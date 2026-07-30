#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Цены по артикулам (наши + конкуренты) → лист «Цены» той же Google-таблицы.

Задача Егора от 30.07.2026: цены нужны «в том же формате, поартикульно», что и
позиции в полках, — чтобы не проходить руками по всем конкурентам несколько раз
в день. Источник строк тот же, что у позиций: товарные группы из таблицы
«Аналитика цен WB» (to_sheets.read_groups), но цена снимается у ВСЕХ артикулов
группы — и у наших, и у конкурентов.

Откуда цена: публичный `card.wb.ru/cards/v4/detail` (до 100 артикулов за запрос,
без авторизации). `sizes[].price.product / 100` — цена карточки; в таблицу идёт
цена с WB-кошельком = floor(цена × 0.98) — ровно та колонка, которую менеджеры
ведут в «Аналитике цен (wb кошелек)» (сверка с MPStats 28.07.2026: именно floor,
не round — 559×0.98=547.8 → 547).

Раскладка листа «Цены» — как у «Матрицы»:
    A..D — Товар | Артикул | Бренд | Чей («наш»/«конкурент»), внутри группы
           сначала наши строки (подсвечены голубым), потом конкуренты
    строка 1 — дата над колонкой (в A1 — подпись про кошелёк)
    строка 2 — время последнего замера этой даты
    новая дата врезается ОДНОЙ колонкой СЛЕВА, прежние съезжают вправо;
    повторный прогон в тот же день обновляет колонку НА МЕСТЕ (меняется только
    время в строке 2) — интрадей-прогоны историю не раздувают.

Раскраска: сравнение со ВЧЕРАШНЕЙ колонкой — подешевело зелёным, подорожало
красным. Интрадей-перезапись тоже сравнивается со вчера, а не сама с собой,
поэтому «цена вернулась» гаснет, а не остаётся висеть.

Три служебных состояния, которые нельзя смешивать (та же логика, что у полок):
«нет в наличии» (карточка есть, цены нет), «нет карточки» (артикул удалён с WB),
«ошибка сбора» (батч не отдался по сети — дырка замера, НЕ факт о товаре).

Запуск: ежедневно — шагом в shelves.yml после позиций; интрадей — отдельный
лёгкий workflow prices.yml (workflow_dispatch), его дергает cron-job.org.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import time
from datetime import datetime

import gspread

import shelf_positions as sp
import to_sheets as ts

SHEET_PRICES = "Цены"
PRICE_FIX = ["Товар", "Артикул", "Бренд", "Чей"]
HEAD_ROWS = 2            # строка 1 — даты, строка 2 — шапка/время замера
NOTE_A1 = "Цена с WB-кошельком, ₽"

COLOR_DOWN = {"red": 0.78, "green": 0.91, "blue": 0.79}   # подешевело со вчера
COLOR_UP = {"red": 0.96, "green": 0.80, "blue": 0.80}     # подорожало со вчера
COLOR_STATE = {"red": 0.93, "green": 0.93, "blue": 0.93}  # нет в наличии / нет карточки
COLOR_OURS = {"red": 0.85, "green": 0.91, "blue": 0.98}   # наши строки, фикс-колонки
WHITE = {"red": 1, "green": 1, "blue": 1}

STATE_NONE = "нет в наличии"
STATE_GONE = "нет карточки"
STATE_FAIL = "ошибка сбора"


def wallet_price(kopecks: int) -> int:
    # Именно floor, не round: сверено с MPStats на 10 парах день/артикул.
    return math.floor(kopecks / 100 * 0.98)


def fetch_prices(nm_ids: list[int], dest: int) -> dict[int, int | str]:
    """{nmID: цена с кошельком | STATE_*} по всем артикулам, батчами по 100."""
    session = sp._session()
    out: dict[int, int | str] = {}
    for i in range(0, len(nm_ids), 100):
        chunk = nm_ids[i:i + 100]
        params = {"appType": "1", "curr": "rub", "spp": "30", "dest": str(dest),
                  "nm": ";".join(str(x) for x in chunk)}
        data = sp._get_json(session, sp.CARD_URL, params)
        if data is None or data is sp.EMPTY:
            # Сетевой сбой всего батча — это НЕ «цены нет», это дырка замера.
            for nm in chunk:
                out[nm] = STATE_FAIL
            continue
        got: dict[int, int | None] = {}
        for p in (data.get("products") or []):
            kop = None
            for size in p.get("sizes") or []:
                pr = (size.get("price") or {}).get("product")
                if pr:
                    kop = pr
                    break
            got[p["id"]] = kop
        for nm in chunk:
            if nm not in got:
                out[nm] = STATE_GONE      # карточка удалена/скрыта на WB
            elif not got[nm]:
                out[nm] = STATE_NONE      # карточка есть, товара в продаже нет
            else:
                out[nm] = wallet_price(got[nm])
        time.sleep(0.2)
    return out


def build_rows(groups: list[dict]) -> list[dict]:
    """Строки листа: группа за группой, внутри группы сначала наши артикулы."""
    rows = []
    for g in groups:
        ours = set(g["ours"])
        for nm in list(g["ours"]) + list(g["competitors"]):
            rows.append({
                "product": g["product"],
                "nm": str(nm),
                "brand": (g.get("brand_of") or {}).get(nm, ""),
                "ours": nm in ours,
            })
    return rows


def _as_int(text: str) -> int | None:
    digits = re.sub(r"[^\d]", "", text or "")
    return int(digits) if digits else None


def write_prices(book, rows: list[dict], prices: dict[int, int | str],
                 when: datetime) -> tuple[str, str]:
    """Колонка цен за дату `when`; вернёт (дата, время) записанного замера."""
    nfix = len(PRICE_FIX)
    date = when.strftime("%d.%m")
    tm = when.strftime("%H:%M")

    ws = ts.ensure_ws(book, SHEET_PRICES, rows=len(rows) + HEAD_ROWS + 50,
                      cols=nfix + 40)
    values = ws.get_all_values()

    fresh = (len(values) < HEAD_ROWS or not values[1]
             or values[1][0].strip() != PRICE_FIX[0])
    if fresh:
        ws.clear()
        ws.update(values=[[NOTE_A1], list(PRICE_FIX)], range_name="A1")
        values = [[NOTE_A1] + [""] * (nfix - 1), list(PRICE_FIX)]

    row1 = values[0] if values else []
    body = values[HEAD_ROWS:]

    # Свежая дата всегда в первой колонке за фиксированными: если сегодня уже
    # писали — она там и стоит, иначе врезаем новую колонку на это же место.
    date_is_new = not (len(row1) > nfix and row1[nfix].strip() == date)

    # Вчерашняя колонка для раскраски — индекс в СТАРОЙ раскладке (до вставки).
    prev_idx = nfix if date_is_new else nfix + 1
    prev: dict[tuple[str, str], int | None] = {}
    row_of: dict[tuple[str, str], int] = {}
    for i, r in enumerate(body):
        key = ((r[0].strip() if len(r) > 0 else ""),
               (r[1].strip() if len(r) > 1 else ""))
        if not key[1]:
            continue
        row_of.setdefault(key, i + HEAD_ROWS + 1)
        if len(r) > prev_idx:
            prev[key] = _as_int(r[prev_idx])

    new_fixed = []
    next_row = len(body) + HEAD_ROWS + 1
    for e in rows:
        key = (e["product"], e["nm"])
        if key not in row_of:
            row_of[key] = next_row
            new_fixed.append(e)
            next_row += 1

    # Только строки: ширину добирает вставка колонки. Резать по колонкам нельзя —
    # gspread держит устаревший col_count и снёс бы правый край с прежними датами.
    if ws.row_count < next_row + 30:
        ws.resize(rows=next_row + 30, cols=ws.col_count)

    if date_is_new:
        book.batch_update({"requests": [{"insertDimension": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": nfix, "endIndex": nfix + 1},
            "inheritFromBefore": True}}]})

    date_col = nfix + 1  # 1-базный номер колонки текущей даты

    updates = [
        {"range": gspread.utils.rowcol_to_a1(1, date_col), "values": [[date]]},
        {"range": gspread.utils.rowcol_to_a1(2, date_col), "values": [[tm]]},
    ]
    if new_fixed:
        first_new = row_of[(new_fixed[0]["product"], new_fixed[0]["nm"])]
        updates.append({
            "range": f"{gspread.utils.rowcol_to_a1(first_new, 1)}:"
                     f"{gspread.utils.rowcol_to_a1(first_new + len(new_fixed) - 1, nfix)}",
            "values": [[e["product"], e["nm"], e["brand"],
                        "наш" if e["ours"] else "конкурент"] for e in new_fixed]})

    used = sorted(row_of[(e["product"], e["nm"])] for e in rows)
    first, last = used[0], used[-1]
    column = [[""] for _ in range(last - first + 1)]
    for e in rows:
        column[row_of[(e["product"], e["nm"])] - first] = \
            [prices.get(int(e["nm"]), STATE_FAIL)]
    updates.append({
        "range": f"{gspread.utils.rowcol_to_a1(first, date_col)}:"
                 f"{gspread.utils.rowcol_to_a1(last, date_col)}",
        "values": column})

    ws.batch_update(updates, value_input_option="USER_ENTERED")

    # Оформление: заморозка, шапка, сброс колонки в белый + раскраска изменений.
    reqs = [
        {"updateSheetProperties": {
            "properties": {"sheetId": ws.id,
                           "gridProperties": {"frozenRowCount": HEAD_ROWS,
                                              "frozenColumnCount": nfix}},
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}},
        {"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": HEAD_ROWS},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True},
                                           "wrapStrategy": "WRAP",
                                           "horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat(textFormat,wrapStrategy,horizontalAlignment)"}},
        # Всю колонку даты — в белый (вставка наследует голубую заливку наших
        # строк из «Чей», а при перезаписи того же дня гасим вчерашнюю раскраску
        # у ячеек, где цена вернулась к вчерашней).
        {"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": HEAD_ROWS,
                      "endRowIndex": max(next_row, last),
                      "startColumnIndex": nfix, "endColumnIndex": nfix + 1},
            "cell": {"userEnteredFormat": {
                "backgroundColor": WHITE,
                "horizontalAlignment": "CENTER",
                "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,numberFormat)"}},
    ]

    # Наши строки подсвечиваем в фиксированных колонках — один раз, при появлении.
    for e in new_fixed:
        if not e["ours"]:
            continue
        r = row_of[(e["product"], e["nm"])]
        reqs.append({"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": r - 1, "endRowIndex": r,
                      "startColumnIndex": 0, "endColumnIndex": nfix},
            "cell": {"userEnteredFormat": {"backgroundColor": COLOR_OURS}},
            "fields": "userEnteredFormat(backgroundColor)"}})

    changed_down = changed_up = 0
    for e in rows:
        val = prices.get(int(e["nm"]), STATE_FAIL)
        if isinstance(val, int):
            old = prev.get((e["product"], e["nm"]))
            if old is None or old == val:
                continue
            color = COLOR_DOWN if val < old else COLOR_UP
            if val < old:
                changed_down += 1
            else:
                changed_up += 1
        elif val == STATE_FAIL:
            continue                      # дырку замера не красим
        else:
            color = COLOR_STATE
        r = row_of[(e["product"], e["nm"])]
        reqs.append({"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": r - 1, "endRowIndex": r,
                      "startColumnIndex": nfix, "endColumnIndex": nfix + 1},
            "cell": {"userEnteredFormat": {"backgroundColor": color}},
            "fields": "userEnteredFormat(backgroundColor)"}})

    for i in range(0, len(reqs), 500):
        book.batch_update({"requests": reqs[i:i + 500]})

    sp.log(f"Изменений со вчера: подешевело {changed_down}, подорожало {changed_up}.")
    return date, tm


# ------------------------------------------------------------------------- main

def main() -> None:
    ap = argparse.ArgumentParser(description="Цены по артикулам групп → лист «Цены»")
    ap.add_argument("--sheet-id", default=os.environ.get("SHEET_ID"))
    ap.add_argument("--source-id", default=os.environ.get("SOURCE_ID"))
    ap.add_argument("--source-tab", default=ts.SOURCE_TAB)
    ap.add_argument("--creds", default=None)
    ap.add_argument("--dest", type=int, default=sp.DEST_MOSCOW)
    ap.add_argument("--save-snapshot", default="snapshot_prices.json")
    ap.add_argument("--dry-run", action="store_true",
                    help="снять цены и сохранить снапшот, в таблицу не писать")
    args = ap.parse_args()

    if not args.sheet_id:
        raise SystemExit("Не задан --sheet-id (или SHEET_ID)")
    if not args.source_id:
        raise SystemExit("Не задан --source-id (или SOURCE_ID) — откуда брать артикулы")

    client = ts.get_client(args.creds)
    groups, skipped = ts.read_groups(client, args.source_id, args.source_tab)
    rows = build_rows(groups)
    nm_ids = sorted({int(e["nm"]) for e in rows})
    sp.log(f"Групп: {len(groups)}; артикулов к съёму цены: {len(nm_ids)} "
           f"(строк в листе: {len(rows)}); строк источника пропущено: {len(skipped)}")

    prices = fetch_prices(nm_ids, args.dest)
    ok = sum(1 for v in prices.values() if isinstance(v, int))
    none = sum(1 for v in prices.values() if v == STATE_NONE)
    gone = sum(1 for v in prices.values() if v == STATE_GONE)
    fail = sum(1 for v in prices.values() if v == STATE_FAIL)
    sp.log(f"Цен получено: {ok}; нет в наличии: {none}; нет карточки: {gone}; "
           f"ошибок сбора: {fail}")

    when = datetime.now(sp.MSK)
    if args.save_snapshot:
        with open(args.save_snapshot, "w", encoding="utf-8") as f:
            json.dump({"snapshot_at": when.isoformat(timespec="seconds"),
                       "dest": args.dest,
                       "source": "card.wb.ru/cards/v4/detail (wb кошелёк, floor ×0.98)",
                       "prices": {str(k): v for k, v in prices.items()}},
                      f, ensure_ascii=False, indent=2)
        sp.log(f"Снапшот сохранён: {args.save_snapshot}")

    if args.dry_run:
        print("dry-run: в таблицу не пишу.")
        return

    book = client.open_by_key(args.sheet_id)
    date, tm = write_prices(book, rows, prices, when)
    print(f"Готово: лист «{SHEET_PRICES}», колонка за {date} обновлена в {tm} "
          f"({ok} цен, строк {len(rows)}).")


if __name__ == "__main__":
    main()
