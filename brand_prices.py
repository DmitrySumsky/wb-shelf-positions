#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Цены по карточкам листа «Цены» книги бренда.

Задача от 07.08.2026: в книге бренда (Health Form) уже лежит лист «Цены» —
копия строк «Полок» (наши карточки и конкуренты, группа за группой), но без
истории цен. Нужен тот же сбор, что в общей книге «Анализ конкурентов ВБ авто»
(`prices_to_sheets.py`), только источник строк — сам лист книги.

Отличие от `prices_to_sheets.py` (там лист собирается скриптом из внешней
таблицы-источника): здесь строки ведёт человек, скрипт их НЕ трогает вообще —
ни значений, ни порядка, ни раскраски. Он читает артикулы и дописывает колонку
за сегодня. Ровно тот же договор, что у `brand_shelves.py` с листом «Полки».

Раскладка листа:
    A..  фиксированные колонки человека (Товар | Артикул конкурента | Бренд |
         ~ Выручка | Тип) — границей служит первая колонка-дата, как в «Полках»
    строка 1 — шапка; над колонкой даты стоит дата ДД.ММ, время замера —
         примечанием к этой ячейке (вторую строку шапки не заводим: лист уже
         ведётся с данными со строки 2, сдвиг сломал бы ссылки человека)
    свежая дата врезается ОДНОЙ колонкой СЛЕВА, прежние съезжают вправо;
    повторный прогон в тот же день переписывает колонку НА МЕСТЕ.

Цена — с WB-кошельком: floor(цена × 0.98), та же формула и тот же источник
(`card.wb.ru/cards/v4/detail`), что в `prices_to_sheets.py`.

Что в ячейке:
    число            — цена с кошельком;
    «нет в наличии»  — карточка есть, товара в продаже нет;
    «нет карточки»   — артикул удалён/скрыт на WB;
    «ошибка сбора»   — батч не отдался по сети (дырка замера, НЕ факт о товаре);
    «нет такого товара» — в строке нет артикула (у бренда нет карточки), снимать
                       нечего; так же помечает пустоту `brand_shelves.py`.

Раскраска колонки: подешевело со вчера — зелёным, подорожало — красным,
текстовые состояния — серым. Сравнение всегда со ВЧЕРАШНЕЙ колонкой, поэтому
повторный прогон в тот же день не красит цену «сама с собой».

Запуск: шагом в shelves.yml после полок либо руками:
    python brand_prices.py --brand "Health Form" --creds ключ.json
"""

from __future__ import annotations

import argparse
import json
import os
from datetime import datetime

import gspread

import brand_shelves as bs
import prices_to_sheets as pts
import shelf_positions as sp
import to_sheets as ts
from vexor_shelves import install_retries

SHEET_PRICES = "Цены"
KEEP_DATES = 90

# Книги, где сбор цен включён. Лист «Цены» ведёт человек, поэтому книга
# добавляется сюда осознанно, а не «раз есть лист — пишем».
BRANDS_WITH_PRICES = ["Health Form"]

STATE_NO_PRODUCT = bs.STATE_NO_PRODUCT
TEXT_STATES = (pts.STATE_NONE, pts.STATE_GONE, pts.STATE_FAIL, STATE_NO_PRODUCT)


def read_price_rows(ws) -> tuple[list[dict], dict, list[list[str]]]:
    """Лист → (строки, раскладка, сырые значения).

    Строка: {"row": номер в листе, "nm": артикул или ""}. Раскладка ищется той
    же функцией, что у «Полок»: границей фиксированной части служит первая
    колонка-дата, колонки — по названиям, а не по буквам.
    """
    values = ws.get_all_values()
    if not values:
        raise SystemExit(f"Лист «{SHEET_PRICES}» пуст")
    head = [str(x).strip() for x in values[0]]
    lay = bs.read_layout(head)
    ia = lay["art"]
    width = max(lay["nfix"], ia) + 1

    rows: list[dict] = []
    for i, raw in enumerate(values[1:], start=2):
        cells = [str(x).strip() for x in (list(raw) + [""] * width)[:width]]
        if not any(cells[:lay["nfix"]]):
            continue                      # разделитель между группами
        art = cells[ia]
        rows.append({"row": i, "nm": art if art.isdigit() else ""})
    return rows, lay, values


def prev_prices(values: list[list[str]], rows: list[dict], col: int) -> dict[int, int]:
    """{номер строки: вчерашняя цена} по колонке `col` (0-based) старой раскладки."""
    out: dict[int, int] = {}
    for r in rows:
        raw = values[r["row"] - 1]
        if len(raw) > col:
            val = pts._as_int(str(raw[col]))
            if val is not None:
                out[r["row"]] = val
    return out


def write_column(book, ws, lay: dict, values: list[list[str]], rows: list[dict],
                 prices: dict[int, int | str], when: datetime) -> tuple[str, int, int]:
    """Колонка цен за сегодня. Вернёт (дата, подешевело, подорожало)."""
    today = when.strftime("%d.%m")
    nfix = lay["nfix"]
    nrows = len(values)

    if today in lay["date_at"]:
        col = lay["date_at"][today] + 1              # повторный прогон в тот же день
        ndates = len(lay["date_at"])
        prev_col = nfix + 1                          # вчера — колонка правее сегодняшней
    else:
        col = nfix + 1
        if ws.col_count < nfix + len(lay["date_at"]) + 1:
            ws.resize(rows=ws.row_count, cols=nfix + len(lay["date_at"]) + 4)
        book.batch_update({"requests": [{"insertDimension": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": nfix, "endIndex": nfix + 1},
            "inheritFromBefore": True}}]})
        ndates = len(lay["date_at"]) + 1
        prev_col = nfix                              # индексы ДО вставки колонки
    prev = prev_prices(values, rows, prev_col)

    cells: dict[int, object] = {}
    for r in rows:
        cells[r["row"]] = (prices.get(int(r["nm"]), pts.STATE_FAIL) if r["nm"]
                           else STATE_NO_PRODUCT)

    letter = bs.col_letter(col)
    column = [[cells.get(i, "")] for i in range(2, nrows + 1)]
    ws.update(values=[[today]], range_name=f"{letter}1", value_input_option="RAW")
    ws.update(values=column, range_name=f"{letter}2:{letter}{nrows}",
              value_input_option="USER_ENTERED")

    reqs: list[dict] = []
    if ndates > KEEP_DATES:
        reqs.append({"deleteDimension": {"range": {
            "sheetId": ws.id, "dimension": "COLUMNS",
            "startIndex": nfix + KEEP_DATES, "endIndex": nfix + ndates}}})
        ndates = KEEP_DATES
    last_col = nfix + ndates

    reqs += [
        {"updateSheetProperties": {
            "properties": {"sheetId": ws.id,
                           "gridProperties": {"frozenRowCount": 1,
                                              "frozenColumnCount": 3}},
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}},
        # Время замера — примечанием к ячейке даты: вторая строка шапки сдвинула
        # бы все строки листа, который ведёт человек.
        {"updateCells": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": col - 1, "endColumnIndex": col},
            "rows": [{"values": [{
                "note": f"Замер {when.strftime('%d.%m.%Y %H:%M')} МСК. "
                        f"Цена с WB-кошельком, ₽ (floor ×0,98)",
                "userEnteredFormat": {"textFormat": {"bold": True},
                                      "horizontalAlignment": "CENTER"}}]}],
            "fields": "note,userEnteredFormat(textFormat,horizontalAlignment)"}},
        # Колонка целиком в белый: вставка наследует оформление соседа слева,
        # а перезапись того же дня должна гасить вчерашнюю раскраску.
        {"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": nrows,
                      "startColumnIndex": col - 1, "endColumnIndex": col},
            "cell": {"userEnteredFormat": {
                "backgroundColor": pts.WHITE,
                "horizontalAlignment": "CENTER",
                "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,numberFormat)"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": nfix, "endIndex": last_col},
            "properties": {"pixelSize": 75}, "fields": "pixelSize"}},
    ]

    down = up = 0
    for r in rows:
        val = cells[r["row"]]
        if isinstance(val, int):
            old = prev.get(r["row"])
            if old is None or old == val:
                continue
            color = pts.COLOR_DOWN if val < old else pts.COLOR_UP
            if val < old:
                down += 1
            else:
                up += 1
        elif val == pts.STATE_FAIL:
            continue                      # дырку замера не красим
        else:
            color = pts.COLOR_STATE
        reqs.append({"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": r["row"] - 1,
                      "endRowIndex": r["row"],
                      "startColumnIndex": col - 1, "endColumnIndex": col},
            "cell": {"userEnteredFormat": {"backgroundColor": color}},
            "fields": "userEnteredFormat.backgroundColor"}})

    for i in range(0, len(reqs), 300):
        book.batch_update({"requests": reqs[i:i + 300]})
    return today, down, up


def run_brand(client, brand: str, args) -> str:
    cfg = bs.BRANDS[brand]
    sp.log(f"=== {brand}: цены ===")
    book = client.open_by_key(os.environ.get(f"SHEET_ID_{brand.upper().replace(' ', '_')}")
                              or cfg["sheet_id"])
    try:
        ws = book.worksheet(SHEET_PRICES)
    except gspread.WorksheetNotFound:
        raise SystemExit(f"В книге {brand} нет листа «{SHEET_PRICES}» — "
                         "строки листа ведёт человек, сам его не завожу")

    rows, lay, values = read_price_rows(ws)
    nm_ids = sorted({int(r["nm"]) for r in rows if r["nm"]})
    sp.log(f"Лист «{SHEET_PRICES}»: строк {len(rows)}, из них без артикула "
           f"{sum(1 for r in rows if not r['nm'])}; уникальных карточек "
           f"{len(nm_ids)}; дат в истории {len(lay['date_at'])}")

    prices = pts.fetch_prices(nm_ids, args.dest)
    ok = sum(1 for v in prices.values() if isinstance(v, int))
    stat = (f"цен {ok}, нет в наличии "
            f"{sum(1 for v in prices.values() if v == pts.STATE_NONE)}, "
            f"нет карточки {sum(1 for v in prices.values() if v == pts.STATE_GONE)}, "
            f"ошибок сбора {sum(1 for v in prices.values() if v == pts.STATE_FAIL)}")
    sp.log(stat)

    if args.save_snapshots:
        path = f"snapshot_prices_{brand.lower().replace(' ', '_')}.json"
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"snapshot_at": datetime.now(sp.MSK).isoformat(timespec="seconds"),
                       "brand": brand, "dest": args.dest,
                       "source": "card.wb.ru/cards/v4/detail (wb кошелёк, floor ×0.98)",
                       "prices": {str(k): v for k, v in prices.items()}},
                      f, ensure_ascii=False, indent=2)
        sp.log(f"Снапшот сохранён: {path}")

    if args.dry_run:
        return f"{brand}: dry-run, {stat}"

    # Строки человек правит когда угодно, а колонка пишется по их номерам —
    # перечитываем лист прямо перед записью (тот же приём, что в brand_shelves).
    rows, lay, values = read_price_rows(ws)
    date, down, up = write_column(book, ws, lay, values, rows, prices,
                                  datetime.now(sp.MSK))
    return (f"{brand}: колонка за {date} записана ({stat}; "
            f"подешевело {down}, подорожало {up})")


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Цены по карточкам листа «Цены» книги бренда")
    ap.add_argument("--brand", default="all",
                    help="имя бренда или all — все книги со сбором цен "
                         f"({', '.join(BRANDS_WITH_PRICES)})")
    ap.add_argument("--creds", default=None)
    ap.add_argument("--dest", type=int, default=sp.DEST_MOSCOW)
    ap.add_argument("--save-snapshots", action="store_true", default=True)
    ap.add_argument("--dry-run", action="store_true", help="в таблицу не писать")
    args = ap.parse_args()

    if args.brand.lower() == "all":
        brands = list(BRANDS_WITH_PRICES)
    else:
        match = [b for b in BRANDS_WITH_PRICES if b.lower() == args.brand.strip().lower()]
        if not match:
            raise SystemExit(f"Сбор цен включён только для: {BRANDS_WITH_PRICES}; "
                             f"пришло {args.brand!r}")
        brands = match

    install_retries()
    client = ts.get_client(args.creds)

    results, failed = [], []
    for brand in brands:
        try:
            results.append(run_brand(client, brand, args))
        except Exception as exc:
            sp.log(f"ОШИБКА по бренду {brand}: {exc.__class__.__name__}: {exc}")
            failed.append(brand)
    for line in results:
        print("Готово:", line)
    if failed:
        raise SystemExit(f"Бренды с ошибкой: {', '.join(failed)}")


if __name__ == "__main__":
    main()
