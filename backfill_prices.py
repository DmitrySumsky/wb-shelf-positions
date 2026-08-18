#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Разовая доливка ИСТОРИИ цен в лист «Цены» (задача 30.07.2026).

Зачем: лист «Цены» заведён 30.07 и до этой доливки содержал ровно один день —
по 203 конкурентам истории не было вообще, сравнивать «дороже/дешевле стало»
было не с чем.

Откуда история: НЕ из MPStats напрямую, а из таблицы-источника «Аналитика цен
WB», лист «Аналитика цен (wb кошелек)» — той самой, откуда `to_sheets.read_groups`
берёт список артикулов для полок и цен. Там менеджеры уже ведут ту же самую
метрику, и её ежедневно дозаливает `21.orders-cloud/.../prices_update.py`.
Проверено перед запуском (30.07.2026):
  * значения источника совпадают с `wallet_price` MPStats один в один (6 проб);
  * все 322 артикула листа «Цены» есть в источнике (покрытие 100%);
  * медиана «Цены»(30.07) / источник(29.07) по 304 парам = 1.0000 — метрика и
    масштаб те же, никакого пересчёта не нужно.
Это заодно даёт глубину БОЛЬШЕ, чем MPStats: у него окно ровно 30 дней, а в
источнике уже накоплено 59 (с 01.06) — и лишний раз его квоту мы не тратим.

Куда пишет: колонками СПРАВА от существующего блока дат, по убыванию даты
(новые левее — как во всём листе). Дневной прогон `prices_to_sheets.py` врезает
свою колонку слева (`insertDimension` в позицию E) и лист целиком не
перезаписывает, поэтому долитые колонки просто съезжают вправо и живут дальше.

Чего НЕ делает: не красит историю «подешевело/подорожало». Раскраска в дневном
прогоне идёт отдельным `repeatCell` на КАЖДУЮ изменившуюся ячейку — на месяц
это ~9 700 запросов; для архива она не нужна, следят за свежей колонкой.
В строке 2 у исторических колонок вместо времени замера стоит «MPStats» — сразу
видно, что это перенесённый дневной срез, а не наш интрадей-замер.

Идемпотентность: дата, которая в листе уже есть, пропускается; повторный запуск
ничего не дублирует. Доливаются только даты СТАРШЕ самой старой в листе —
свежие колонки остаются делом дневного прогона.

Запуск (разово, локально):
  python backfill_prices.py --sheet-id <ID приёмника> --source-id <ID источника> \
      --creds <sa.json> [--days 30] [--dry-run]
"""

from __future__ import annotations

import argparse
import os
import re
from datetime import date, datetime

import gspread

import prices_to_sheets as pp
import to_sheets as ts

DATE_RE = re.compile(r"^\d{1,2}\.\d{1,2}$")
TIME_NOTE = "MPStats"      # строка 2 исторических колонок: не наш замер по времени


def as_int(text) -> int | None:
    """«1 234 ₽» → 1234. Пусто/мусор → None (пустую ячейку не выдумываем)."""
    digits = re.sub(r"[^\d]", "", str(text))
    return int(digits) if digits else None


def to_date(label: str, today: date) -> date:
    """«ДД.ММ» → дата. Месяц больше текущего = прошлый год (как в prices_update)."""
    d, m = (int(x) for x in label.strip().split("."))
    year = today.year if m <= today.month else today.year - 1
    return date(year, m, d)


def read_source_history(client, source_id: str, source_tab: str):
    """({артикул: {«ДД.ММ»: цена}}, [даты слева направо]) из таблицы-источника."""
    ws = client.open_by_key(source_id).worksheet(source_tab)
    values = ws.get_all_values()
    hdr = values[0] if values else []
    cols = {i: h.strip() for i, h in enumerate(hdr) if DATE_RE.match(h.strip())}
    hist: dict[str, dict[str, int]] = {}
    for row in values[1:]:
        art = row[1].strip() if len(row) > 1 else ""
        # артикул может стоять в НЕСКОЛЬКИХ товарных группах источника — цена у
        # него одна и та же, поэтому берём первую встреченную строку
        if not art.isdigit() or art in hist:
            continue
        hist[art] = {d: v for i, d in cols.items()
                     if len(row) > i and (v := as_int(row[i])) is not None}
    return hist, [cols[i] for i in sorted(cols)]


def main() -> None:
    ap = argparse.ArgumentParser(description="Доливка истории цен в лист «Цены»")
    ap.add_argument("--sheet-id", default=os.environ.get("SHEET_ID"))
    ap.add_argument("--source-id", default=os.environ.get("SOURCE_ID"))
    ap.add_argument("--source-tab", default=ts.SOURCE_TAB)
    ap.add_argument("--creds", default=None)
    ap.add_argument("--days", type=int, default=30,
                    help="сколько дней истории долить (по умолчанию 30)")
    ap.add_argument("--dry-run", action="store_true",
                    help="показать план и сверку, в таблицу не писать")
    args = ap.parse_args()

    if not args.sheet_id:
        raise SystemExit("Не задан --sheet-id (или SHEET_ID)")
    if not args.source_id:
        raise SystemExit("Не задан --source-id (или SOURCE_ID)")

    today = datetime.now(pp.sp.MSK).date()
    client = ts.get_client(args.creds)

    hist, src_dates = read_source_history(client, args.source_id, args.source_tab)
    pp.sp.log(f"Источник: артикулов {len(hist)}, дат {len(src_dates)} "
              f"({src_dates[0]} … {src_dates[-1]})")

    book = client.open_by_key(args.sheet_id)
    ws = book.worksheet(pp.SHEET_PRICES)
    values = ws.get_all_values()
    nfix = len(pp.PRICE_FIX)
    hdr = values[0] if values else []

    have_idx = [i for i, h in enumerate(hdr) if DATE_RE.match(h.strip())]
    if not have_idx:
        raise SystemExit("В листе «Цены» нет ни одной колонки с датой — "
                         "сначала должен пройти дневной prices_to_sheets.py")
    have = {hdr[i].strip() for i in have_idx}
    oldest_have = min(to_date(hdr[i], today) for i in have_idx)
    last_col0 = max(have_idx)

    # только СТАРШЕ самой старой в листе: свежие даты — дело дневного прогона
    todo = [d for d in src_dates
            if d not in have and to_date(d, today) < oldest_have][:args.days]
    if not todo:
        print("Доливать нечего: в листе уже есть все даты источника, которые старше "
              f"{oldest_have:%d.%m}.")
        return

    body = values[pp.HEAD_ROWS:]
    art_of: dict[int, str] = {}
    for i, r in enumerate(body):
        art = r[1].strip() if len(r) > 1 else ""
        if art.isdigit():
            art_of[i + pp.HEAD_ROWS + 1] = art
    if not art_of:
        raise SystemExit("В листе «Цены» не нашлось строк с артикулами")
    first, last = min(art_of), max(art_of)

    grid = []
    for rn in range(first, last + 1):
        h = hist.get(art_of.get(rn, ""), {})
        grid.append([h.get(d, "") for d in todo])

    filled = sum(1 for row in grid for v in row if v != "")
    holes = len(art_of) * len(todo) - filled
    no_hist = sorted({art_of[rn] for rn in art_of if not hist.get(art_of[rn])})

    start_col1 = last_col0 + 2                       # сразу за блоком дат, 1-базно
    end_col1 = start_col1 + len(todo) - 1
    pp.sp.log(f"Доливаю {len(todo)} дат ({todo[0]} … {todo[-1]}) в колонки "
              f"{gspread.utils.rowcol_to_a1(1, start_col1)[:-1]}–"
              f"{gspread.utils.rowcol_to_a1(1, end_col1)[:-1]}, "
              f"строк {len(art_of)} (с {first} по {last})")
    pp.sp.log(f"Заполнится ячеек: {filled}, останется пустыми: {holes}"
              + (f"; артикулов без истории вовсе: {len(no_hist)}" if no_hist else ""))

    if args.dry_run:
        print("\ndry-run: в таблицу не пишу. Примеры первых строк:")
        for rn in range(first, min(first + 5, last + 1)):
            print(f"  {art_of.get(rn, '—'):>12}: "
                  + ", ".join(f"{d}={v or '—'}" for d, v in zip(todo[:6], grid[rn - first])))
        return

    need_cols = max(end_col1, ws.col_count)
    if ws.col_count < end_col1:
        # явным числом через updateSheetProperties, а не ws.resize(): gspread
        # держит col_count с момента открытия листа и после чужой вставки
        # колонок способен урезать правый край (грабля 28.07.2026)
        book.batch_update({"requests": [{"updateSheetProperties": {
            "properties": {"sheetId": ws.id,
                           "gridProperties": {"columnCount": need_cols}},
            "fields": "gridProperties.columnCount"}}]})
        pp.sp.log(f"Колонок в листе: расширено до {need_cols}")

    a1 = gspread.utils.rowcol_to_a1
    ws.batch_update([
        {"range": f"{a1(1, start_col1)}:{a1(1, end_col1)}", "values": [todo]},
        {"range": f"{a1(2, start_col1)}:{a1(2, end_col1)}",
         "values": [[TIME_NOTE] * len(todo)]},
        {"range": f"{a1(first, start_col1)}:{a1(last, end_col1)}", "values": grid},
    ], value_input_option="USER_ENTERED")

    book.batch_update({"requests": [
        {"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": pp.HEAD_ROWS,
                      "startColumnIndex": start_col1 - 1, "endColumnIndex": end_col1},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True},
                                           "wrapStrategy": "WRAP",
                                           "horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat(textFormat,wrapStrategy,horizontalAlignment)"}},
        {"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": pp.HEAD_ROWS,
                      "endRowIndex": last, "startColumnIndex": start_col1 - 1,
                      "endColumnIndex": end_col1},
            "cell": {"userEnteredFormat": {
                "backgroundColor": pp.WHITE,
                "horizontalAlignment": "CENTER",
                "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,numberFormat)"}},
    ]})

    print(f"Готово: долито {len(todo)} дат ({todo[-1]} … {todo[0]}), "
          f"ячеек с ценой {filled}, пустых {holes}. "
          f"История без раскраски — следят за свежей колонкой.")
    if no_hist:
        print(f"Без истории в источнике ({len(no_hist)} арт.): "
              + ", ".join(no_hist[:12]) + (" …" if len(no_hist) > 12 else ""))


if __name__ == "__main__":
    main()
