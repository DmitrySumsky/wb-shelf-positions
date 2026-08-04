#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Отдельная книга бренда: позиции карточек NATURI в полках конкурентов.

Задача Артура от 04.08.2026: разнести бренды по отдельным таблицам. Образец —
лист «Натури пример» в общей книге «Анализ конкурентов ВБ авто»: плоский список,
где строка = полка (карточка конкурента) внутри группы «наш товар», а в колонках
дат стоит позиция НАШЕЙ карточки в этой полке, свежая дата слева.

Раскладка листа «Полки» книги-приёмника:
    A  Товар                 \\
    B  Артикул конкурента     |
    C  Бренд конкурента       |  ведёт человек (Артур): строки, порядок,
    D  ~ Выручка конкурента   |  «Тип конкурента», «Прогрев» — только руками
    E  Тип конкурента         |
    F  Прогрев                |
    G  Прогрев к-во          /
    H.. даты ДД.ММ, свежая СЛЕВА

Единственный источник правды по строкам — САМ лист приёмника: скрипт читает
A–G, дописывает колонку за сегодня и больше ничего в них не трогает (ни значений,
ни оформления). Первый запуск, когда листа ещё нет, копирует его из «Натури
пример» вместе с раскраской и чистит пустые колонки-даты образца.

Группа = подряд идущие строки с одинаковым «Товар». Наша карточка группы —
ПЕРВАЯ строка группы с нашим брендом в колонке C (в образце она выделена тёмно-
зелёным). Остальные строки группы, включая наши же другие карточки, считаются
полками: в них ищется карточка группы.

Что в ячейке даты:
    строка полки        — позиция нашей карточки (число),
                          «—» = полку проверили, нас там нет,
                          «нет карточки» = артикул полки удалён/скрыт на WB,
                          «ошибка сбора» = полка не отдалась (это НЕ «нас там нет»);
    строка нашей карточки — «N из M»: в скольких полках группы нашлись.

Запуск: шагом в shelves.yml (08:00 МСК) либо руками:
    python naturi_shelves.py --creds ключ.json
"""

from __future__ import annotations

import argparse
import json
import os
import re
from datetime import datetime

import gspread

import shelf_positions as sp
import to_sheets as ts
from vexor_shelves import install_retries

# Книга «Анализ конкурентов Naturi». Не секрет: ID есть в памяти проекта.
NATURI_SHEET_ID = "1XLby8VEOKQtuXrm4OCiQ-PFiTMXfaUe7gF054feoh_0"

# Откуда взять раскладку при самом первом запуске (книга «Анализ конкурентов ВБ авто»).
TEMPLATE_SHEET_ID = "1hqCt4QnCnqrLrRUZD3hSCuDd3k-2PFpaZdHNaxE4Nzk"
TEMPLATE_TAB = "Натури пример"

SHEET_DST = "Полки"
OUR_BRAND = "NATURI"

FIX_COLS = ["Товар", "Артикул конкурента", "Бренд конкурента", "~ Выручка конкурента",
            "Тип конкурента", "Прогрев", "Прогрев к-во"]
NFIX = len(FIX_COLS)              # A..G
KEEP_DATES = 90

NOT_IN_SHELF = "—"
STATE_GONE = "нет карточки"
STATE_FAIL = "ошибка сбора"

# Раскраска позиций — условным форматированием (шесть правил на любой размер листа),
# как в листе «Полки» книги VEXOR: ячеек тут десятки тысяч, красить их поштучно дорого.
BANDS = [
    (10, {"red": 0.72, "green": 0.88, "blue": 0.72}),
    (30, {"red": 0.85, "green": 0.94, "blue": 0.78}),
    (100, {"red": 1.00, "green": 0.95, "blue": 0.75}),
    (300, {"red": 0.99, "green": 0.87, "blue": 0.75}),
]
COLOR_DEEP = {"red": 0.96, "green": 0.80, "blue": 0.80}
COLOR_NONE = {"red": 0.93, "green": 0.93, "blue": 0.93}


def col_letter(i1: int) -> str:
    s, i = "", i1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def norm_date(text: str) -> str | None:
    """«04.08» / «04.08.2026» → «04.08». Не дата — None.

    Google на USER_ENTERED превращает «04.08» в настоящую дату и показывает с годом,
    поэтому шапку читаем терпимо к формату, а пишем всегда RAW.
    """
    m = re.fullmatch(r"(\d{1,2})\.(\d{1,2})(?:\.\d{2,4})?", str(text).strip())
    return f"{int(m.group(1)):02d}.{int(m.group(2)):02d}" if m else None


# --------------------------------------------------------- лист: завести / прочитать

def bootstrap(client, book, template_id: str, template_tab: str):
    """Первый запуск: копируем раскладку образца вместе с оформлением.

    Копия делается средствами Google (`copyTo`), а не переписыванием значений, —
    так переезжают ширины колонок, заливка строк наших карточек и жёлтая колонка
    «Тип конкурента». Пустые колонки-даты образца сносим: свои даты скрипт
    заведёт сам, а чужие пустые только сбивали бы счёт истории.
    """
    src = client.open_by_key(template_id).worksheet(template_tab)
    sp.log(f"Листа «{SHEET_DST}» в книге нет — копирую раскладку из «{template_tab}»")
    info = src.copy_to(book.id)
    ws = book.get_worksheet_by_id(info["sheetId"])
    ws.update_title(SHEET_DST)
    ws.update_index(0)
    if ws.col_count > NFIX:
        # Заморозку снимаем ДО удаления: у образца заморожены как раз A–G, а
        # Google отвечает «not possible to delete all non-frozen columns», если
        # после удаления остались бы одни замороженные. Ставим её обратно ниже.
        book.batch_update({"requests": [
            {"updateSheetProperties": {
                "properties": {"sheetId": ws.id,
                               "gridProperties": {"frozenColumnCount": 0}},
                "fields": "gridProperties.frozenColumnCount"}},
            {"deleteDimension": {"range": {
                "sheetId": ws.id, "dimension": "COLUMNS",
                "startIndex": NFIX, "endIndex": ws.col_count}}},
        ]})
    return ws


def read_sheet(ws) -> tuple[list[dict], list[str], list[list[str]]]:
    """Лист → (строки, даты слева направо, сырые значения).

    Строка: {"row", "kind": "our"|"shelf"|"skip", "art", "block"}.
    """
    values = ws.get_all_values()
    if not values:
        raise SystemExit(f"Лист «{SHEET_DST}» пуст")
    head = [str(x).strip() for x in values[0]]
    if head[:2] != FIX_COLS[:2]:
        raise SystemExit(f"Шапка «{SHEET_DST}» не та, что ожидалась: {head[:NFIX]}")

    dates: list[str] = []
    for i in range(NFIX, len(head)):
        d = norm_date(head[i])
        if d and d not in dates:
            dates.append(d)

    rows: list[dict] = []
    block = None
    block_has_our = set()
    for i, raw in enumerate(values[1:], start=2):
        cells = [str(x).strip() for x in (list(raw) + [""] * NFIX)[:NFIX]]
        product, art, brand = cells[0], cells[1], cells[2]
        if not product:
            rows.append({"row": i, "kind": "skip", "art": "", "block": None})
            continue
        if product != block:
            block = product
        is_our = brand.upper() == OUR_BRAND.upper()
        # Наша карточка группы — первая наша строка блока. Остальные наши карточки
        # (в образце светло-зелёные) считаются полками: в них тоже интересно стоять.
        kind = "shelf"
        if is_our and block not in block_has_our:
            kind = "our"
            block_has_our.add(block)
        if not art.isdigit():
            kind = "skip"
        rows.append({"row": i, "kind": kind, "art": art, "block": block})
    return rows, dates, values


def build_groups(rows: list[dict]) -> list[dict]:
    """Строки листа → группы для sp.run_groups."""
    order: list[str] = []
    acc: dict[str, dict] = {}
    for r in rows:
        if r["kind"] == "skip" or not r["block"]:
            continue
        g = acc.get(r["block"])
        if g is None:
            g = acc[r["block"]] = {"product": r["block"], "ours": [], "competitors": []}
            order.append(r["block"])
        nm = int(r["art"])
        if r["kind"] == "our":
            g["ours"].append(nm)
        elif nm not in g["competitors"]:      # один артикул в блоке может повторяться
            g["competitors"].append(nm)

    good, skipped = [], []
    for name in order:
        g = acc[name]
        (good if g["ours"] and g["competitors"] else skipped).append(g)
    if skipped:
        sp.log("Группы без нашей карточки или без полок пропущены: "
               + ", ".join(f"{g['product']} (наших {len(g['ours'])}, "
                           f"полок {len(g['competitors'])})" for g in skipped))
    return good


# --------------------------------------------------------------- значения за сегодня

def cell_values(snapshot: dict, groups: list[dict], rows: list[dict]) -> dict[int, object]:
    """{номер строки листа: значение за сегодня}."""
    failed = set(snapshot.get("failed_shelves", []))
    missing = set(snapshot.get("missing_shelves", []))
    positions = snapshot.get("positions", {})
    our_of = {g["product"]: str(g["ours"][0]) for g in groups}

    out: dict[int, object] = {}
    for r in rows:
        if r["kind"] == "skip":
            continue
        our = our_of.get(r["block"])
        if not our:
            continue
        if r["kind"] == "shelf":
            art = r["art"]
            if art in failed:
                out[r["row"]] = STATE_FAIL
            elif art in missing:
                out[r["row"]] = STATE_GONE
            else:
                out[r["row"]] = positions.get(our, {}).get(art) or NOT_IN_SHELF
        else:
            row_pos = positions.get(our, {})
            checked = [c for c in row_pos if c not in failed]
            found = sum(1 for c in checked if row_pos.get(c))
            # «5 из 12», а не «5/12»: Google на USER_ENTERED сделал бы из «5/12» дату.
            out[r["row"]] = f"{found} из {len(checked)}" if checked else ""
    return out


# ----------------------------------------------------------------------- запись

def write_column(book, ws, rows: list[dict], dates: list[str], values: list[list[str]],
                 today_vals: dict[int, object], when: datetime) -> tuple[str, int]:
    """Колонка за сегодня. Даты левее — свежая слева, прежние съезжают вправо.

    Колонки A–G не трогаем вообще: их ведёт человек, и любое переписывание
    затёрло бы его правки и раскраску.
    """
    today = when.strftime("%d.%m")
    nrows = len(values)

    if today in dates:
        col = NFIX + 1 + dates.index(today)          # повторный прогон в тот же день
    else:
        col = NFIX + 1
        if ws.col_count < NFIX + len(dates) + 1:
            ws.resize(rows=ws.row_count, cols=NFIX + len(dates) + 4)
        # inheritFromBefore=True — оформление берётся из колонки G слева, а не из
        # вчерашней даты справа: иначе в свежую колонку переехала бы её заливка.
        book.batch_update({"requests": [{"insertDimension": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": NFIX, "endIndex": NFIX + 1},
            "inheritFromBefore": True}}]})
        dates = [today] + dates

    letter = col_letter(col)
    column = [[today_vals.get(i, "")] for i in range(2, nrows + 1)]
    ws.update(values=[[today]], range_name=f"{letter}1", value_input_option="RAW")
    ws.update(values=column, range_name=f"{letter}2:{letter}{nrows}",
              value_input_option="USER_ENTERED")

    ndates = len(dates)
    reqs: list[dict] = []
    # Хвост истории глубже KEEP_DATES сносим — иначе лист растёт вправо без конца.
    if ndates > KEEP_DATES:
        reqs.append({"deleteDimension": {"range": {
            "sheetId": ws.id, "dimension": "COLUMNS",
            "startIndex": NFIX + KEEP_DATES, "endIndex": NFIX + ndates}}})
        ndates = KEEP_DATES

    last_col = NFIX + ndates
    reqs += [
        {"updateSheetProperties": {
            "properties": {"sheetId": ws.id,
                           "gridProperties": {"frozenRowCount": 1,
                                              "frozenColumnCount": 3}},
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}},
        {"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": NFIX, "endColumnIndex": last_col},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True},
                                           "horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat(textFormat,horizontalAlignment)"}},
        {"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": nrows,
                      "startColumnIndex": NFIX, "endColumnIndex": last_col},
            "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat.horizontalAlignment"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": NFIX, "endIndex": last_col},
            "properties": {"pixelSize": 60}, "fields": "pixelSize"}},
    ]

    # Условное форматирование пересоздаём целиком: правил всегда шесть, и они
    # не размножаются от прогона к прогону.
    meta = book.fetch_sheet_metadata(
        {"fields": "sheets(properties.sheetId,conditionalFormats)"})
    for s in meta.get("sheets", []):
        if s["properties"]["sheetId"] != ws.id:
            continue
        for k in range(len(s.get("conditionalFormats", [])) - 1, -1, -1):
            reqs.append({"deleteConditionalFormatRule": {"sheetId": ws.id, "index": k}})

    rng = {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": nrows,
           "startColumnIndex": NFIX, "endColumnIndex": last_col}
    prev = 0
    for limit, color in BANDS:
        reqs.append({"addConditionalFormatRule": {"index": 0, "rule": {
            "ranges": [rng],
            "booleanRule": {"condition": {"type": "NUMBER_BETWEEN", "values": [
                {"userEnteredValue": str(prev + 1)}, {"userEnteredValue": str(limit)}]},
                "format": {"backgroundColor": color}}}}})
        prev = limit
    reqs.append({"addConditionalFormatRule": {"index": 0, "rule": {
        "ranges": [rng],
        "booleanRule": {"condition": {"type": "NUMBER_GREATER", "values": [
            {"userEnteredValue": str(BANDS[-1][0])}]},
            "format": {"backgroundColor": COLOR_DEEP}}}}})
    reqs.append({"addConditionalFormatRule": {"index": 0, "rule": {
        "ranges": [rng],
        "booleanRule": {"condition": {"type": "TEXT_EQ", "values": [
            {"userEnteredValue": NOT_IN_SHELF}]},
            "format": {"backgroundColor": COLOR_NONE}}}}})

    for i in range(0, len(reqs), 300):
        book.batch_update({"requests": reqs[i:i + 300]})
    return today, ndates


# ------------------------------------------------------------------------- main

def main() -> None:
    ap = argparse.ArgumentParser(
        description="Позиции NATURI в полках → лист «Полки» отдельной книги бренда")
    ap.add_argument("--sheet-id",
                    default=os.environ.get("NATURI_SHEET_ID") or NATURI_SHEET_ID)
    ap.add_argument("--template-id",
                    default=os.environ.get("SHEET_ID") or TEMPLATE_SHEET_ID)
    ap.add_argument("--template-tab", default=TEMPLATE_TAB)
    ap.add_argument("--creds", default=None)
    ap.add_argument("--dest", type=int, default=sp.DEST_MOSCOW)
    ap.add_argument("--workers", type=int, default=4)
    ap.add_argument("--max-positions", type=int, default=600)
    ap.add_argument("--save-snapshot", default="snapshot_naturi.json")
    ap.add_argument("--snapshot-in", default=None,
                    help="взять готовый снапшот вместо обхода полок (отладка)")
    ap.add_argument("--dry-run", action="store_true", help="в таблицу не писать")
    args = ap.parse_args()

    install_retries()
    client = ts.get_client(args.creds)
    book = client.open_by_key(args.sheet_id)
    try:
        ws = book.worksheet(SHEET_DST)
    except gspread.WorksheetNotFound:
        ws = bootstrap(client, book, args.template_id, args.template_tab)

    rows, dates, values = read_sheet(ws)
    groups = build_groups(rows)
    shelves = {c for g in groups for c in g["competitors"]}
    sp.log(f"Лист «{SHEET_DST}»: строк {len(values) - 1}, групп {len(groups)}, "
           f"наших карточек {len(groups)}, уникальных полок к обходу {len(shelves)}, "
           f"дат в истории {len(dates)}")

    if args.snapshot_in:
        with open(args.snapshot_in, encoding="utf-8") as f:
            snapshot = json.load(f)
    else:
        snapshot = sp.run_groups(groups, dest=args.dest, workers=args.workers,
                                 max_positions=args.max_positions)
        if args.save_snapshot:
            with open(args.save_snapshot, "w", encoding="utf-8") as f:
                json.dump(snapshot, f, ensure_ascii=False, indent=2)
            sp.log(f"Снапшот сохранён: {args.save_snapshot}")

    vals = cell_values(snapshot, groups, rows)
    nums = [v for v in vals.values() if isinstance(v, int)]
    sp.log(f"Позиций найдено: {len(nums)}; нас нет в полке: "
           f"{sum(1 for v in vals.values() if v == NOT_IN_SHELF)}; "
           f"нет карточки: {sum(1 for v in vals.values() if v == STATE_GONE)}; "
           f"ошибок сбора: {sum(1 for v in vals.values() if v == STATE_FAIL)}"
           + (f"; медиана позиции {sorted(nums)[len(nums) // 2]}" if nums else ""))

    if args.dry_run:
        print("dry-run: в таблицу не пишу.")
        return

    date, ndates = write_column(book, ws, rows, dates, values, vals, datetime.now(sp.MSK))
    print(f"Готово: лист «{SHEET_DST}», колонка за {date} записана "
          f"(строк {len(values) - 1}, дат в истории {ndates}).")


if __name__ == "__main__":
    main()
