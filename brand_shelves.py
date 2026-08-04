#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Отдельная книга бренда: позиции карточек бренда в полках конкурентов.

Задача Артура от 04.08.2026: разнести бренды по отдельным таблицам. Образец —
лист «Натури пример» в общей книге «Анализ конкурентов ВБ авто»: плоский список,
где строка = полка (карточка конкурента) внутри группы «наш товар», а в колонках
дат стоит позиция НАШЕЙ карточки в этой полке, свежая дата слева.

Брендов четыре, книга у каждого своя, а список конкурентов общий: он один и тот
же в каждой группе, меняется только контрольная карточка. Поэтому лист бренда
заводится из того же образца — строки бренда поднимаются в начало своей группы,
первая из них (та, у которой «Тип конкурента» совпадает с типом группы) и есть
контрольная карточка, остальные строки остаются полками в прежнем порядке.

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

Запуск: шагом в shelves.yml (08:00 МСК), кнопкой в самой книге (workflow
brand-shelves.yml) либо руками:
    python brand_shelves.py --brand all --creds ключ.json
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

# Книги брендов (папка «Анализ конкурентов» на Drive). ID не секрет.
# `aliases` — как бренд написан в колонке «Бренд конкурента» образца; сравнение
# точное по нормализованной строке, иначе «Health Form» поймал бы «HealthIs».
BRANDS: dict[str, dict] = {
    "NATURI": {"sheet_id": "1XLby8VEOKQtuXrm4OCiQ-PFiTMXfaUe7gF054feoh_0",
               "aliases": ["NATURI"]},
    "SUNSHINE": {"sheet_id": "1xustRP7HtPHNiZTRnGZ1FU7T_nPLnyeP9PD4VxUvw1o",
                 "aliases": ["Sunshine Nutrition", "SUNSHINE", "Sunshine"]},
    "Health Form": {"sheet_id": "1KxORIJezgLt85dLl6D_wHnsRYfNMwn1as9S_L32Swvc",
                    "aliases": ["Health Form", "HealthForm"]},
    "4ME": {"sheet_id": "1slwJjO4mEn7umu6vH0bMbzuWGuAkfeVdekzTUExj-Ro",
            "aliases": ["4Me Nutrition", "4ME", "4Me"]},
}

# Откуда взять раскладку при самом первом запуске (книга «Анализ конкурентов ВБ авто»).
TEMPLATE_SHEET_ID = "1hqCt4QnCnqrLrRUZD3hSCuDd3k-2PFpaZdHNaxE4Nzk"
TEMPLATE_TAB = "Натури пример"

SHEET_DST = "Полки"

COLOR_CONTROL = {"red": 0.42, "green": 0.66, "blue": 0.31}   # контрольная карточка
COLOR_OURS = {"red": 0.85, "green": 0.92, "blue": 0.83}      # прочие карточки бренда
WHITE = {"red": 1, "green": 1, "blue": 1}
COL_WIDTHS = [330, 115, 150, 140, 110, 80, 100]              # A..G

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

def norm_brand(text: str) -> str:
    return " ".join(str(text).split()).strip().lower()


def reorder_for_brand(rows: list[list[str]], aliases: list[str]) -> tuple[list[list[str]], int]:
    """Строки образца → строки листа бренда: его карточки в начале своей группы.

    Список конкурентов у всех брендов общий, разная только контрольная карточка.
    В образце (он нарисован под NATURI) карточки бренда стоят где-то в середине
    группы — поднимаем их наверх, а контрольной делаем ту, у которой «Тип
    конкурента» совпадает с типом группы (тип группы = тип её первой строки в
    образце). Совпадения нет — берём первую карточку бренда.
    """
    ours = {norm_brand(a) for a in aliases}
    out: list[list[str]] = []
    controls = 0

    block: list[list[str]] = []
    def flush() -> None:
        nonlocal controls
        if not block:
            return
        kind = block[0][4] if len(block[0]) > 4 else ""
        mine = [r for r in block if norm_brand(r[2]) in ours]
        rest = [r for r in block if norm_brand(r[2]) not in ours]
        if mine:
            same = [r for r in mine if (r[4] if len(r) > 4 else "") == kind]
            control = same[0] if same else mine[0]
            mine = [control] + [r for r in mine if r is not control]
            controls += 1
        out.extend(mine + rest)
        block.clear()

    product = ""
    for raw in rows:
        r = [str(x).strip() for x in (list(raw) + [""] * NFIX)[:NFIX]]
        if not any(r):
            flush()
            out.append(r)
            product = ""
            continue
        if r[0] and r[0] != product:
            flush()
            product = r[0]
        r[0] = product          # товар дублируем в каждой строке: после
        block.append(r)         # перестановки первая строка группы уже другая
    flush()
    return out, controls


def bootstrap(client, book, brand: str, template_id: str, template_tab: str):
    """Первый запуск: лист бренда из образца «Натури пример».

    Значения переписываются (строки бренда переезжают в начало своих групп),
    поэтому оформление задаём сами: ширины, шапка, заморозка и заливка строк
    наших карточек — тёмно-зелёная у контрольной, светло-зелёная у остальных.
    Дальше лист ведёт человек, и скрипт эти колонки больше не трогает.
    """
    src = client.open_by_key(template_id).worksheet(template_tab)
    sp.log(f"Листа «{SHEET_DST}» в книге нет — собираю из «{template_tab}» под {brand}")
    values = src.get_all_values()
    head = [str(x).strip() for x in values[0]][:NFIX]
    body, controls = reorder_for_brand(values[1:], BRANDS[brand]["aliases"])
    if not controls:
        raise SystemExit(f"В образце нет ни одной карточки бренда {brand} — нечего мерить")

    ws = book.add_worksheet(title=SHEET_DST, rows=len(body) + 200, cols=NFIX + 10)
    ws.update_index(0)
    ws.update(values=[head] + body, range_name=f"A1:{col_letter(NFIX)}{len(body) + 1}",
              value_input_option="USER_ENTERED")

    ours = {norm_brand(a) for a in BRANDS[brand]["aliases"]}
    reqs = [
        {"updateSheetProperties": {
            "properties": {"sheetId": ws.id,
                           "gridProperties": {"frozenRowCount": 1, "frozenColumnCount": 3}},
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}},
        {"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": NFIX},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True},
                                           "wrapStrategy": "WRAP",
                                           "horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat(textFormat,wrapStrategy,horizontalAlignment)"}},
    ]
    for i, w in enumerate(COL_WIDTHS):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": i, "endIndex": i + 1},
            "properties": {"pixelSize": w}, "fields": "pixelSize"}})

    product = ""
    for i, r in enumerate(body, start=2):
        if not any(r):
            product = ""
            continue
        first = r[0] != product
        product = r[0]
        if norm_brand(r[2]) not in ours:
            continue
        color = COLOR_CONTROL if first else COLOR_OURS
        reqs.append({"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": i - 1, "endRowIndex": i,
                      "startColumnIndex": 0, "endColumnIndex": NFIX},
            "cell": {"userEnteredFormat": {"backgroundColor": color}},
            "fields": "userEnteredFormat.backgroundColor"}})

    for i in range(0, len(reqs), 300):
        book.batch_update({"requests": reqs[i:i + 300]})
    sp.log(f"Лист заведён: строк {len(body)}, контрольных карточек {controls}")
    return ws


def read_layout(head: list[str]) -> dict:
    """Шапка → раскладка листа: где фиксированные колонки и где какая дата.

    Ничего не зашито по буквам: границей служит ПЕРВАЯ колонка-дата, а нужные
    поля ищутся по названию. Так лист переживает и добавленную человеком
    колонку («Комментарий»), и перестановку колонок местами.
    """
    date_at: dict[str, int] = {}          # дата → индекс колонки (0-based)
    for i, cell in enumerate(head):
        d = norm_date(cell)
        if d and d not in date_at:
            date_at[d] = i
    nfix = min(date_at.values()) if date_at else len([h for h in head if h])
    nfix = max(nfix, 3)                   # A–C нужны всегда

    def col_of(title: str, default: int) -> int:
        for i, cell in enumerate(head[:nfix]):
            if cell.lower() == title.lower():
                return i
        return default

    return {"nfix": nfix, "date_at": date_at,
            "product": col_of(FIX_COLS[0], 0),
            "art": col_of(FIX_COLS[1], 1),
            "brand": col_of(FIX_COLS[2], 2)}


def read_sheet(ws, aliases: list[str]) -> tuple[list[dict], dict, list[list[str]]]:
    """Лист → (строки, раскладка, сырые значения).

    Строка: {"row", "kind": "our"|"shelf"|"skip", "art", "block"}.

    Строки Артур правит руками в любой момент: вставленный конкурент просто
    появляется в разборе следующим прогоном и с этого дня заполняется, а прежние
    даты у него остаются пустыми. Товар в колонке A можно не дублировать —
    пустая A наследует группу строки выше (так делают, когда вставляют строку
    внутрь блока).
    """
    values = ws.get_all_values()
    if not values:
        raise SystemExit(f"Лист «{SHEET_DST}» пуст")
    head = [str(x).strip() for x in values[0]]
    lay = read_layout(head)
    ic, ia, ib = lay["product"], lay["art"], lay["brand"]
    width = max(lay["nfix"], ic, ia, ib) + 1
    ours = {norm_brand(a) for a in aliases}

    rows: list[dict] = []
    block = None
    block_has_our = set()
    for i, raw in enumerate(values[1:], start=2):
        cells = [str(x).strip() for x in (list(raw) + [""] * width)[:width]]
        product, art, brand = cells[ic], cells[ia], cells[ib]
        if not product and not art and not brand:
            # Пустая строка — конец блока: следующий артикул без товара в A
            # не должен приклеиться к группе, стоящей выше через разрыв.
            rows.append({"row": i, "kind": "skip", "art": "", "block": None})
            block = None
            continue
        if product:
            block = product
        elif block is None:
            rows.append({"row": i, "kind": "skip", "art": "", "block": None})
            continue
        is_our = norm_brand(brand) in ours
        # Наша карточка группы — первая наша строка блока. Остальные наши карточки
        # (в образце светло-зелёные) считаются полками: в них тоже интересно стоять.
        kind = "shelf"
        if not art.isdigit():
            # Артикул не заполнен (в образце так у наливной L-Carnitine) — строка
            # не полка и не наша карточка; блок ещё может получить нашу ниже.
            kind = "skip"
        elif is_our and block not in block_has_our:
            kind = "our"
            block_has_our.add(block)
        rows.append({"row": i, "kind": kind, "art": art, "block": block})
    return rows, lay, values


def misplaced_controls(rows: list[dict]) -> list[str]:
    """Группы, где контрольная карточка стоит не первой строкой.

    Контрольная — первая строка группы С НАШИМ брендом; если выше неё в группе
    стоят конкуренты, «N из M» окажется в середине блока, а глазами это читается
    как съехавшая строка. Сами строки не переставляем (их ведёт человек), только
    говорим, где непорядок.
    """
    first: dict[str, int] = {}
    for r in rows:
        if r["block"] and r["block"] not in first:
            first[r["block"]] = r["row"]
    return [f"{r['block']}: контрольная в строке {r['row']}, "
            f"а группа начинается со строки {first[r['block']]}"
            for r in rows if r["kind"] == "our" and r["row"] != first[r["block"]]]


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
    visited = set(snapshot.get("shelves", {}))     # какие полки вообще обходили
    our_of = {g["product"]: str(g["ours"][0]) for g in groups}

    out: dict[int, object] = {}
    for r in rows:
        if r["kind"] == "skip":
            continue
        our = our_of.get(r["block"])
        if not our or our not in positions:
            # Контрольной карточки в снапшоте нет (группу завели уже после
            # обхода) — колонку за сегодня по ней не заполняем вообще.
            continue
        if r["kind"] == "shelf":
            art = r["art"]
            if visited and art not in visited:
                # Строку добавили руками уже после обхода — данных за сегодня по
                # ней просто нет. Оставляем пусто: «—» означало бы «проверили и
                # нас там нет», а мы её не проверяли.
                out[r["row"]] = ""
            elif art in failed:
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

def write_column(book, ws, lay: dict, values: list[list[str]],
                 today_vals: dict[int, object], when: datetime) -> tuple[str, int]:
    """Колонка за сегодня. Даты левее — свежая слева, прежние съезжают вправо.

    Фиксированные колонки не трогаем вообще: их ведёт человек, и любое
    переписывание затёрло бы его правки и раскраску. Колонка ищется по её
    фактическому индексу в шапке, а не по порядковому номеру среди дат, —
    иначе вставленная человеком колонка сдвинула бы запись на соседнюю дату.
    """
    today = when.strftime("%d.%m")
    nrows = len(values)
    nfix = lay["nfix"]

    if today in lay["date_at"]:
        col = lay["date_at"][today] + 1              # повторный прогон в тот же день
        ndates = len(lay["date_at"])
    else:
        col = nfix + 1
        if ws.col_count < nfix + len(lay["date_at"]) + 1:
            ws.resize(rows=ws.row_count, cols=nfix + len(lay["date_at"]) + 4)
        # inheritFromBefore=True — оформление берётся из колонки слева (последней
        # фиксированной), а не из вчерашней даты: иначе переехала бы её заливка.
        book.batch_update({"requests": [{"insertDimension": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": nfix, "endIndex": nfix + 1},
            "inheritFromBefore": True}}]})
        ndates = len(lay["date_at"]) + 1

    letter = col_letter(col)
    column = [[today_vals.get(i, "")] for i in range(2, nrows + 1)]
    ws.update(values=[[today]], range_name=f"{letter}1", value_input_option="RAW")
    ws.update(values=column, range_name=f"{letter}2:{letter}{nrows}",
              value_input_option="USER_ENTERED")

    reqs: list[dict] = []
    # Хвост истории глубже KEEP_DATES сносим — иначе лист растёт вправо без конца.
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
        {"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": nfix, "endColumnIndex": last_col},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True},
                                           "horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat(textFormat,horizontalAlignment)"}},
        {"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": nrows,
                      "startColumnIndex": nfix, "endColumnIndex": last_col},
            "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat.horizontalAlignment"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": nfix, "endIndex": last_col},
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
           "startColumnIndex": nfix, "endColumnIndex": last_col}
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

def run_brand(client, brand: str, args) -> str:
    """Один бренд: завести лист при необходимости, обойти полки, дописать колонку."""
    cfg = BRANDS[brand]
    sp.log(f"=== {brand} ===")
    book = client.open_by_key(os.environ.get(f"SHEET_ID_{brand.upper().replace(' ', '_')}")
                              or cfg["sheet_id"])
    try:
        ws = book.worksheet(SHEET_DST)
    except gspread.WorksheetNotFound:
        if args.dry_run:
            return f"{brand}: листа нет, dry-run — не завожу"
        ws = bootstrap(client, book, brand, args.template_id, args.template_tab)

    rows, lay, values = read_sheet(ws, cfg["aliases"])
    groups = build_groups(rows)
    shelves = {c for g in groups for c in g["competitors"]}
    sp.log(f"Лист «{SHEET_DST}»: строк {len(values) - 1}, групп {len(groups)}, "
           f"уникальных полок к обходу {len(shelves)}, "
           f"фиксированных колонок {lay['nfix']}, дат в истории {len(lay['date_at'])}")
    for line in misplaced_controls(rows):
        sp.log(f"ВНИМАНИЕ, карточка бренда не первой строкой группы — {line}")

    snap_path = args.snapshot_in or f"snapshot_{brand.lower().replace(' ', '_')}.json"
    if args.snapshot_in:
        with open(args.snapshot_in, encoding="utf-8") as f:
            snapshot = json.load(f)
    else:
        snapshot = sp.run_groups(groups, dest=args.dest, workers=args.workers,
                                 max_positions=args.max_positions)
        if args.save_snapshots:
            with open(snap_path, "w", encoding="utf-8") as f:
                json.dump(snapshot, f, ensure_ascii=False, indent=2)
            sp.log(f"Снапшот сохранён: {snap_path}")

    if not args.dry_run:
        # Обход полок идёт ~4 минуты, и всё это время Артур может править строки.
        # Колонка пишется по НОМЕРАМ строк, поэтому раскладку перечитываем прямо
        # перед записью: иначе значения лягут на старую нумерацию и разъедутся
        # (так вышло 04.08 в книге 4ME). Полки, появившиеся после обхода,
        # остаются пустыми — их сегодня не проверяли.
        fresh_rows, fresh_lay, fresh_values = read_sheet(ws, cfg["aliases"])
        if len(fresh_values) != len(values) or [r["art"] for r in fresh_rows] != [
                r["art"] for r in rows]:
            sp.log("Строки листа изменились за время обхода — пишу по свежей раскладке")
        rows, lay, values = fresh_rows, fresh_lay, fresh_values
        groups = build_groups(rows)

    vals = cell_values(snapshot, groups, rows)
    nums = [v for v in vals.values() if isinstance(v, int)]
    stat = (f"позиций {len(nums)}, нас нет в полке "
            f"{sum(1 for v in vals.values() if v == NOT_IN_SHELF)}, "
            f"нет карточки {sum(1 for v in vals.values() if v == STATE_GONE)}, "
            f"ошибок сбора {sum(1 for v in vals.values() if v == STATE_FAIL)}"
            + (f", медиана {sorted(nums)[len(nums) // 2]}" if nums else ""))
    sp.log(stat)

    if args.dry_run:
        return f"{brand}: dry-run, {stat}"
    date, ndates = write_column(book, ws, lay, values, vals, datetime.now(sp.MSK))
    return f"{brand}: колонка за {date} записана ({stat}; дат в истории {ndates})"


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Позиции бренда в полках → лист «Полки» его отдельной книги")
    ap.add_argument("--brand", default="all",
                    help="NATURI | SUNSHINE | 'Health Form' | 4ME | all")
    ap.add_argument("--template-id",
                    default=os.environ.get("SHEET_ID") or TEMPLATE_SHEET_ID)
    ap.add_argument("--template-tab", default=TEMPLATE_TAB)
    ap.add_argument("--creds", default=None)
    ap.add_argument("--dest", type=int, default=sp.DEST_MOSCOW)
    ap.add_argument("--workers", type=int, default=4)
    ap.add_argument("--max-positions", type=int, default=600)
    ap.add_argument("--save-snapshots", action="store_true", default=True)
    ap.add_argument("--snapshot-in", default=None,
                    help="взять готовый снапшот вместо обхода полок (отладка)")
    ap.add_argument("--dry-run", action="store_true", help="в таблицу не писать")
    args = ap.parse_args()

    if args.brand.lower() == "all":
        brands = list(BRANDS)
    else:
        match = [b for b in BRANDS if b.lower() == args.brand.strip().lower()]
        if not match:
            raise SystemExit(f"Неизвестный бренд {args.brand!r}; известны: {list(BRANDS)}")
        brands = match

    install_retries()
    client = ts.get_client(args.creds)

    results, failed = [], []
    for brand in brands:
        try:
            results.append(run_brand(client, brand, args))
        except Exception as exc:
            # Один бренд не должен ронять остальные: книга может быть закрыта,
            # лист переименован, полки не отдаться — остальные три обновятся.
            sp.log(f"ОШИБКА по бренду {brand}: {exc.__class__.__name__}: {exc}")
            failed.append(brand)
    for line in results:
        print("Готово:", line)
    if failed:
        raise SystemExit(f"Бренды с ошибкой: {', '.join(failed)}")


if __name__ == "__main__":
    main()
