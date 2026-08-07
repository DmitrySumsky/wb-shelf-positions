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
полками: в них ищется карточка группы. Если у этой первой строки артикул не
заполнен («нет», прочерк, пусто — товар ещё не заведён на WB), группа не мерится
вообще и колонку за день не получает: роль контрольной НЕ переходит к следующей
нашей карточке, иначе группа наберёт позиции другого товара (v1.2).

Что в ячейке даты:
    строка полки        — позиция нашей карточки (число),
                          «—» = полку проверили, нас там нет,
                          «нет карточки» = артикул полки удалён/скрыт на WB,
                          «ошибка сбора» = полка не отдалась (это НЕ «нас там нет»);
    строка нашей карточки — «N из M»: в скольких полках группы нашлись;
    вся группа целиком   — «нет такого товара»: у бренда нет карточки, мерить
                          нечем (v1.4; раньше такие группы оставались пустыми и
                          выглядели как обрыв прогона).

Новый товар (v1.3, 06.08.2026) заводится в ОДНОЙ книге-доноре (по умолчанию
NATURI, её ведёт Артур), а в остальные группа переезжает сама с `--sync-groups`:
полки те же, контрольной встаёт карточка этого бренда — её артикул в книге-доноре
уже стоит строкой-конкурентом. Карточки у бренда нет — приезжает «нет», и группа
честно остаётся неизмеримой. Существующие строки книг при этом не трогаются:
новые группы дописываются в конец листа.

Запуск: шагом в shelves.yml (08:00 МСК), кнопкой в самой книге (workflow
brand-shelves.yml) либо руками:
    python brand_shelves.py --brand all --creds ключ.json
    python brand_shelves.py --brand all --sync-groups --creds ключ.json
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

# v1.3. Что переносится при добавлении товара из книги-донора: только описание
# полки. «Ключ», «Прогрев», «Артикул для прогрева» у каждого бренда свои — их
# ведёт менеджер книги, и чужие значения там были бы враньём.
SYNC_FIELDS = ["Товар", "Артикул конкурента", "Бренд конкурента",
               "~ Выручка конкурента", "Тип конкурента"]
SYNC_FROM = "NATURI"              # книга-донор по умолчанию: её ведёт Артур
# Товары, которые в другой книге называются иначе и потому не переносятся:
# у 4ME хром 270 мг, у донора — 250 мг, товар один и тот же.
SYNC_SKIP = ["Chromium Picolinate 250 mg 120 caps"]

NOT_IN_SHELF = "—"
STATE_GONE = "нет карточки"
STATE_FAIL = "ошибка сбора"
# v1.4. Группа, которую мерить нечем (у бренда нет карточки товара). Раньше её
# строки оставались пустыми, и по листу это читалось как «скрипт сломался на
# середине». Пишем прямым текстом — видно, что прошли весь лист.
STATE_NO_PRODUCT = "нет такого товара"

# Раскраска позиций — условным форматированием (шесть правил на любой размер листа),
# как в листе «Полки» книги VEXOR: ячеек тут десятки тысяч, красить их поштучно дорого.
BANDS = [  # шесть правил на позиции + седьмое на текстовые пометки
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


def read_sheet(ws, aliases: list[str]) -> tuple[list[dict], dict, list[list[str]], set]:
    """Лист → (строки, раскладка, сырые значения, группы без карточки бренда).

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
    block_has_our: set[str] = set()
    no_card: set[str] = set()
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
        # Наша карточка группы — ПЕРВАЯ наша строка блока, и роль контрольной у неё
        # уже не отнять. Остальные наши карточки (в образце светло-зелёные)
        # считаются полками: в них тоже интересно стоять.
        kind = "shelf"
        if is_our and block not in block_has_our:
            block_has_our.add(block)
            if art.isdigit():
                kind = "our"
            else:
                # v1.2. Карточки товара ещё нет: в артикуле «нет», прочерк или
                # пусто. Раньше контрольной молча становилась СЛЕДУЮЩАЯ наша
                # карточка блока — и группа получала чужие позиции, неотличимые
                # от настоящих (05.08: `Choline + Inositol + Ginkgo` мерился по
                # холину, `AAKG 180` — по 90-капсульному). Группа помечается как
                # неизмеримая и колонку за день не получает вовсе.
                kind = "skip"
                no_card.add(block)
        elif not art.isdigit():
            # Артикул не заполнен (в образце так у наливной L-Carnitine) — строка
            # не полка и не наша карточка.
            kind = "skip"
        rows.append({"row": i, "kind": kind, "art": art, "block": block})
    return rows, lay, values, no_card


# ------------------------------------------------- новые товары: перенос из книги-донора

def blocks_of(head: list[str], values: list[list[str]],
              lay: dict) -> tuple[list[str], dict[str, list[dict]]]:
    """Лист → группы: порядок товаров и строки каждой группы как {заголовок: значение}.

    Берётся только фиксированная часть листа: колонки-даты у книг свои, чужую
    историю переносить нельзя. Заголовки — ключи в нижнем регистре, поэтому
    книги с разным порядком колонок (в 4ME «Бренд» стоит перед «Артикулом»)
    сходятся между собой без карт соответствия.
    """
    nfix = lay["nfix"]
    ic = lay["product"]
    titles = [str(head[i]).strip().lower() if i < len(head) else "" for i in range(nfix)]

    order: list[str] = []
    blocks: dict[str, list[dict]] = {}
    block = None
    for raw in values[1:]:
        cells = [str(x).strip() for x in (list(raw) + [""] * nfix)[:nfix]]
        if not any(cells):
            block = None
            continue
        if cells[ic]:
            block = cells[ic]
        elif block is None:
            continue
        row = {titles[i]: cells[i] for i in range(nfix) if titles[i]}
        row[titles[ic]] = block           # товар дублируем в каждой строке группы
        if block not in blocks:
            blocks[block] = []
            order.append(block)
        blocks[block].append(row)
    return order, blocks


def rows_for_brand(rows: list[dict], product: str, brand: str,
                   used: set[str] | None = None) -> list[dict]:
    """Строки группы из книги-донора → порядок для книги бренда.

    Тот же приём, что при заведении листа: карточки бренда поднимаются в начало
    группы, контрольной становится та, у которой «Тип конкурента» совпадает с
    типом группы. Карточка бренда-донора при этом становится обычной полкой.

    Два случая, когда контрольной карточки честно нет (v1.3):
    строк бренда в группе нет вовсе (товара у него не бывает) — и артикул уже
    работает контрольным в другой группе этой книги (так у `Choline` и
    `Choline + Inositol + Ginkgo`: во второй группе стоит карточка первой).
    Тогда в начало группы встаёт строка с «нет» в артикуле: группа видна в
    листе, но не мерится — правило v1.2 «чужой карточкой не мерим».
    """
    kt, ka, kb, kk = (SYNC_FIELDS[0].lower(), SYNC_FIELDS[1].lower(),
                      SYNC_FIELDS[2].lower(), SYNC_FIELDS[4].lower())
    used = used or set()
    ours = {norm_brand(a) for a in BRANDS[brand]["aliases"]}
    kind = rows[0].get(kk, "") if rows else ""
    mine = [r for r in rows if norm_brand(r.get(kb, "")) in ours]
    rest = [r for r in rows if norm_brand(r.get(kb, "")) not in ours]

    same = [r for r in mine if r.get(kk, "") == kind]
    order = same + [r for r in mine if r not in same]
    control = next((r for r in order
                    if r.get(ka, "").isdigit() and r.get(ka, "") not in used), None)
    if control is None:
        control = next((r for r in order if not r.get(ka, "").isdigit()), None)
    if control is None:
        control = {kt: product, ka: "нет", kb: BRANDS[brand]["aliases"][0], kk: kind}
        mine = [control] + mine
    else:
        mine = [control] + [r for r in mine if r is not control]
    return mine + rest


def sync_groups(book, ws, brand: str, head: list[str], values: list[list[str]], lay: dict,
                ref_head: list[str], ref_values: list[list[str]], ref_lay: dict,
                dry_run: bool = False, skip: set[str] | None = None) -> list[str]:
    """Дописать в лист товары, которые есть в книге-доноре, а здесь ещё нет.

    Список конкурентов у брендов общий, поэтому новый товар достаточно завести
    в одной книге — в остальные группа переносится целиком: полки в том же
    порядке, контрольной встаёт карточка этого бренда (её артикул в книге-доноре
    уже есть строкой-конкурентом). Карточки у бренда нет — строка приезжает с
    «нет» в артикуле, группа честно остаётся неизмеримой до появления товара.

    Существующие строки не трогаются вообще: новые группы дописываются в конец
    листа через пустую строку-разделитель.
    """
    ref_order, ref_blocks = blocks_of(ref_head, ref_values, ref_lay)
    _, cur_blocks = blocks_of(head, values, lay)
    # `skip` — тот же товар, названный в книге иначе: у 4ME хром 270 мг, а в
    # книге-доноре 250 мг. Совпадение имён проверять нечем, поэтому список ведёт
    # человек, иначе в лист приедет вторая группа того же товара.
    skip = {norm_brand(s) for s in (skip or set())}
    missing = [p for p in ref_order
               if p not in cur_blocks and norm_brand(p) not in skip]
    if not missing:
        return []

    nfix = lay["nfix"]
    titles = [str(head[i]).strip().lower() if i < len(head) else "" for i in range(nfix)]
    keep = {f.lower() for f in SYNC_FIELDS}
    ours = {norm_brand(a) for a in BRANDS[brand]["aliases"]}
    ka, kb = SYNC_FIELDS[1].lower(), SYNC_FIELDS[2].lower()

    # Артикулы, которые уже работают контрольными в этой книге: второй раз тот же
    # артикул контрольным не ставим — обе группы получили бы одни и те же позиции.
    used = set()
    for rows in cur_blocks.values():
        own = next((r for r in rows if norm_brand(r.get(kb, "")) in ours), None)
        if own and own.get(ka, "").isdigit():
            used.add(own[ka])

    out: list[list[str]] = []
    paint: list[tuple[int, dict]] = []      # (индекс в out, цвет)
    added: list[str] = []
    for product in missing:
        rows = rows_for_brand(ref_blocks[product], product, brand, used)
        control_art = rows[0].get(ka, "") if rows else ""
        if control_art.isdigit():
            used.add(control_art)
        out.append([""] * nfix)             # разделитель: группы не слипаются
        first_our = True
        for r in rows:
            if norm_brand(r.get(kb, "")) in ours:
                paint.append((len(out), COLOR_CONTROL if first_our else COLOR_OURS))
                first_our = False
            out.append([r.get(t, "") if t in keep else "" for t in titles])
        added.append(f"{product} ({len(rows)} строк"
                     + ("" if control_art.isdigit() else ", карточки бренда нет") + ")")

    if dry_run:
        return [f"dry-run, не дописано: {a}" for a in added]

    start = len(values) + 1                 # первая свободная строка листа
    end = start + len(out) - 1
    if ws.row_count < end:
        ws.add_rows(end - ws.row_count + 50)
    ws.update(values=out, range_name=f"A{start}:{col_letter(nfix)}{end}",
              value_input_option="USER_ENTERED")

    reqs = [{"repeatCell": {
        "range": {"sheetId": ws.id, "startRowIndex": start + i - 1, "endRowIndex": start + i,
                  "startColumnIndex": 0, "endColumnIndex": nfix},
        "cell": {"userEnteredFormat": {"backgroundColor": color}},
        "fields": "userEnteredFormat.backgroundColor"}} for i, color in paint]
    for i in range(0, len(reqs), 300):
        book.batch_update({"requests": reqs[i:i + 300]})
    return added


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


def build_groups(rows: list[dict], no_card: set | None = None) -> list[dict]:
    """Строки листа → группы для sp.run_groups.

    Группа, у которой карточка бренда заведена без артикула (`no_card`), не
    считается вообще: мерить нечем, а любое число в её строках было бы позицией
    другого товара.
    """
    no_card = no_card or set()
    order: list[str] = []
    acc: dict[str, dict] = {}
    for r in rows:
        if r["kind"] == "skip" or not r["block"] or r["block"] in no_card:
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
    if no_card:
        sp.log("Группы без артикула у карточки бренда пропущены (мерить нечем, "
               "в колонке за день будет «нет такого товара»): "
               + ", ".join(sorted(no_card)))

    # v1.2. Одна и та же карточка контрольной в двух группах — почти всегда
    # означает, что у одной из них своей карточки нет, а мы этого не заметили.
    seen: dict[int, str] = {}
    for g in good:
        nm = g["ours"][0]
        if nm in seen:
            sp.log(f"ВНИМАНИЕ: контрольная карточка {nm} сразу в двух группах — "
                   f"«{seen[nm]}» и «{g['product']}»; позиции у них совпадут")
        else:
            seen[nm] = g["product"]
    return good


# --------------------------------------------------------------- значения за сегодня

def cell_values(snapshot: dict, groups: list[dict], rows: list[dict]) -> dict[int, object]:
    """{номер строки листа: значение за сегодня}."""
    failed = set(snapshot.get("failed_shelves", []))
    missing = set(snapshot.get("missing_shelves", []))
    positions = snapshot.get("positions", {})
    visited = set(snapshot.get("shelves", {}))     # какие полки вообще обходили
    our_of = {g["product"]: str(g["ours"][0]) for g in groups}

    # v1.4. Группы без контрольной карточки: у бренда нет такого товара («нет»
    # или пусто в артикуле нашей строки) либо строк бренда в группе нет вовсе.
    # Помечаем ВСЕ строки такой группы, включая пустые-по-артикулу.
    with_our = {r["block"] for r in rows if r["kind"] == "our"}

    out: dict[int, object] = {}
    for r in rows:
        if r["block"] and r["block"] not in with_our:
            out[r["row"]] = STATE_NO_PRODUCT
            continue
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

    # Условное форматирование пересоздаём целиком: правил всегда семь, и они
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
    for text in (NOT_IN_SHELF, STATE_NO_PRODUCT):
        reqs.append({"addConditionalFormatRule": {"index": 0, "rule": {
            "ranges": [rng],
            "booleanRule": {"condition": {"type": "TEXT_EQ", "values": [
                {"userEnteredValue": text}]},
                "format": {"backgroundColor": COLOR_NONE}}}}})

    for i in range(0, len(reqs), 300):
        book.batch_update({"requests": reqs[i:i + 300]})
    return today, ndates


# ------------------------------------------------------------------------- main

def run_brand(client, brand: str, args, ref: tuple | None = None) -> str:
    """Один бренд: завести лист при необходимости, обойти полки, дописать колонку.

    `ref` — (шапка, значения, раскладка) листа книги-донора: заданы, значит перед
    обходом в лист доедут товары, которых в нём ещё нет (v1.3).
    """
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

    rows, lay, values, no_card = read_sheet(ws, cfg["aliases"])
    if ref is not None:
        added = sync_groups(book, ws, brand, [str(x).strip() for x in values[0]],
                            values, lay, *ref, dry_run=args.dry_run,
                            skip=set(filter(None, args.sync_skip.split(";"))))
        if added:
            sp.log(f"Новых товаров из книги-донора: {len(added)} — " + "; ".join(added))
            if not args.dry_run:
                rows, lay, values, no_card = read_sheet(ws, cfg["aliases"])

    groups = build_groups(rows, no_card)
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
        fresh_rows, fresh_lay, fresh_values, fresh_no_card = read_sheet(ws, cfg["aliases"])
        if len(fresh_values) != len(values) or [r["art"] for r in fresh_rows] != [
                r["art"] for r in rows]:
            sp.log("Строки листа изменились за время обхода — пишу по свежей раскладке")
        rows, lay, values, no_card = fresh_rows, fresh_lay, fresh_values, fresh_no_card
        groups = build_groups(rows, no_card)

    vals = cell_values(snapshot, groups, rows)
    nums = [v for v in vals.values() if isinstance(v, int)]
    stat = (f"позиций {len(nums)}, нас нет в полке "
            f"{sum(1 for v in vals.values() if v == NOT_IN_SHELF)}, "
            f"нет карточки {sum(1 for v in vals.values() if v == STATE_GONE)}, "
            f"ошибок сбора {sum(1 for v in vals.values() if v == STATE_FAIL)}, "
            f"строк без товара у бренда "
            f"{sum(1 for v in vals.values() if v == STATE_NO_PRODUCT)}"
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
    ap.add_argument("--sync-groups", action="store_true",
                    help="дописать в книги товары, которые есть в книге-доноре")
    ap.add_argument("--sync-from", default=SYNC_FROM, help="книга-донор новых товаров")
    ap.add_argument("--sync-skip", default=";".join(SYNC_SKIP),
                    help="товары, которые не переносить (через «;»)")
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

    # Книга-донор читается один раз на прогон: список конкурентов у брендов общий,
    # и новый товар достаточно завести в ней одной.
    donor, ref = None, None
    if args.sync_groups:
        match = [b for b in BRANDS if b.lower() == args.sync_from.strip().lower()]
        if not match:
            raise SystemExit(f"Неизвестная книга-донор {args.sync_from!r}")
        donor = match[0]
        dvals = client.open_by_key(BRANDS[donor]["sheet_id"]).worksheet(SHEET_DST) \
                      .get_all_values()
        dhead = [str(x).strip() for x in dvals[0]]
        ref = (dhead, dvals, read_layout(dhead))
        sp.log(f"Новые товары беру из книги {donor}")

    results, failed = [], []
    for brand in brands:
        try:
            results.append(run_brand(client, brand, args, None if brand == donor else ref))
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
