#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Лист «Полки» таблицы «Анализ цен конкурентов Вексор»: позиция НАШЕЙ карточки
в рекомендательной полке каждого конкурента, по дням.

Задача от 31.07.2026: в книге VEXOR уже есть лист «Сводная» (цены конкурентов
по дням, ведёт `21.orders-cloud/prices_update.py`) и ручные колонки «Прогрев» /
«Прогрев к-во». Нужен лист-близнец, где в тех же строках и в тех же колонках-датах
вместо цены стоит наша позиция в полке этого конкурента — чтобы прогрев было с чем
сопоставлять.

Раскладка «Полок» повторяет «Сводную» строка в строку:
    A  Название товара      \\
    B  Артикул ВБ            |  копируются из «Сводной» каждый прогон
    C  Прогрев               |  (единственный источник правды — «Сводная»)
    D  Прогрев к-во          |
    E  Бренд                /
    F  Средняя, 30 дн       — формула по 30 свежим датам
    G.. даты ДД.ММ, свежая СЛЕВА

Что в ячейке даты:
    строка конкурента — позиция нашей карточки в его полке (число),
                        «—» = полку проверили, нас там нет,
                        «нет карточки» = артикул конкурента удалён/скрыт на WB,
                        «ошибка сбора» = полка не отдалась (это НЕ «нас там нет»);
    строка VEXOR      — сводка по товару: «N/M» — в скольких полках блока нашлись.

Товарный блок = кусок «Сводной» между пустыми строками-разделителями (так они и
заведены: чёрная строка = разделитель, первая строка блока — наш VEXOR). Позиции
считаются ТОЛЬКО внутри блока: наша «Мастика» ищется в полках мастик, а не мовилей.
Если в блоке две наши карточки — в строке конкурента стоит ЛУЧШАЯ (минимальная)
из их позиций, а «N/M» у каждой нашей строки своё.

История: лист пересобирается целиком каждый прогон (строки следуют за «Сводной»,
даже если менеджер вставил конкурента в середину), прежние даты читаются с самого
листа и переносятся по ключу «блок + артикул + номер повтора». Глубина — KEEP_DATES
колонок, дальше самые старые отбрасываются.

Запуск: шагом в shelves.yml (08:00 МСК) либо руками:
    python vexor_shelves.py --creds ключ.json
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

# Книга «Анализ цен конкурентов Вексор». Не секрет: ID уже лежит в ключах
# отчёта заказов (PRICES_GSHEET_ID) и в памяти проекта.
VEXOR_SHEET_ID = "1arRAuQHixoZlkI0PVcTMZwgZiWQQPH3G7ADCnLm-a8w"

SHEET_SRC = "Сводная"
SHEET_DST = "Полки"
OUR_BRAND = "VEXOR"

# Колонки до блока дат. Первые пять копируются из «Сводной» как есть —
# в том числе ручные «Прогрев» и «Прогрев к-во».
FIX_COLS = ["Название товара", "Артикул ВБ", "Прогрев", "Прогрев к-во", "Бренд"]
AVG_COL = "Средняя, 30 дн"
NFIX = len(FIX_COLS) + 1          # A..F
KEEP_DATES = 90                   # сколько колонок истории держать

NOT_IN_SHELF = "—"
STATE_GONE = "нет карточки"
STATE_FAIL = "ошибка сбора"

COLOR_OURS = {"red": 0.85, "green": 0.94, "blue": 0.83}    # строка VEXOR
COLOR_SEP = {"red": 0.15, "green": 0.15, "blue": 0.15}     # чёрный разделитель
WHITE = {"red": 1, "green": 1, "blue": 1}

# Раскраска позиций: чем ближе к началу полки, тем зеленее. Делается условным
# форматированием (5 правил на весь блок дат), а не покраской ячеек: правил
# ровно пять на любой размер листа, а ячеек тут больше десяти тысяч.
BANDS = [
    (10, {"red": 0.72, "green": 0.88, "blue": 0.72}),
    (30, {"red": 0.85, "green": 0.94, "blue": 0.78}),
    (100, {"red": 1.00, "green": 0.95, "blue": 0.75}),
    (300, {"red": 0.99, "green": 0.87, "blue": 0.75}),
]
COLOR_DEEP = {"red": 0.96, "green": 0.80, "blue": 0.80}
COLOR_NONE = {"red": 0.93, "green": 0.93, "blue": 0.93}


# ------------------------------------------------------------------ чтение блоков

def read_blocks(book) -> tuple[list[dict], list[list[str]]]:
    """«Сводная» → (блоки, строки листа).

    Строка листа: {"kind": "sep"|"our"|"comp", "fix": [A..E], "block", "occ"}.
    Разделители сохраняются, чтобы «Полки» повторяли «Сводную» строка в строку.
    """
    ws = book.worksheet(SHEET_SRC)
    values = ws.get("A1:E2000")
    head = [str(x).strip() for x in (values[0] if values else [])]
    if head[:2] != FIX_COLS[:2]:
        raise SystemExit(f"Шапка «{SHEET_SRC}» не та, что ожидалась: {head[:5]}")

    rows: list[dict] = []
    blocks: list[dict] = []
    cur: dict | None = None
    seen_titles: dict[str, int] = {}

    for raw in values[1:]:
        cells = [str(x).strip() for x in (list(raw) + [""] * 5)[:5]]
        name, art, warm, warm_n, brand = cells
        if not any([name, art, brand]):
            rows.append({"kind": "sep", "fix": [""] * 5})
            cur = None
            continue
        if cur is None:
            title = name or f"(без названия, строка {len(rows) + 2})"
            seen_titles[title] = seen_titles.get(title, 0) + 1
            if seen_titles[title] > 1:                 # два блока с одним именем
                title = f"{title} #{seen_titles[title]}"
            cur = {"title": title, "ours": [], "competitors": []}
            blocks.append(cur)
        occ = sum(1 for r in rows
                  if r.get("block") == cur["title"] and r["fix"][1] == art)
        is_our = brand.upper() == OUR_BRAND.upper()
        if art.isdigit():
            (cur["ours"] if is_our else cur["competitors"]).append(int(art))
        rows.append({"kind": "our" if is_our else "comp", "fix": cells,
                     "block": cur["title"], "occ": occ})

    # Блок без наших карточек искать не в чем — полки его конкурентов не обходим.
    good = [b for b in blocks if b["ours"] and b["competitors"]]
    skipped = [b["title"] for b in blocks if b not in good]
    if skipped:
        sp.log(f"Блоки без нашей карточки или без конкурентов пропущены: {skipped}")
    groups = [{"product": b["title"], "ours": b["ours"],
               "competitors": b["competitors"]} for b in good]
    return groups, rows


# --------------------------------------------------------------- сборка значений

def cell_values(snapshot: dict, groups: list[dict], rows: list[dict]) -> dict:
    """{(блок, артикул, повтор): значение ячейки за сегодня}."""
    failed = set(snapshot.get("failed_shelves", []))
    missing = set(snapshot.get("missing_shelves", []))
    positions = snapshot.get("positions", {})
    ours_of = {g["product"]: [str(x) for x in g["ours"]] for g in groups}

    out: dict[tuple[str, str, int], object] = {}
    for r in rows:
        if r["kind"] == "sep":
            continue
        key = (r["block"], r["fix"][1], r["occ"])
        art = r["fix"][1]
        our_list = ours_of.get(r["block"], [])
        if not our_list or not art.isdigit():
            continue
        if r["kind"] == "comp":
            if art in failed:
                out[key] = STATE_FAIL
            elif art in missing:
                out[key] = STATE_GONE
            else:
                # Лучшая позиция среди наших карточек блока (обычно она одна).
                got = [positions.get(o, {}).get(art) for o in our_list]
                got = [p for p in got if p]
                out[key] = min(got) if got else NOT_IN_SHELF
        else:
            row_pos = positions.get(art, {})
            checked = [c for c in row_pos if c not in failed]
            found = sum(1 for c in checked if row_pos.get(c))
            # «5 из 12», а не «5/12»: Google на USER_ENTERED превратил бы
            # «5/12» в дату 12 мая
            out[key] = f"{found} из {len(checked)}" if checked else ""
    return out


# ----------------------------------------------------------------- запись листа

def col_letter(i1: int) -> str:
    s, i = "", i1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def norm_date(text: str) -> str | None:
    """«30.07» / «30.07.2026» → «30.07». Не дата — None.

    Google на USER_ENTERED легко превращает «30.07» в настоящую дату и показывает
    её уже с годом, поэтому шапка читается терпимо к формату, а сравнивается
    всегда в коротком виде.
    """
    m = re.fullmatch(r"(\d{1,2})\.(\d{1,2})(?:\.\d{2,4})?", str(text).strip())
    return f"{int(m.group(1)):02d}.{int(m.group(2)):02d}" if m else None


def read_history(ws) -> tuple[list[str], dict]:
    """Что уже есть на листе: (даты слева направо, {ключ: {дата: значение}}).

    Ключ строится ровно так же, как в read_blocks («блок + артикул + повтор»),
    иначе после вставки конкурента в середину «Сводной» история съедет на строку.
    """
    values = ws.get_all_values()
    if not values or len(values[0]) < NFIX:
        return [], {}
    head = [str(x).strip() for x in values[0]]
    if head[0] != FIX_COLS[0]:
        return [], {}          # лист пуст или прежнего образца — начинаем с нуля
    idx_of: dict[str, int] = {}
    for i in range(NFIX, len(head)):
        d = norm_date(head[i])
        if d and d not in idx_of:
            idx_of[d] = i
    dates = list(idx_of)

    hist: dict[tuple[str, str, int], dict[str, str]] = {}
    block, occ_seen, seen_titles = None, {}, {}
    for raw in values[1:]:
        row = [str(x).strip() for x in raw] + [""] * len(head)
        name, art, _, _, brand = row[:5]
        if not any([name, art, brand]):
            block, occ_seen = None, {}
            continue
        if block is None:
            seen_titles[name] = seen_titles.get(name, 0) + 1
            block = name if seen_titles[name] == 1 else f"{name} #{seen_titles[name]}"
        occ = occ_seen.get(art, 0)
        occ_seen[art] = occ + 1
        cells = {d: row[i] for d, i in idx_of.items() if row[i] != ""}
        if cells:
            hist[(block, art, occ)] = cells
    return dates, hist


def write_shelves(book, rows: list[dict], today_vals: dict, when: datetime) -> tuple[str, int]:
    """Полная пересборка листа «Полки» под текущую «Сводную»."""
    today = when.strftime("%d.%m")
    ws = ts.ensure_ws(book, SHEET_DST, rows=len(rows) + 40, cols=NFIX + 8)
    old_dates, hist = read_history(ws)

    # Даты: сегодня слева, дальше прежние по убыванию (как в «Сводной»).
    dates = [today] + [d for d in old_dates if d != today]
    dates = dates[:KEEP_DATES]

    header = FIX_COLS + [AVG_COL] + dates
    if ws.col_count < len(header):
        ws.resize(rows=ws.row_count, cols=len(header) + 4)
    body = []
    fmt_our, fmt_sep = [], []
    for i, r in enumerate(rows, start=2):
        if r["kind"] == "sep":
            body.append([""] * len(header))
            fmt_sep.append(i)
            continue
        key = (r["block"], r["fix"][1], r["occ"])
        old = hist.get(key, {})
        line = list(r["fix"])
        # Средняя считается от СВОЕЙ колонки: COLUMN() = 6 в F, значит окно
        # начинается с G — первой даты. Вставит менеджер колонку слева —
        # формула переедет вместе с блоком и останется верной.
        line.append(f"=IFERROR(ROUND(AVERAGE(OFFSET($A{i};0;COLUMN();1;"
                    f"MIN(30;COLUMNS($1:$1)-COLUMN())));1);\"\")")
        for d in dates:
            line.append(today_vals.get(key, "") if d == today else old.get(d, ""))
        body.append(line)
        if r["kind"] == "our":
            fmt_our.append(i)

    last_col = col_letter(len(header))
    nrows = len(body) + 1
    # Шапку пишем RAW: на USER_ENTERED Google превращает «31.07» в дату и
    # показывает её с годом — следующий прогон не узнал бы свою же колонку.
    ws.update(values=[header], range_name=f"A1:{last_col}1",
              value_input_option="RAW")
    ws.update(values=body, range_name=f"A2:{last_col}{nrows}",
              value_input_option="USER_ENTERED")
    # Хвост прежней раскладки (строк стало меньше / дат стало меньше) — вычистить,
    # иначе внизу останутся строки от прошлой версии «Сводной».
    if ws.row_count > nrows:
        ws.batch_clear([f"A{nrows + 1}:{col_letter(ws.col_count)}{ws.row_count}"])
    if ws.col_count > len(header):
        ws.batch_clear([f"{col_letter(len(header) + 1)}1:"
                        f"{col_letter(ws.col_count)}{nrows}"])

    reqs = [
        {"updateSheetProperties": {
            "properties": {"sheetId": ws.id,
                           "gridProperties": {"frozenRowCount": 1,
                                              "frozenColumnCount": NFIX}},
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}},
        {"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": nrows,
                      "startColumnIndex": 0, "endColumnIndex": len(header)},
            "cell": {"userEnteredFormat": {"backgroundColor": WHITE,
                                           "textFormat": {"bold": False}}},
            "fields": "userEnteredFormat(backgroundColor,textFormat.bold)"}},
        {"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": len(header)},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True},
                                           "horizontalAlignment": "CENTER",
                                           "wrapStrategy": "WRAP"}},
            "fields": "userEnteredFormat(textFormat,horizontalAlignment,wrapStrategy)"}},
        {"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": nrows,
                      "startColumnIndex": NFIX - 1, "endColumnIndex": len(header)},
            "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat.horizontalAlignment"}},
    ]
    for i, w in enumerate([320, 105, 80, 100, 140, 95]):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": i, "endIndex": i + 1},
            "properties": {"pixelSize": w}, "fields": "pixelSize"}})
    reqs.append({"updateDimensionProperties": {
        "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                  "startIndex": NFIX, "endIndex": len(header)},
        "properties": {"pixelSize": 55}, "fields": "pixelSize"}})

    def band(rows_1based, color, bold):
        out = []
        for r in rows_1based:                      # соседние строки не сливаем:
            out.append({"repeatCell": {            # их единицы, а код проще
                "range": {"sheetId": ws.id, "startRowIndex": r - 1, "endRowIndex": r,
                          "startColumnIndex": 0, "endColumnIndex": len(header)},
                "cell": {"userEnteredFormat": {"backgroundColor": color,
                                               "textFormat": {"bold": bold}}},
                "fields": "userEnteredFormat(backgroundColor,textFormat.bold)"}})
        return out

    reqs += band(fmt_our, COLOR_OURS, True)
    reqs += band(fmt_sep, COLOR_SEP, False)

    # Условное форматирование блока дат пересоздаём: правил ровно столько,
    # сколько полос, и они не размножаются от прогона к прогону.
    meta = book.fetch_sheet_metadata(
        {"fields": "sheets(properties.sheetId,conditionalFormats)"})
    for s in meta.get("sheets", []):
        if s["properties"]["sheetId"] != ws.id:
            continue
        for k in range(len(s.get("conditionalFormats", [])) - 1, -1, -1):
            reqs.append({"deleteConditionalFormatRule": {"sheetId": ws.id, "index": k}})

    rng = {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": nrows,
           "startColumnIndex": NFIX, "endColumnIndex": len(header)}
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
    return today, len(dates)


# ------------------------------------------------------------------------- main

def main() -> None:
    ap = argparse.ArgumentParser(description="Позиции в полках → лист «Полки» книги VEXOR")
    ap.add_argument("--sheet-id", default=os.environ.get("VEXOR_SHEET_ID") or VEXOR_SHEET_ID)
    ap.add_argument("--creds", default=None)
    ap.add_argument("--dest", type=int, default=sp.DEST_MOSCOW)
    ap.add_argument("--workers", type=int, default=4)
    ap.add_argument("--max-positions", type=int, default=600)
    ap.add_argument("--save-snapshot", default="snapshot_vexor.json")
    ap.add_argument("--snapshot-in", default=None,
                    help="взять готовый снапшот вместо обхода полок (отладка)")
    ap.add_argument("--dry-run", action="store_true", help="в таблицу не писать")
    args = ap.parse_args()

    client = ts.get_client(args.creds)
    book = client.open_by_key(args.sheet_id)
    groups, rows = read_blocks(book)
    shelves = sorted({c for g in groups for c in g["competitors"]})
    sp.log(f"«{SHEET_SRC}»: строк {len(rows)}, блоков {len(groups)}, "
           f"наших карточек {sum(len(g['ours']) for g in groups)}, "
           f"уникальных полок к обходу {len(shelves)}")

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

    date, ndates = write_shelves(book, rows, vals, datetime.now(sp.MSK))
    print(f"Готово: лист «{SHEET_DST}», колонка за {date} записана "
          f"({len(rows)} строк, дат в истории {ndates}).")


if __name__ == "__main__":
    main()
