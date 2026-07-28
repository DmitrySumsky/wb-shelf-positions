#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Запись позиций в полках в Google-таблицу.

Источник списков — таблица «Аналитика цен WB», лист «Аналитика цен (wb кошелек)»:
колонка A — название товара (оно же группа), B — артикул, C — бренд. Строки одной
группы идут подряд. Наши бренды задаются в OUR_BRANDS, остальные в группе считаются
конкурентами. Позиции считаются ТОЛЬКО внутри группы.

Листы результата:
  «Матрица»  — строка = полка конкурента, на каждую дату блок из 4 колонок (наши
               бренды) с датой над ним. Свежая дата — слева, прежние съезжают вправо.
  «Сводка»   — строка = наш артикул: в скольких полках найден, лучшая/средняя/худшая.
  «История»  — строка = пара (наш артикул × конкурент), столбец = дата.

Доступ — сервисный аккаунт. Таблицу-приёмник создаёт человек и расшаривает на
e-mail SA как «Редактор»: у сервисного аккаунта вне Workspace нет своей квоты Drive.

Ключ: --creds путь | GCP_SA_JSON (содержимое) | GOOGLE_APPLICATION_CREDENTIALS
ID таблиц: --sheet-id / SHEET_ID (приёмник), --source-id / SOURCE_ID (списки)
"""

from __future__ import annotations

import argparse
import json
import os
import tempfile
from datetime import datetime, timedelta, timezone

import gspread

MSK = timezone(timedelta(hours=3))

SHEET_MATRIX = "Матрица"
SHEET_SUMMARY = "Сводка"
SHEET_HISTORY = "История"

SOURCE_TAB = "Аналитика цен (wb кошелек)"
OUR_BRANDS = ["NATURI", "SUNSHINE", "Health Form", "4ME"]

# Пороговая раскраска. Красим статически при перезаписи, а не правилами условного
# формата: сотни правил на большом листе заметно тормозят таблицу.
COLOR_BANDS = [
    (10, {"red": 0.78, "green": 0.91, "blue": 0.79}),    # 1–10    зелёный
    (30, {"red": 0.89, "green": 0.96, "blue": 0.81}),    # 11–30   салатовый
    (100, {"red": 1.00, "green": 0.95, "blue": 0.75}),   # 31–100  жёлтый
    (300, {"red": 0.99, "green": 0.87, "blue": 0.75}),   # 101–300 оранжевый
]
COLOR_DEEP = {"red": 0.96, "green": 0.80, "blue": 0.80}  # глубже 300
COLOR_NONE = {"red": 0.93, "green": 0.93, "blue": 0.93}  # нет в полке


def band_color(pos):
    if pos is None:
        return COLOR_NONE
    for limit, color in COLOR_BANDS:
        if pos <= limit:
            return color
    return COLOR_DEEP


# --------------------------------------------------------------------------- доступ

def get_client(creds_path: str | None):
    if creds_path:
        return gspread.service_account(filename=creds_path)
    raw = os.environ.get("GCP_SA_JSON")
    if raw:
        with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as f:
            f.write(raw)
            tmp = f.name
        return gspread.service_account(filename=tmp)
    env = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if env:
        return gspread.service_account(filename=env)
    raise SystemExit("Нет ключа сервисного аккаунта: --creds, GCP_SA_JSON или GOOGLE_APPLICATION_CREDENTIALS")


def ensure_ws(book, title: str, rows: int, cols: int):
    try:
        ws = book.worksheet(title)
    except gspread.WorksheetNotFound:
        return book.add_worksheet(title=title, rows=rows, cols=cols)
    # Лист нового образца — 26 колонок / 1000 строк. Расширяем ДО записи, иначе
    # batchUpdate отвечает 400 «beyond the last column».
    if ws.col_count < cols or ws.row_count < rows:
        ws.resize(rows=max(rows, ws.row_count), cols=max(cols, ws.col_count))
    return ws


# --------------------------------------------------------------- чтение списков

def read_groups(client, source_id: str, tab: str = SOURCE_TAB,
                our_brands: list[str] | None = None) -> tuple[list[dict], list[str]]:
    """Товарные группы из таблицы цен. Возвращает (группы, список пропущенных строк)."""
    our_brands = our_brands or OUR_BRANDS
    our_norm = {b.strip().lower(): b for b in our_brands}

    ws = client.open_by_key(source_id).worksheet(tab)
    values = ws.get_all_values()[1:]

    groups: dict[str, dict] = {}
    order: list[str] = []
    skipped: list[str] = []

    for row in values:
        product = (row[0] if len(row) > 0 else "").strip()
        nm_raw = (row[1] if len(row) > 1 else "").strip()
        brand = (row[2] if len(row) > 2 else "").strip()
        if not product:
            continue
        if product not in groups:
            groups[product] = {"product": product, "ours": [], "competitors": [],
                               "brand_of": {}}
            order.append(product)
        if not nm_raw.isdigit():
            # В таблице встречается «отсутствует» и пустые ячейки — товара у бренда нет.
            if nm_raw or brand:
                skipped.append(f"{product} / {brand or '—'} / {nm_raw or 'пусто'}")
            continue
        nm = int(nm_raw)
        g = groups[product]
        g["brand_of"][nm] = brand
        if brand.strip().lower() in our_norm:
            g["ours"].append(nm)
        else:
            g["competitors"].append(nm)

    result = []
    for name in order:
        g = groups[name]
        if g["ours"] and g["competitors"]:
            result.append(g)
        else:
            skipped.append(f"{name}: пропущена целиком "
                           f"(наших {len(g['ours'])}, конкурентов {len(g['competitors'])})")
    return result, skipped


def brand_of(snap: dict, nm: int) -> str:
    """Бренд из исходной таблицы, если есть; иначе из карточки WB."""
    for g in snap.get("groups", []):
        b = (g.get("brand_of") or {}).get(str(nm)) or (g.get("brand_of") or {}).get(nm)
        if b:
            return b
    return snap["cards"].get(str(nm), {}).get("brand", "")


def iter_pairs(snap: dict):
    for g in snap["groups"]:
        for our in g["ours"]:
            for comp in g["competitors"]:
                yield g["product"], our, comp


# ------------------------------------------------------------------------ листы

MATRIX_FIX = ["Товар", "Артикул конкурента", "Бренд конкурента", "Товаров в полке"]
MATRIX_HEAD_ROWS = 2      # строка 1 — даты, строка 2 — названия колонок


def build_entries(snap: dict, our_brands: list[str]) -> list[dict]:
    """Замер в виде строк «полка конкурента → позиции наших брендов»."""
    failed = set(snap.get("failed_shelves", []))
    missing = set(snap.get("missing_shelves", []))
    entries = []
    for g in snap["groups"]:
        our_by_brand = {}
        for our in g["ours"]:
            our_by_brand[brand_of(snap, our).strip().lower()] = our
        for comp in g["competitors"]:
            # Три разных случая, которые нельзя смешивать: полка не проверена
            # из-за сбоя, карточки конкурента больше нет на WB, и «нас там нет».
            flag = ("ошибка сбора" if str(comp) in failed else
                    "нет карточки" if str(comp) in missing else None)
            cells = []
            for br in our_brands:
                our = our_by_brand.get(br.strip().lower())
                if not our:
                    cells.append("")
                elif flag:
                    cells.append(flag)
                else:
                    pos = snap["positions"].get(str(our), {}).get(str(comp))
                    cells.append(pos if pos else "—")
            entries.append({
                "product": g["product"],
                "comp": str(comp),
                "brand": brand_of(snap, comp),
                "total": flag or snap["shelves"].get(str(comp), {}).get("total", ""),
                "cells": cells,
            })
    return entries


def write_matrix(book, snap: dict, our_brands: list[str]) -> str:
    """
    Матрица с накоплением по датам.

        A..D  — Товар | Артикул конкурента | Бренд конкурента | Товаров в полке
        строка 1 — дата над блоком, строка 2 — названия наших брендов
        каждый прогон вставляет блок из len(our_brands) колонок СЛЕВА, сразу
        за фиксированными колонками, и сдвигает прежние даты вправо

    Лист НЕ перезаписывается: новые конкуренты дописываются строками вниз,
    новая дата — блоком слева (самое свежее всегда под рукой, листать вправо
    не надо). Повторный прогон в тот же день обновляет свой блок на месте.
    """
    nfix = len(MATRIX_FIX)
    nb = len(our_brands)
    date = datetime.fromisoformat(snap["snapshot_at"]).strftime("%d.%m")
    entries = build_entries(snap, our_brands)

    ws = ensure_ws(book, SHEET_MATRIX,
                   rows=len(entries) + MATRIX_HEAD_ROWS + 50, cols=nfix + nb * 10)
    values = ws.get_all_values()

    # Шапка нового образца стоит во ВТОРОЙ строке (в первой — даты). Если там не
    # она — лист либо пуст, либо остался от прежней однодневной раскладки.
    fresh = (len(values) < MATRIX_HEAD_ROWS
             or not values[1]
             or values[1][0].strip() != MATRIX_FIX[0])
    if fresh:
        ws.clear()
        ws.update(values=[MATRIX_FIX + our_brands], range_name="A2")
        values = [[""] * nfix, MATRIX_FIX + our_brands]

    row1 = values[0] if values else []
    row2 = values[1] if len(values) > 1 else MATRIX_FIX + our_brands
    body = values[MATRIX_HEAD_ROWS:]

    # Существующие блоки дат: подпись стоит в первой колонке блока.
    n_blocks = max(0, (max(len(row1), len(row2)) - nfix + nb - 1) // nb)
    date_col0 = None
    for b in range(n_blocks):
        i = nfix + b * nb
        if i < len(row1) and row1[i].strip() == date:
            date_col0 = i
            break

    date_is_new = date_col0 is None
    if date_is_new:
        date_col0 = nfix

    # Существующие строки по ключу (товар, артикул конкурента).
    row_of = {}
    for i, r in enumerate(body):
        key = ((r[0].strip() if len(r) > 0 else ""), (r[1].strip() if len(r) > 1 else ""))
        if key[1]:
            row_of[key] = i + MATRIX_HEAD_ROWS + 1

    new_fixed = []
    next_row = len(body) + MATRIX_HEAD_ROWS + 1
    for e in entries:
        key = (e["product"], e["comp"])
        if key not in row_of:
            row_of[key] = next_row
            new_fixed.append(e)
            next_row += 1

    # Только строки: ширину добирает сама вставка блока ниже. Резать лист по
    # колонкам тут нельзя — gspread держит устаревший col_count и снёс бы
    # правый край с прежними датами.
    if ws.row_count < next_row + 30:
        ws.resize(rows=next_row + 30, cols=ws.col_count)

    # Новой даты на листе нет — врезаем под неё блок колонок сразу за
    # фиксированной частью. Google сам сдвигает вправо и данные, и мерджи
    # прежних дат, поэтому пересобирать старые блоки не нужно.
    if date_is_new:
        book.batch_update({"requests": [{"insertDimension": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": nfix, "endIndex": nfix + nb},
            # Наследуем оформление от фиксированных колонок слева, а не от
            # вчерашнего блока справа: иначе в новый блок переедет вчерашняя
            # заливка и пустые ячейки остались бы крашеными.
            "inheritFromBefore": True}}]})

    updates = []

    # Заголовок блока: дата и названия брендов.
    updates.append({"range": gspread.utils.rowcol_to_a1(1, date_col0 + 1), "values": [[date]]})
    updates.append({"range": gspread.utils.rowcol_to_a1(2, date_col0 + 1), "values": [our_brands]})

    # Фиксированная часть новых строк.
    if new_fixed:
        first_new = row_of[(new_fixed[0]["product"], new_fixed[0]["comp"])]
        updates.append({
            "range": f"{gspread.utils.rowcol_to_a1(first_new, 1)}:"
                     f"{gspread.utils.rowcol_to_a1(first_new + len(new_fixed) - 1, nfix)}",
            "values": [[e["product"], e["comp"], e["brand"], e["total"]] for e in new_fixed]})

    used = sorted(row_of[(e["product"], e["comp"])] for e in entries)
    first, last = used[0], used[-1]

    # «Товаров в полке» — колонка одна на все даты, держим актуальной.
    totals = [[""] for _ in range(last - first + 1)]
    block = [[""] * nb for _ in range(last - first + 1)]
    for e in entries:
        r = row_of[(e["product"], e["comp"])] - first
        totals[r] = [e["total"]]
        block[r] = e["cells"]
    updates.append({
        "range": f"{gspread.utils.rowcol_to_a1(first, nfix)}:"
                 f"{gspread.utils.rowcol_to_a1(last, nfix)}",
        "values": totals})
    updates.append({
        "range": f"{gspread.utils.rowcol_to_a1(first, date_col0 + 1)}:"
                 f"{gspread.utils.rowcol_to_a1(last, date_col0 + nb)}",
        "values": block})

    ws.batch_update(updates, value_input_option="USER_ENTERED")

    # Оформление: шапка, заморозка, объединение даты над блоком, раскраска.
    reqs = [
        {"updateSheetProperties": {
            "properties": {"sheetId": ws.id,
                           "gridProperties": {"frozenRowCount": MATRIX_HEAD_ROWS,
                                              "frozenColumnCount": nfix}},
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}},
        {"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": MATRIX_HEAD_ROWS},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True},
                                           "wrapStrategy": "WRAP",
                                           "horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat(textFormat,wrapStrategy,horizontalAlignment)"}},
    ]
    # Свежевставленный блок гасим до белого: раскрашиваем ниже только те ячейки,
    # где сегодня есть значение, а унаследованная заливка иначе осталась бы висеть.
    if date_is_new:
        reqs.append({"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": MATRIX_HEAD_ROWS,
                      "endRowIndex": max(next_row, last),
                      "startColumnIndex": date_col0, "endColumnIndex": date_col0 + nb},
            "cell": {"userEnteredFormat": {"backgroundColor": {"red": 1, "green": 1, "blue": 1},
                                           "horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment)"}})

    # Мердж даты над блоком. Блок целиком лежит правее замороженных колонок,
    # поэтому границу frozenColumnCount не пересекает (иначе Google отдаёт 400).
    if nb > 1:
        reqs.append({"mergeCells": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": date_col0, "endColumnIndex": date_col0 + nb},
            "mergeType": "MERGE_ALL"}})

    for e in entries:
        r = row_of[(e["product"], e["comp"])]
        for bi, val in enumerate(e["cells"]):
            if val == "":
                continue
            pos = val if isinstance(val, int) else None
            color = band_color(pos) if pos else COLOR_NONE
            reqs.append({"repeatCell": {
                "range": {"sheetId": ws.id, "startRowIndex": r - 1, "endRowIndex": r,
                          "startColumnIndex": date_col0 + bi,
                          "endColumnIndex": date_col0 + bi + 1},
                "cell": {"userEnteredFormat": {"backgroundColor": color,
                                               "horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat(backgroundColor,horizontalAlignment)"}})
    for i in range(0, len(reqs), 500):
        book.batch_update({"requests": reqs[i:i + 500]})
    return date


def write_summary(book, snap: dict) -> None:
    """Строка = наш артикул: где стоим по всей своей группе."""
    header = ["Товар", "Наш бренд", "Наш артикул", "Полок проверено", "Найдено",
              "Лучшая", "Средняя", "Худшая", "Полок не проверено"]
    rows = [header]
    skip = set(snap.get("failed_shelves", [])) | set(snap.get("missing_shelves", []))
    for g in snap["groups"]:
        # Непроверенные полки из знаменателя исключаем — иначе «найдено 4 из 6»
        # читалось бы как «в двух полках нас нет», хотя мы их просто не проверили.
        ok_comps = [c for c in g["competitors"] if str(c) not in skip]
        broken = len(g["competitors"]) - len(ok_comps)
        for our in g["ours"]:
            hits = [p for p in (snap["positions"].get(str(our), {}).get(str(c))
                                for c in ok_comps) if p]
            rows.append([
                g["product"], brand_of(snap, our), our,
                len(ok_comps), len(hits),
                min(hits) if hits else "—",
                round(sum(hits) / len(hits)) if hits else "—",
                max(hits) if hits else "—",
                broken if broken else "",
            ])

    ws = ensure_ws(book, SHEET_SUMMARY, rows=len(rows) + 20, cols=len(header) + 2)
    ws.clear()
    ws.update(values=rows, range_name="A1", value_input_option="USER_ENTERED")
    book.batch_update({"requests": [
        {"updateSheetProperties": {
            "properties": {"sheetId": ws.id,
                           "gridProperties": {"frozenRowCount": 1, "frozenColumnCount": 3}},
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}},
        {"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True},
                                           "wrapStrategy": "WRAP"}},
            "fields": "userEnteredFormat(textFormat,wrapStrategy)"}},
    ]})


FIXED_COLS = ["Товар", "Наш бренд", "Наш артикул", "Артикул конкурента", "Бренд конкурента"]


def append_history(book, snap: dict) -> int:
    """
    История: строка = пара (наш артикул × конкурент), столбец = дата.

    Плоский журнал по строке на замер проще, но при 136 наших × ~5 конкурентов это
    ~740 строк в сутки — 2,4 млн ячеек за год при лимите таблицы в 10 млн.
    Раскладка по датам держит объём в ~0,3 млн.
    """
    date = snap["snapshot_at"][:10]
    nfix = len(FIXED_COLS)
    pairs = list(iter_pairs(snap))

    ws = ensure_ws(book, SHEET_HISTORY, rows=len(pairs) + 100, cols=nfix + 40)
    values = ws.get_all_values()
    if not values or not values[0] or not values[0][0]:
        ws.update(values=[FIXED_COLS], range_name="A1")
        ws.freeze(rows=1, cols=nfix)
        values = [FIXED_COLS]

    header = list(values[0])
    body = values[1:]

    row_of = {}
    for i, row in enumerate(body):
        if len(row) >= 4 and row[2].strip() and row[3].strip():
            row_of[(row[2].strip(), row[3].strip())] = i + 2

    date_is_new = date not in header
    date_col = (len(header) + 1) if date_is_new else (header.index(date) + 1)

    new_rows = []
    next_row = len(body) + 2
    for product, our, comp in pairs:
        key = (str(our), str(comp))
        if key not in row_of:
            row_of[key] = next_row
            new_rows.append([product, brand_of(snap, our), our, comp, brand_of(snap, comp)])
            next_row += 1

    if ws.row_count < next_row + 50 or ws.col_count < date_col + 5:
        ws.resize(rows=max(next_row + 50, ws.row_count),
                  cols=max(date_col + 5, ws.col_count))

    if date_is_new:
        ws.update(values=[[date]], range_name=gspread.utils.rowcol_to_a1(1, date_col))
    if new_rows:
        ws.update(values=new_rows,
                  range_name=gspread.utils.rowcol_to_a1(len(body) + 2, 1),
                  value_input_option="USER_ENTERED")

    used = sorted(row_of[(str(o), str(c))] for _, o, c in pairs)
    first, last = used[0], used[-1]
    failed = set(snap.get("failed_shelves", []))
    missing = set(snap.get("missing_shelves", []))
    column = [[""] for _ in range(last - first + 1)]
    for _, our, comp in pairs:
        if str(comp) in failed:
            value = "ошибка сбора"    # полку не проверили, а не «нас там нет»
        elif str(comp) in missing:
            value = "нет карточки"    # артикул конкурента удалён с WB
        else:
            pos = snap["positions"].get(str(our), {}).get(str(comp))
            value = pos if pos else "—"
        column[row_of[(str(our), str(comp))] - first] = [value]

    ws.update(values=column,
              range_name=f"{gspread.utils.rowcol_to_a1(first, date_col)}:"
                         f"{gspread.utils.rowcol_to_a1(last, date_col)}",
              value_input_option="USER_ENTERED")
    return len(pairs)


# ------------------------------------------------------------------------- main

def main() -> None:
    ap = argparse.ArgumentParser(description="Позиции в полках конкурентов → Google-таблица")
    ap.add_argument("--sheet-id", default=os.environ.get("SHEET_ID"))
    ap.add_argument("--source-id", default=os.environ.get("SOURCE_ID"))
    ap.add_argument("--source-tab", default=SOURCE_TAB)
    ap.add_argument("--creds", default=None)
    ap.add_argument("--snapshot", default=None, help="готовый snapshot.json вместо сбора")
    ap.add_argument("--save-snapshot", default="snapshot.json")
    ap.add_argument("--dest", type=int, default=None)
    ap.add_argument("--workers", type=int, default=6)
    ap.add_argument("--dry-run", action="store_true", help="собрать и сохранить, в таблицу не писать")
    args = ap.parse_args()

    if not args.sheet_id:
        raise SystemExit("Не задан --sheet-id (или SHEET_ID)")

    client = get_client(args.creds)

    if args.snapshot:
        with open(args.snapshot, encoding="utf-8") as f:
            snap = json.load(f)
    else:
        if not args.source_id:
            raise SystemExit("Не задан --source-id (или SOURCE_ID) — откуда брать артикулы")
        import shelf_positions as sp
        groups, skipped = read_groups(client, args.source_id, args.source_tab)
        print(f"Групп к обходу: {len(groups)}; строк пропущено: {len(skipped)}")
        for s in skipped[:15]:
            print("  пропуск:", s)
        if len(skipped) > 15:
            print(f"  … и ещё {len(skipped) - 15}")
        snap = sp.run_groups(groups, dest=args.dest or sp.DEST_MOSCOW,
                             workers=args.workers)
        snap["skipped"] = skipped
        if args.save_snapshot:
            with open(args.save_snapshot, "w", encoding="utf-8") as f:
                json.dump(snap, f, ensure_ascii=False, indent=2)
            print(f"Снапшот сохранён: {args.save_snapshot}")

    if args.dry_run:
        print("dry-run: в таблицу не пишу.")
        return

    book = client.open_by_key(args.sheet_id)
    date = write_matrix(book, snap, OUR_BRANDS)
    write_summary(book, snap)
    n = append_history(book, snap)
    print(f"Готово: в «Матрицу» добавлен блок за {date}, «Сводка» обновлена, "
          f"в «Историю» записано {n} пар ({snap['snapshot_at']}).")


if __name__ == "__main__":
    main()
