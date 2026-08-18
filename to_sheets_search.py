#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ЭКСПЕРИМЕНТАЛЬНЫЙ параллельный прогон: конкуренты берутся из поисковой выдачи WB,
а не из ручного списка в таблице цен.

Отличий от боевого `to_sheets.py` ровно два:
  * источник конкурентов — лист «Запросы (поиск)» в таблице-приёмнике
    (слово, глубина топа, наши артикулы), а не таблица «Аналитика цен WB»;
  * результат пишется в листы с суффиксом «(поиск)».

Боевые листы «Матрица» / «Сводка» / «История» и боевой `snapshot.json` не
затрагиваются: имена листов подменяются только в памяти этого процесса, снапшот
кладётся в `snapshot_search.json`. Логику обхода полок, накопление дат и раскраску
переиспользуем как есть — иначе два контура разъедутся по смыслу.

Запуск:
    python to_sheets_search.py --creds ключ.json --sheet-id <ID>
    python to_sheets_search.py --queries "магний,коллаген" --top-n 150 --dry-run
"""

from __future__ import annotations

import argparse
import json
import os

import gspread

import to_sheets as ts
import search_groups as sg
import shelf_positions as sp

# Подмена имён листов. Константы модуля читаются внутри его функций как глобальные,
# поэтому переопределение здесь уводит запись в параллельные листы, а сам файл
# to_sheets.py остаётся нетронутым — боевой прогон запускается отдельным процессом.
SUFFIX = " (поиск)"
ts.SHEET_MATRIX = "Матрица" + SUFFIX
ts.SHEET_SUMMARY = "Сводка" + SUFFIX
ts.SHEET_HISTORY = "История" + SUFFIX

QUERIES_TAB = "Запросы" + SUFFIX
REPORT_TAB = "Отчёт" + SUFFIX

QUERIES_HEADER = ["Запрос", "Топ N конкурентов", "Наши артикулы (через запятую)",
                  "Фильтр по категории", "Включён"]
# Стартовый набор: слова, которыми покупатель ищет товар из наших же групп
# («Magnesium Chelate…», «Inositol…», «Collagen…», «Taurine…»). Дальше список
# правит менеджер прямо в таблице — код его не перезаписывает.
QUERIES_SEED = [
    ["магний", 100, "", "да", "да"],
    ["инозитол", 100, "", "да", "да"],
    ["коллаген", 100, "", "да", "да"],
    ["таурин", 100, "", "да", "да"],
    ["омега 3", 100, "", "да", "нет"],
    ["цинк пиколинат", 100, "", "да", "нет"],
]

NO_WORDS = {"нет", "no", "false", "0", "выкл", "off", "-"}


def _flag(value: str, default: bool = True) -> bool:
    v = (value or "").strip().lower()
    if not v:
        return default
    return v not in NO_WORDS


def read_queries(book) -> list[dict]:
    """
    Запросы из листа «Запросы (поиск)». Листа нет — заводим со стартовым набором
    и работаем по нему; дальше список живёт в таблице, код его не трогает.
    """
    try:
        ws = book.worksheet(QUERIES_TAB)
    except gspread.WorksheetNotFound:
        ws = book.add_worksheet(title=QUERIES_TAB, rows=50, cols=len(QUERIES_HEADER) + 2)
        ws.update(values=[QUERIES_HEADER] + QUERIES_SEED, range_name="A1",
                  value_input_option="USER_ENTERED")
        ws.freeze(rows=1)
        book.batch_update({"requests": [{"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True},
                                           "wrapStrategy": "WRAP"}},
            "fields": "userEnteredFormat(textFormat,wrapStrategy)"}}]})
        print(f"Создан лист «{QUERIES_TAB}» со стартовым набором запросов — "
              f"дальше правьте его прямо в таблице.")

    queries = []
    for row in ws.get_all_values()[1:]:
        cell = lambda i: (row[i].strip() if len(row) > i else "")
        query = cell(0)
        if not query or not _flag(cell(4)):
            continue
        top_raw = cell(1).replace(" ", "")
        our_raw = [x.strip() for x in cell(2).replace(";", ",").split(",")]
        queries.append({
            "query": query,
            "top_n": int(top_raw) if top_raw.isdigit() else 200,
            "our_nm": [int(x) for x in our_raw if x.isdigit()],
            "subject_filter": _flag(cell(3)),
        })
    return queries


def write_report(book, snap: dict) -> None:
    """Что именно собрали и что отбросили — иначе «топ-200» читается как «все»."""
    rows = [["Запрос", "Наших карточек", "Конкурентов (полок)", "Строк в выдаче"]]
    for g in snap.get("groups", []):
        rows.append([g["product"], len(g["ours"]), len(g["competitors"]),
                     len(g["ours"]) * len(g["competitors"])])
    rows.append([])
    rows.append([f"Замер: {snap['snapshot_at']}", f"Полок обойдено: {len(snap['competitors'])}",
                 f"Не отдалось: {len(snap.get('failed_shelves', []))}",
                 f"Нет карточки: {len(snap.get('missing_shelves', []))}"])
    rows.append([])
    rows.append(["Подробности сбора:"])
    for line in snap.get("search_report", []):
        rows.append([line])

    ws = ts.ensure_ws(book, REPORT_TAB, rows=len(rows) + 20, cols=6)
    ws.clear()
    ws.update(values=rows, range_name="A1", value_input_option="USER_ENTERED")
    book.batch_update({"requests": [{"repeatCell": {
        "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1},
        "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
        "fields": "userEnteredFormat(textFormat)"}}]})


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Позиции в полках конкурентов ИЗ ПОИСКОВОЙ ВЫДАЧИ → Google-таблица")
    ap.add_argument("--sheet-id", default=os.environ.get("SHEET_ID"))
    ap.add_argument("--creds", default=None)
    ap.add_argument("--queries", default=None,
                    help="запросы через запятую вместо листа «Запросы (поиск)»")
    ap.add_argument("--top-n", type=int, default=200, help="конкурентов на запрос (для --queries)")
    ap.add_argument("--max-shelves", type=int, default=600,
                    help="потолок уникальных полок за прогон, 0 = без ограничения")
    ap.add_argument("--snapshot", default=None, help="готовый снапшот вместо сбора")
    ap.add_argument("--save-snapshot", default="snapshot_search.json")
    ap.add_argument("--dest", type=int, default=None)
    ap.add_argument("--workers", type=int, default=6)
    ap.add_argument("--dry-run", action="store_true", help="собрать и сохранить, в таблицу не писать")
    args = ap.parse_args()

    client = None
    book = None
    if not args.dry_run or not args.queries:
        if not args.sheet_id:
            raise SystemExit("Не задан --sheet-id (или SHEET_ID)")
        client = ts.get_client(args.creds)
        book = client.open_by_key(args.sheet_id)

    if args.snapshot:
        with open(args.snapshot, encoding="utf-8") as f:
            snap = json.load(f)
    else:
        if args.queries:
            queries = [{"query": q.strip(), "top_n": args.top_n, "our_nm": [],
                        "subject_filter": True}
                       for q in args.queries.split(",") if q.strip()]
        else:
            queries = read_queries(book)
        if not queries:
            raise SystemExit("Нет включённых запросов — заполните лист «Запросы (поиск)».")

        groups, report = sg.build_groups(
            queries, ts.OUR_BRANDS,
            dest=args.dest or sp.DEST_MOSCOW,
            max_shelves=args.max_shelves or None)
        if not groups:
            raise SystemExit("Ни по одному запросу не собралось группы — см. лог выше.")

        snap = sp.run_groups(groups, dest=args.dest or sp.DEST_MOSCOW, workers=args.workers)
        snap["search_report"] = report
        snap["source_kind"] = "search"
        if args.save_snapshot:
            with open(args.save_snapshot, "w", encoding="utf-8") as f:
                json.dump(snap, f, ensure_ascii=False, indent=2)
            print(f"Снапшот сохранён: {args.save_snapshot}")

    if args.dry_run:
        print("dry-run: в таблицу не пишу.")
        return

    date = ts.write_matrix(book, snap, ts.OUR_BRANDS)
    ts.write_summary(book, snap)
    n = ts.append_history(book, snap)
    write_report(book, snap)
    print(f"Готово (ветка «поиск»): блок за {date} в «{ts.SHEET_MATRIX}», "
          f"«{ts.SHEET_SUMMARY}» обновлена, в «{ts.SHEET_HISTORY}» записано {n} пар.")


if __name__ == "__main__":
    main()
