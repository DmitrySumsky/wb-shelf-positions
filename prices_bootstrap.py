#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Разовая заводка листа «Цены» в ОТДЕЛЬНОЙ книге цен бренда.

v1.1.0 — 18.08.2026
  ШЕСТОЙ БРЕНД ORZAX БЕЗ ИСТОРИИ ДОНОРА — заводка книги цен кабинета ИП1.
  • `--no-history`: лист заводится пустым по датам. У Orzax свои карточки
    (перекупы того же товара), в доноре Health Form их нет ни одной, и без
    флага в книгу приехали бы 70 пустых колонок-дат «на всякий случай».

v1.0.0 — 10.08.2026
  ЦЕНЫ ЖИЛИ ЛИСТОМ ВНУТРИ КНИГИ ПОЛОК — просьба пользователя 10.08 разнести их
  по своим таблицам (по одной на бренд, ID — в `brands.json`).
  • книга-донор истории (Health Form) переезжает целиком листом `copy_to`:
    69 дат, ручные колонки «Капсул»/«Какую ставим», раскраска — всё сохраняется;
  • остальные книги собираются из листа «Полки» СВОЕГО бренда (состав свежее,
    чем у старого листа цен), а история цен переносится из книги-донора по
    артикулу: список конкурентов у брендов общий, цена карточки от бренда книги
    не зависит;
  • колонка «Капсул» тоже переносится по артикулу — это свойство карточки, а не
    книги. «Какую ставим» остаётся пустой: это решение по своему товару;
  • `--sync-rows` дописывает в конец существующего листа цен группы, которые
    появились в «Полках» позже. История прежних строк не двигается.

Скрипт разовый: в расписание не ставится, каждый запуск — осознанный. Уже
заведённый лист не трогает (кроме `--sync-rows`), чтобы повтор не затирал
работу человека.

Запуск:
    python prices_bootstrap.py --brand all --creds ключ.json
    python prices_bootstrap.py --brand "Health Form" --archive-source
    python prices_bootstrap.py --brand NATURI --sync-rows
"""

from __future__ import annotations

import argparse
from datetime import datetime

import gspread

import brand_prices as bp
import brand_shelves as bs
import shelf_positions as sp
import to_sheets as ts
import wb_config
from vexor_shelves import install_retries

SHEET_PRICES = "Цены"
SHEET_SHELVES = "Полки"

# Откуда берётся история цен и ручная колонка «Капсул» для новых книг.
HISTORY_FROM = "Health Form"

# Фиксированные колонки листа цен — как их завёл Артур в книге Health Form.
FIX_COLS = ["Товар", "Артикул конкурента", "Бренд конкурента",
            "~ Выручка конкурента", "Капсул", "Какую ставим, кошелек WB"]
FIX_WIDTHS = [330, 130, 150, 140, 70, 150]
DATE_WIDTH = 75

HEAD_BG = {"red": 0.85, "green": 0.85, "blue": 0.85}


def read_fixed(ws, want: list[str]) -> tuple[list[dict], dict]:
    """Лист → (строки с нужными полями, раскладка). Порядок строк сохраняется."""
    values = ws.get_all_values()
    if not values:
        raise SystemExit(f"Лист «{ws.title}» пуст")
    head = [str(x).strip() for x in values[0]]
    lay = bs.read_layout(head)
    idx = {}
    for title in want:
        idx[title] = next((i for i, cell in enumerate(head[:lay["nfix"]])
                           if cell.lower() == title.lower()), None)
    rows = []
    for raw in values[1:]:
        cells = [str(x).strip() for x in raw]
        if not any(cells[:lay["nfix"]]):
            continue
        rows.append({t: (cells[i] if i is not None and i < len(cells) else "")
                     for t, i in idx.items()})
    return rows, lay


def history_map(ws) -> tuple[list[str], dict[str, dict[str, str]], dict[str, str]]:
    """Лист-донор → (даты в порядке листа, {артикул: {дата: значение}}, {артикул: капсул}).

    Дубли артикула в доноре разрешаются в пользу ПЕРВОЙ строки: строки одного
    артикула в разных группах — это одна и та же карточка WB с одной ценой.
    """
    values = ws.get_all_values()
    head = [str(x).strip() for x in values[0]]
    lay = bs.read_layout(head)
    dates = [d for d, _ in sorted(lay["date_at"].items(), key=lambda kv: kv[1])]
    caps_at = next((i for i, cell in enumerate(head[:lay["nfix"]])
                    if cell.lower() == "капсул"), None)

    hist: dict[str, dict[str, str]] = {}
    caps: dict[str, str] = {}
    for raw in values[1:]:
        cells = [str(x).strip() for x in raw]
        if not any(cells[:lay["nfix"]]):
            continue
        art = cells[lay["art"]] if len(cells) > lay["art"] else ""
        if not art.isdigit() or art in hist:
            continue
        hist[art] = {d: (cells[i] if i < len(cells) else "")
                     for d, i in lay["date_at"].items()}
        if caps_at is not None and caps_at < len(cells):
            caps[art] = cells[caps_at]
    return dates, hist, caps


def build_grid(shelf_rows: list[dict], dates: list[str],
               hist: dict[str, dict[str, str]], caps: dict[str, str]) -> list[list[str]]:
    """Строки «Полок» + история донора → готовая таблица листа цен (со шапкой)."""
    grid = [FIX_COLS + list(dates)]
    for r in shelf_rows:
        art = r.get("Артикул конкурента", "")
        h = hist.get(art, {}) if art.isdigit() else {}
        grid.append([
            r.get("Товар", ""),
            art,
            r.get("Бренд конкурента", ""),
            r.get("~ Выручка конкурента", ""),
            caps.get(art, "") if art.isdigit() else "",
            "",                                   # «Какую ставим» ведёт человек
        ] + [h.get(d, "") for d in dates])
    return grid


def drop_default_sheet(book) -> None:
    """Снести пустой «Лист1»/«Sheet1», который Google создаёт вместе с книгой."""
    for ws in book.worksheets():
        if ws.title in ("Лист1", "Sheet1") and len(book.worksheets()) > 1:
            if not any(any(c.strip() for c in row) for row in ws.get_all_values()):
                book.del_worksheet(ws)
                sp.log(f"  удалён пустой лист «{ws.title}»")


def format_sheet(book, ws, ndates: int, nrows: int) -> None:
    """Шапка, заморозка, ширины, формат чисел, градиент по блокам."""
    nfix = len(FIX_COLS)
    reqs = [
        {"updateSheetProperties": {
            "properties": {"sheetId": ws.id,
                           "gridProperties": {"frozenRowCount": 1,
                                              "frozenColumnCount": 3}},
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}},
        {"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1},
            "cell": {"userEnteredFormat": {
                "backgroundColor": HEAD_BG,
                "textFormat": {"bold": True},
                "horizontalAlignment": "CENTER",
                "wrapStrategy": "WRAP"}},
            "fields": "userEnteredFormat(backgroundColor,textFormat,"
                      "horizontalAlignment,wrapStrategy)"}},
    ]
    # Книга без истории (ORZAX, v1.1.0) заводится вообще без колонок-дат: пустой
    # диапазон Sheets не принимает, а первую колонку всё равно вставит прогон.
    if ndates:
        reqs += [
            {"repeatCell": {
                "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": nrows,
                          "startColumnIndex": nfix, "endColumnIndex": nfix + ndates},
                "cell": {"userEnteredFormat": {
                    "horizontalAlignment": "CENTER",
                    "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}},
                "fields": "userEnteredFormat(horizontalAlignment,numberFormat)"}},
            {"updateDimensionProperties": {
                "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                          "startIndex": nfix, "endIndex": nfix + ndates},
                "properties": {"pixelSize": DATE_WIDTH}, "fields": "pixelSize"}},
        ]
    for i, width in enumerate(FIX_WIDTHS):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": i, "endIndex": i + 1},
            "properties": {"pixelSize": width}, "fields": "pixelSize"}})
    book.batch_update({"requests": reqs})


def write_grid(ws, grid: list[list[str]]) -> None:
    """Залить таблицу кусками по 300 строк: одна простыня на 90 тыс. ячеек не проходит."""
    width = max(len(r) for r in grid)
    body = [row + [""] * (width - len(row)) for row in grid]
    last_col = bs.col_letter(width)
    for i in range(0, len(body), 300):
        chunk = body[i:i + 300]
        ws.update(values=chunk,
                  range_name=f"A{i + 1}:{last_col}{i + len(chunk)}",
                  value_input_option="RAW")     # даты шапки не должны стать датами


def migrate_donor(client, brand: str, archive_source: bool) -> str:
    """Книга-донор: лист «Цены» переезжает из книги полок в книгу цен целиком."""
    src_book = client.open_by_key(bs.BRANDS[brand]["sheet_id"])
    dst_book = client.open_by_key(wb_config.prices_book(brand))
    src = src_book.worksheet(SHEET_PRICES)

    res = src.copy_to(dst_book.id)
    new = dst_book.get_worksheet_by_id(res["sheetId"])
    new.update_title(SHEET_PRICES)
    drop_default_sheet(dst_book)
    sp.log(f"  лист перенесён целиком: {new.row_count}×{new.col_count}")

    if archive_source:
        stamp = datetime.now(sp.MSK).strftime("%d.%m")
        src.update_title(f"{SHEET_PRICES} (архив до {stamp})")
        sp.log(f"  исходный лист переименован в «{src.title}»")
    return f"{brand}: лист «{SHEET_PRICES}» перенесён в свою книгу"


def build_book(client, brand: str, dates: list[str],
               hist: dict[str, dict[str, str]], caps: dict[str, str]) -> str:
    """Новая книга цен: строки из «Полок» бренда + история из донора."""
    shelf_book = client.open_by_key(bs.BRANDS[brand]["sheet_id"])
    shelf_rows, _ = read_fixed(shelf_book.worksheet(SHEET_SHELVES), FIX_COLS[:4])
    grid = build_grid(shelf_rows, dates, hist, caps)

    book = client.open_by_key(wb_config.prices_book(brand))
    ws = ts.ensure_ws(book, SHEET_PRICES,
                      rows=len(grid) + 50, cols=len(FIX_COLS) + len(dates) + 10)
    write_grid(ws, grid)
    format_sheet(book, ws, len(dates), len(grid))
    drop_default_sheet(book)

    lay = bs.read_layout(grid[0])
    blocks = bp.read_blocks(grid, lay)
    nrules = bp.rebuild_gradients(book, ws, lay, blocks, len(dates))
    filled = sum(1 for r in grid[1:] if r[1].isdigit() and r[1] in hist)
    return (f"{brand}: лист заведён — строк {len(grid) - 1}, дат {len(dates)}, "
            f"история подхвачена у {filled} карточек, блоков {len(blocks)}, "
            f"правил градиента {nrules}")


def sync_rows(client, brand: str) -> str:
    """Дописать в конец листа цен группы, которых там ещё нет (появились в «Полках»)."""
    shelf_book = client.open_by_key(bs.BRANDS[brand]["sheet_id"])
    shelf_rows, _ = read_fixed(shelf_book.worksheet(SHEET_SHELVES), FIX_COLS[:4])
    book = client.open_by_key(wb_config.prices_book(brand))
    ws = book.worksheet(SHEET_PRICES)
    price_rows, lay = read_fixed(ws, FIX_COLS[:4])

    have = {r["Товар"] for r in price_rows if r["Товар"]}
    add: list[dict] = []
    for r in shelf_rows:
        if r["Товар"] and r["Товар"] not in have:
            add.append(r)
    if not add:
        return f"{brand}: новых групп нет"

    groups = sorted({r["Товар"] for r in add})
    start = len(ws.get_all_values()) + 2          # пустая строка-разделитель
    body = [[r["Товар"], r["Артикул конкурента"], r["Бренд конкурента"],
             r["~ Выручка конкурента"], "", ""] for r in add]
    ws.update(values=body,
              range_name=f"A{start}:{bs.col_letter(len(FIX_COLS))}{start + len(body) - 1}",
              value_input_option="RAW")
    return (f"{brand}: дописано групп {len(groups)}, строк {len(body)} "
            f"({', '.join(groups[:5])}{'…' if len(groups) > 5 else ''})")


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Разовая заводка листа «Цены» в книге цен бренда")
    ap.add_argument("--brand", default="all")
    ap.add_argument("--creds", default=None)
    ap.add_argument("--sync-rows", action="store_true",
                    help="дописать в существующий лист недостающие группы")
    ap.add_argument("--archive-source", action="store_true",
                    help="переименовать исходный лист «Цены» в книге полок")
    ap.add_argument("--no-history", action="store_true",
                    help="завести лист без истории донора (карточек бренда в нём нет)")
    ap.add_argument("--force", action="store_true",
                    help="перезаписать уже заведённый лист (по умолчанию пропускаю)")
    args = ap.parse_args()

    brands = (wb_config.books_with_prices() if args.brand.lower() == "all"
              else [wb_config.brand_cfg(args.brand)[0]])

    install_retries()
    client = ts.get_client(args.creds)

    if args.sync_rows:
        for brand in brands:
            print("Готово:", sync_rows(client, brand))
        return

    # Историю читаем из книги-донора один раз: она общая для всех новых книг.
    if args.no_history:
        dates, hist, caps = [], {}, {}
        sp.log("Заводка без истории: колонок-дат не будет, первую допишет прогон цен")
    else:
        donor_book = client.open_by_key(bs.BRANDS[HISTORY_FROM]["sheet_id"])
        try:
            donor_ws = donor_book.worksheet(SHEET_PRICES)
        except gspread.WorksheetNotFound:               # донор уже переехал
            donor_ws = client.open_by_key(
                wb_config.prices_book(HISTORY_FROM)).worksheet(SHEET_PRICES)
        dates, hist, caps = history_map(donor_ws)
        sp.log(f"Донор истории — {HISTORY_FROM}: дат {len(dates)} "
               f"({dates[0]} … {dates[-1]}), карточек {len(hist)}, "
               f"из них с «Капсул» {sum(1 for v in caps.values() if v)}")

    results = []
    for brand in brands:
        sp.log(f"=== {brand} ===")
        book = client.open_by_key(wb_config.prices_book(brand))
        titles = [w.title for w in book.worksheets()]
        if SHEET_PRICES in titles and not args.force:
            results.append(f"{brand}: лист «{SHEET_PRICES}» уже есть — пропускаю "
                           "(--force перезапишет, --sync-rows дополнит)")
            continue
        if brand == HISTORY_FROM:
            results.append(migrate_donor(client, brand, args.archive_source))
        else:
            results.append(build_book(client, brand, dates, hist, caps))

    for line in results:
        print("Готово:", line)


if __name__ == "__main__":
    main()
