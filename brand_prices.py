#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Цены по карточкам листа «Цены» ОТДЕЛЬНОЙ книги цен бренда.

v2.0.0 — 10.08.2026
  ЦЕНЫ ПЕРЕЕХАЛИ ИЗ КНИГИ ПОЛОК В СВОЮ КНИГУ (просьба пользователя 10.08):
  • книга берётся из `brands.json` (поле `prices_book`), а не из книги полок;
  • сбор включён у всех четырёх брендов, а не у одной Health Form;
  • артикулы всех книг снимаются ОДНИМ обходом WB (пересечение списков 93 %),
    дальше та же карта цен раскладывается по книгам — 4 книги стоят столько же
    сетевых запросов, сколько раньше одна;
  • падение одной книги не роняет остальные, в конце — сводка и код возврата.

v1.9.0 — 10.08.2026 · градиент держится на всей истории листа
v1.7.0 — 10.08.2026 · градиент цен внутри блока товара, свой на каждую дату
v1.5.0 — 07.08.2026 · первая версия: цены по карточкам листа «Цены»

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

Запуск: шагом в shelves.yml после полок, интрадей-прогоном prices.yml либо руками:
    python brand_prices.py --brand all --creds ключ.json
"""

from __future__ import annotations

import argparse
import json
import os
from datetime import datetime

import gspread

import brand_shelves as bs
import notify
import prices_to_sheets as pts
import shelf_positions as sp
import to_sheets as ts
import wb_config
from vexor_shelves import install_retries

SHEET_PRICES = "Цены"
KEEP_DATES = wb_config.KEEP_DATES

# Градиент «дёшево зелёное → дорого красное» внутри блока товара, СВОЙ на каждую
# дату (правило = блок × колонка): так цвет отвечает на вопрос «кто дешевле
# конкурентов В ЭТОТ день», а не «за все дни разом». Правила Артура от 07.08
# стояли на диапазоне блок × H:L целиком и новую колонку не захватывали — она
# вставляется левее и выпадала из их диапазона. Теперь правила пересобирает
# скрипт после каждой записи.
# Держим градиент на ВСЕЙ истории листа (10.08.2026, просьба Артура): правил
# получается блоки × даты — на боевом листе Health Form это 44 × 69 ≈ 3 000.
# Sheets их тянет, но если книга начнёт открываться медленно, лечится снижением
# этого числа: свежие даты покроются, глубина останется числами.
GRADIENT_DATES = wb_config.GRADIENT_DATES
GRADIENT_MIN = {"red": 0.34117648, "green": 0.73333335, "blue": 0.5411765}
GRADIENT_MAX = {"red": 0.9019608, "green": 0.4862745, "blue": 0.4509804}


def brands_with_prices() -> list[str]:
    """Бренды, у которых в `brands.json` заведена книга цен.

    Лист «Цены» ведёт человек, поэтому книга подключается осознанно — полем
    `prices` в настройках, а не «раз есть лист — пишем».
    """
    return wb_config.books_with_prices()

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


def read_blocks(values: list[list[str]], lay: dict) -> list[tuple[int, int]]:
    """Блоки товара как (первая строка, последняя строка), 1-based по листу.

    Блок — подряд идущие строки с одинаковым «Товар» (пустое имя наследует
    строку выше, как в «Полках»); пустая строка-разделитель блок закрывает.
    """
    it = lay["product"]
    blocks: list[tuple[int, int]] = []
    name, start, last = None, None, None
    for i, raw in enumerate(values[1:], start=2):
        cells = [str(x).strip() for x in raw]
        fixed = cells[:lay["nfix"]]
        cur = cells[it].strip() if len(cells) > it else ""
        if not any(fixed):
            if start:
                blocks.append((start, last))
            name, start, last = None, None, None
            continue
        if cur and cur != name:
            if start:
                blocks.append((start, last))
            name, start = cur, i
        elif start is None:
            name, start = cur or name, i
        last = i
    if start:
        blocks.append((start, last))
    return [b for b in blocks if b[1] > b[0]]      # блок из одной строки не красим


def rebuild_gradients(book, ws, lay: dict, blocks: list[tuple[int, int]],
                      ndates: int) -> int:
    """Правило-градиент на каждый блок × каждую свежую дату. Вернёт число правил.

    Прежние правила скрипта снимаются целиком: строки и даты переезжают, и
    подправить диапазон существующего правила дешевле только на бумаге. Правила
    человека вне зоны дат (на фиксированных колонках) не трогаются.
    """
    nfix, sid = lay["nfix"], ws.id
    meta = book.fetch_sheet_metadata(
        {"fields": "sheets(properties/sheetId,conditionalFormats)"})
    old = []
    for s in meta.get("sheets", []):
        if s.get("properties", {}).get("sheetId") != sid:
            continue
        old = s.get("conditionalFormats", []) or []
    drop = [i for i, rule in enumerate(old)
            if rule.get("gradientRule") and all(
                r.get("startColumnIndex", 0) >= nfix for r in rule.get("ranges", []))]

    reqs: list[dict] = []
    for i in reversed(drop):                       # с конца: индексы съезжают
        reqs.append({"deleteConditionalFormatRule": {"sheetId": sid, "index": i}})
    for col in range(nfix, nfix + min(ndates, GRADIENT_DATES)):
        for top, bottom in blocks:
            reqs.append({"addConditionalFormatRule": {"index": 0, "rule": {
                "ranges": [{"sheetId": sid,
                            "startRowIndex": top - 1, "endRowIndex": bottom,
                            "startColumnIndex": col, "endColumnIndex": col + 1}],
                "gradientRule": {
                    "minpoint": {"colorStyle": {"rgbColor": GRADIENT_MIN},
                                 "type": "MIN"},
                    "midpoint": {"colorStyle": {"rgbColor": pts.WHITE},
                                 "type": "PERCENTILE", "value": "50"},
                    "maxpoint": {"colorStyle": {"rgbColor": GRADIENT_MAX},
                                 "type": "MAX"}}}}})
    for i in range(0, len(reqs), 300):
        book.batch_update({"requests": reqs[i:i + 300]})
    return len(reqs) - len(drop)


def write_column(book, ws, lay: dict, values: list[list[str]], rows: list[dict],
                 prices: dict[int, int | str], when: datetime,
                 gradient: bool = True) -> tuple[str, int, int]:
    """Колонка цен за сегодня. Вернёт (дата, подешевело, подорожало).

    `gradient=False` — не пересобирать условное форматирование: правила уже
    покрывают сегодняшнюю колонку, а пересборка это ~3 000 правил и ~20 с на
    книгу. Интрадей-прогоны (каждые 2 часа) переписывают колонку НА МЕСТЕ, им
    пересобирать нечего; раз в сутки, когда колонка вставляется новая, градиент
    строится заново.
    """
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
            # Цвет числа целиком за градиентом блока (дёшево зелёное — дорого
            # красное): условное форматирование рисуется ПОВЕРХ ручного фона,
            # и раскраска «подешевело со вчера» под ним всё равно не видна.
            # Считаем движение только для итоговой строки прогона.
            old = prev.get(r["row"])
            if old is not None and old != val:
                if val < old:
                    down += 1
                else:
                    up += 1
            continue
        if val == pts.STATE_FAIL:
            continue                      # дырку замера не красим
        color = pts.COLOR_STATE           # «нет в наличии» и прочий текст — серым
        reqs.append({"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": r["row"] - 1,
                      "endRowIndex": r["row"],
                      "startColumnIndex": col - 1, "endColumnIndex": col},
            "cell": {"userEnteredFormat": {"backgroundColor": color}},
            "fields": "userEnteredFormat.backgroundColor"}})

    for i in range(0, len(reqs), 300):
        book.batch_update({"requests": reqs[i:i + 300]})

    if gradient:
        blocks = read_blocks(values, lay)
        nrules = rebuild_gradients(book, ws, lay, blocks, ndates)
        sp.log(f"Градиент по блокам: блоков {len(blocks)}, дат под градиентом "
               f"{min(ndates, GRADIENT_DATES)}, правил {nrules}")
    return today, down, up


def open_prices_ws(client, brand: str):
    """(книга цен, лист «Цены») бренда. Книга — из `brands.json`."""
    book = client.open_by_key(wb_config.prices_book(brand))
    try:
        return book, book.worksheet(SHEET_PRICES)
    except gspread.WorksheetNotFound:
        raise SystemExit(
            f"В книге цен бренда {brand} нет листа «{SHEET_PRICES}». "
            "Строки листа ведёт человек, сам его не завожу — лист заводится "
            "один раз скриптом prices_bootstrap.py")


def stat_of(prices: dict[int, int | str], nm_ids: list[int]) -> str:
    mine = [prices.get(nm, pts.STATE_FAIL) for nm in nm_ids]
    return (f"цен {sum(1 for v in mine if isinstance(v, int))}, "
            f"нет в наличии {sum(1 for v in mine if v == pts.STATE_NONE)}, "
            f"нет карточки {sum(1 for v in mine if v == pts.STATE_GONE)}, "
            f"ошибок сбора {sum(1 for v in mine if v == pts.STATE_FAIL)}")


def run_all(client, brands: list[str], args) -> tuple[list[str], list[str]]:
    """Один обход WB на все книги цен. Вернёт (итоги, бренды с ошибкой).

    Списки карточек у книг пересекаются на 93 % (список конкурентов общий, у
    книг различается только своя карточка), а цена карточки от бренда книги не
    зависит. Поэтому сначала читаем все листы, потом снимаем ОБЪЕДИНЕНИЕ
    артикулов одним проходом, и уже готовую карту раскладываем по книгам.
    """
    plans: list[dict] = []
    failed: list[str] = []
    for brand in brands:
        try:
            book, ws = open_prices_ws(client, brand)
            rows, lay, _ = read_price_rows(ws)
            nm_ids = sorted({int(r["nm"]) for r in rows if r["nm"]})
            sp.log(f"{brand}: лист «{SHEET_PRICES}» — строк {len(rows)}, без "
                   f"артикула {sum(1 for r in rows if not r['nm'])}, карточек "
                   f"{len(nm_ids)}, дат в истории {len(lay['date_at'])}")
            plans.append({"brand": brand, "book": book, "ws": ws, "nm_ids": nm_ids})
        except Exception as exc:                        # noqa: BLE001
            sp.log(f"ОШИБКА чтения книги цен {brand}: {exc.__class__.__name__}: {exc}")
            failed.append(brand)

    if not plans:
        return [], failed

    union = sorted({nm for p in plans for nm in p["nm_ids"]})
    sp.log(f"Снимаю цены: книг {len(plans)}, уникальных карточек {len(union)} "
           f"(сумма по книгам {sum(len(p['nm_ids']) for p in plans)})")
    prices = pts.fetch_prices(union, args.dest)
    sp.log("Итог сбора: " + stat_of(prices, union))

    results: list[str] = []
    for plan in plans:
        brand, ws, book = plan["brand"], plan["ws"], plan["book"]
        stat = stat_of(prices, plan["nm_ids"])
        if args.save_snapshots:
            path = f"snapshot_prices_{brand.lower().replace(' ', '_')}.json"
            with open(path, "w", encoding="utf-8") as f:
                json.dump({"snapshot_at": datetime.now(sp.MSK).isoformat(timespec="seconds"),
                           "brand": brand, "dest": args.dest,
                           "source": "card.wb.ru/cards/v4/detail (wb кошелёк, floor ×0.98)",
                           "prices": {str(nm): prices.get(nm) for nm in plan["nm_ids"]}},
                          f, ensure_ascii=False, indent=2)
        if args.dry_run:
            results.append(f"{brand}: dry-run, {stat}")
            continue
        try:
            # Строки человек правит когда угодно, а колонка пишется по их
            # номерам — перечитываем лист прямо перед записью (тот же приём,
            # что в brand_shelves).
            rows, lay, values = read_price_rows(ws)
            date, down, up = write_column(book, ws, lay, values, rows, prices,
                                          datetime.now(sp.MSK),
                                          gradient=not args.no_gradient)
            results.append(f"{brand}: колонка за {date} записана ({stat}; "
                           f"подешевело {down}, подорожало {up})")
        except Exception as exc:                        # noqa: BLE001
            sp.log(f"ОШИБКА записи книги цен {brand}: {exc.__class__.__name__}: {exc}")
            failed.append(brand)
    return results, failed


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Цены по карточкам листа «Цены» книги цен бренда")
    ap.add_argument("--brand", default="all",
                    help="имя бренда или all — все книги со сбором цен "
                         f"({', '.join(brands_with_prices())})")
    ap.add_argument("--creds", default=None)
    ap.add_argument("--dest", type=int, default=wb_config.DEST)
    ap.add_argument("--save-snapshots", action="store_true", default=True)
    ap.add_argument("--dry-run", action="store_true", help="в таблицу не писать")
    ap.add_argument("--notify-fail", action="store_true",
                    help="упавшие книги — сообщением в Telegram")
    ap.add_argument("--no-gradient", action="store_true",
                    help="не пересобирать условное форматирование (интрадей)")
    args = ap.parse_args()

    enabled = brands_with_prices()
    if args.brand.lower() == "all":
        brands = list(enabled)
    else:
        match = [b for b in enabled if b.lower() == args.brand.strip().lower()]
        if not match:
            known = [b for b in bs.BRANDS if b.lower() == args.brand.strip().lower()]
            if known:
                # Кнопка «Обновить цены» стоит во всех книгах, а книга цен
                # заведена не у всех: это не ошибка прогона, просто нечего снимать.
                print(f"Готово: у бренда {known[0]} сбор цен не включён "
                      f"(есть у: {', '.join(enabled)}) — нечего снимать")
                return
            raise SystemExit(f"Сбор цен включён только для: {enabled}; "
                             f"пришло {args.brand!r}")
        brands = match

    install_retries()
    client = ts.get_client(args.creds)

    results, failed = run_all(client, brands, args)
    for line in results:
        print("Готово:", line)
    if failed:
        if args.notify_fail:
            notify.fail("Цены брендов: книги с ошибкой — " + ", ".join(failed),
                        "\n".join(results))
        raise SystemExit(f"Бренды с ошибкой: {', '.join(failed)}")


if __name__ == "__main__":
    main()
