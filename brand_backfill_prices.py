#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Разовая доливка ИСТОРИИ цен в лист «Цены» книги БРЕНДА (задача 10.08.2026).

Зачем: лист «Цены» книги Health Form заведён 07.08 и держит две даты — сравнивать
цену конкурента не с чем. Артур попросил «прогрузить историю за прошлые даты».

Отличие от `backfill_prices.py` (он доливает лист «Цены» ОБЩЕЙ книги): там
раскладка фиксирована и вторая строка шапки хранит время замера, здесь строки
ведёт человек, шапка одна строка, а фиксированные колонки ищутся по названиям
(`brand_shelves.read_layout`). Договор тот же, что у дневного прогона: значения
и порядок строк не трогаем, дописываем только колонки.

Два источника, в таком порядке (второй берётся только для тех, кого нет в первом):

1. **Таблица «Аналитика цен WB»**, лист «Аналитика цен (wb кошелек)» — там ту же
   метрику ежедневно ведёт `21.orders-cloud/.../prices_update.py`. Глубина ~70
   дней против 30-дневного окна MPStats, квота не тратится. Тождество метрики
   доказано при доливке общей книги 30.07.2026 (см. `_memory/PATTERNS.md`).
2. **MPStats** `wb/get/item/{nm}/sales`, поле `wallet_price` — для карточек, которых
   в соседней таблице нет (в книге бренда конкурентов больше: 663 артикула против
   326 в источнике). Окно MPStats — 30 дней, поэтому у этих строк история короче,
   и это нормально: пустая ячейка честнее выдуманной.

Куда пишет: колонками СПРАВА от существующего блока дат, по убыванию даты (свежая
левее — как во всём листе). Дневной прогон врезает свою колонку слева и лист не
перезаписывает, поэтому долитые колонки просто съезжают вправо.

Идемпотентность: дата, которая в листе уже есть, пропускается; доливаются только
даты СТАРШЕ самой старой в листе — свежие остаются делом дневного прогона.

Раскраска: после доливки пересобирается градиент по блокам (`brand_prices.
rebuild_gradients`) — он покрывает свежие `GRADIENT_DATES` колонок, глубокая
история остаётся числами без заливки.

Запуск (разово, локально):
    python brand_backfill_prices.py --brand "Health Form" \
        --source-id <ID «Аналитика цен WB»> --creds <sa.json> \
        [--mpstats-key-file ...] [--days 90] [--no-mpstats] [--dry-run]
"""

from __future__ import annotations

import argparse
import json
import os
import re
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime

import gspread

import backfill_prices as bfl
import brand_prices as bp
import brand_shelves as bs
import shelf_positions as sp
import to_sheets as ts
from vexor_shelves import install_retries

DEFAULT_MP_KEY_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "21.orders-cloud", "api_keys.txt")
NOTE_SOURCE = "История: таблица «Аналитика цен WB» (тот же wallet_price)"
NOTE_MPSTATS = "История: MPStats wallet_price"


def read_mp_token(path: str) -> str:
    """Токен MPStats из файла ключей 21.orders-cloud (строка MPSTATS_TOKEN=...)."""
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line.startswith("MPSTATS_TOKEN"):
                return line.split("=", 1)[1].strip().strip('"\'')
    raise SystemExit(f"В {path} нет строки MPSTATS_TOKEN=")


def mp_history(nm: str, token: str, tries: int = 4) -> dict[str, int]:
    """MPStats: {ISO-дата: цена с кошельком}. Недоступно — пустой словарь."""
    url = f"https://mpstats.io/api/wb/get/item/{nm}/sales"
    req = urllib.request.Request(
        url, headers={"X-Mpstats-TOKEN": token, "Content-Type": "application/json"})
    rows = None
    for i in range(tries):
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                rows = json.loads(r.read())
            break
        except urllib.error.HTTPError as e:
            if e.code in (429, 500, 502, 503):
                sp.time.sleep(1.0 * (i + 1))
                continue
            return {}
        except Exception:
            sp.time.sleep(0.7 * (i + 1))
    out: dict[str, int] = {}
    if isinstance(rows, list):
        for r in rows:
            d, v = r.get("data"), (r.get("wallet_price") or r.get("final_price"))
            if d and v:
                out[d] = round(v)
    return out


def iso_to_label(iso: str) -> str:
    """«2026-08-01» → «01.08» (формат колонок листа)."""
    y, m, d = iso.split("-")
    return f"{int(d):02d}.{int(m):02d}"


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Доливка истории цен в лист «Цены» книги бренда")
    ap.add_argument("--brand", default="Health Form")
    ap.add_argument("--source-id", default=os.environ.get("SOURCE_ID"))
    ap.add_argument("--source-tab", default=ts.SOURCE_TAB)
    ap.add_argument("--creds", default=None)
    ap.add_argument("--days", type=int, default=90,
                    help="сколько дат истории долить (по умолчанию 90)")
    ap.add_argument("--mpstats-key-file", default=DEFAULT_MP_KEY_FILE)
    ap.add_argument("--no-mpstats", action="store_true",
                    help="только соседняя таблица, MPStats не дёргать")
    ap.add_argument("--workers", type=int, default=6)
    ap.add_argument("--dry-run", action="store_true",
                    help="показать план и покрытие, в таблицу не писать")
    args = ap.parse_args()

    if not args.source_id:
        raise SystemExit("Не задан --source-id (таблица «Аналитика цен WB»)")
    if args.brand not in bs.BRANDS:
        raise SystemExit(f"Неизвестный бренд {args.brand!r}")

    install_retries()
    today = datetime.now(sp.MSK).date()
    client = ts.get_client(args.creds)

    hist, src_dates = bfl.read_source_history(client, args.source_id, args.source_tab)
    sp.log(f"Соседняя таблица: артикулов {len(hist)}, дат {len(src_dates)}")

    book = client.open_by_key(bs.BRANDS[args.brand]["sheet_id"])
    ws = book.worksheet(bp.SHEET_PRICES)
    values = ws.get_all_values()
    head = [str(x).strip() for x in values[0]]
    lay = bs.read_layout(head)
    if not lay["date_at"]:
        raise SystemExit("В листе «Цены» нет ни одной колонки-даты — сначала "
                         "должен пройти дневной brand_prices.py")

    ia, nfix = lay["art"], lay["nfix"]
    art_of: dict[int, str] = {}
    for i, raw in enumerate(values[1:], start=2):
        cells = [str(x).strip() for x in raw]
        art = cells[ia] if len(cells) > ia else ""
        if art.isdigit():
            art_of[i] = art
    if not art_of:
        raise SystemExit("В листе «Цены» не нашлось строк с артикулами")
    arts = sorted(set(art_of.values()))

    have = set(lay["date_at"])
    oldest = min(bfl.to_date(d, today) for d in have)
    todo = [d for d in src_dates
            if d not in have and bfl.to_date(d, today) < oldest][:args.days]

    # MPStats — только для карточек, которых нет в соседней таблице: она и глубже,
    # и бесплатна, а квоту MPStats тратить на уже известные цены незачем.
    missing = [a for a in arts if a not in hist]
    sp.log(f"Лист «{bp.SHEET_PRICES}»: строк с артикулом {len(art_of)}, уникальных "
           f"{len(arts)}; в соседней таблице {len(arts) - len(missing)}, "
           f"нет там {len(missing)}")

    mp_hist: dict[str, dict[str, int]] = {}
    if missing and not args.no_mpstats:
        token = read_mp_token(args.mpstats_key_file)
        sp.log(f"MPStats: тяну историю по {len(missing)} карточкам (окно 30 дней)")
        with ThreadPoolExecutor(max_workers=args.workers) as ex:
            for nm, h in zip(missing, ex.map(lambda a: mp_history(a, token), missing)):
                if h:
                    mp_hist[nm] = {iso_to_label(k): v for k, v in h.items()}
        got = sum(1 for v in mp_hist.values() if v)
        sp.log(f"MPStats ответил по {got} карточкам из {len(missing)}")
        extra = sorted({d for h in mp_hist.values() for d in h},
                       key=lambda x: bfl.to_date(x, today), reverse=True)
        todo += [d for d in extra
                 if d not in have and d not in todo and bfl.to_date(d, today) < oldest]

    todo = sorted(todo, key=lambda x: bfl.to_date(x, today), reverse=True)[:args.days]
    if not todo:
        print(f"Доливать нечего: все даты источников старше {oldest:%d.%m} уже в листе.")
        return

    first, last = min(art_of), max(art_of)
    grid: list[list[object]] = []
    for rn in range(first, last + 1):
        a = art_of.get(rn, "")
        h = hist.get(a) or mp_hist.get(a) or {}
        grid.append([h.get(d, "") for d in todo])

    filled = sum(1 for row in grid for v in row if v != "")
    sp.log(f"Доливаю {len(todo)} дат ({todo[0]} … {todo[-1]}) в строки {first}–{last}; "
           f"заполнится ячеек {filled} из {len(art_of) * len(todo)}")

    if args.dry_run:
        print("dry-run: в таблицу не пишу. Первые строки:")
        for rn in range(first, min(first + 5, last + 1)):
            print(f"  {art_of.get(rn, '—'):>12}: " + ", ".join(
                f"{d}={v or '—'}" for d, v in zip(todo[:6], grid[rn - first])))
        return

    ndates_before = len(lay["date_at"])
    start_col1 = nfix + ndates_before + 1          # сразу за блоком дат, 1-базно
    end_col1 = start_col1 + len(todo) - 1
    if ws.col_count < end_col1:
        book.batch_update({"requests": [{"updateSheetProperties": {
            "properties": {"sheetId": ws.id,
                           "gridProperties": {"columnCount": end_col1 + 2}},
            "fields": "gridProperties.columnCount"}}]})

    ws.update(values=[todo],
              range_name=(f"{bs.col_letter(start_col1)}1:"
                          f"{bs.col_letter(end_col1)}1"),
              value_input_option="RAW")
    ws.update(values=grid,
              range_name=(f"{bs.col_letter(start_col1)}{first}:"
                          f"{bs.col_letter(end_col1)}{last}"),
              value_input_option="USER_ENTERED")

    src_note = NOTE_SOURCE if not mp_hist else f"{NOTE_SOURCE}; {NOTE_MPSTATS}"
    reqs = [
        {"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": start_col1 - 1, "endColumnIndex": end_col1},
            "cell": {"note": src_note,
                     "userEnteredFormat": {"textFormat": {"bold": True},
                                           "horizontalAlignment": "CENTER"}},
            "fields": "note,userEnteredFormat(textFormat,horizontalAlignment)"}},
        {"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": last,
                      "startColumnIndex": start_col1 - 1, "endColumnIndex": end_col1},
            "cell": {"userEnteredFormat": {
                "horizontalAlignment": "CENTER",
                "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}},
            "fields": "userEnteredFormat(horizontalAlignment,numberFormat)"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": start_col1 - 1, "endIndex": end_col1},
            "properties": {"pixelSize": 75}, "fields": "pixelSize"}},
    ]
    book.batch_update({"requests": reqs})

    lay["date_at"] = dict(lay["date_at"], **{d: start_col1 - 1 + i
                                             for i, d in enumerate(todo)})
    blocks = bp.read_blocks(values, lay)
    nrules = bp.rebuild_gradients(book, ws, lay, blocks, ndates_before + len(todo))
    print(f"Готово: долито {len(todo)} дат ({todo[-1]} … {todo[0]}), ячеек {filled}; "
          f"градиент пересобран ({nrules} правил на {len(blocks)} блоков)")


if __name__ == "__main__":
    main()
