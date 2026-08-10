#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сводка по книгам брендов в Telegram: что реально лежит в таблицах.

v1.0.0 — 10.08.2026
  КНИГА 4ME МОЛЧА НЕ ОБНОВЛЯЛАСЬ С 07.08 — прогон упирался в лимит времени
  задания и убивался на последнем бренде; снаружи это выглядело как «таблица
  сломалась». Сводка читает САМИ ТАБЛИЦЫ, а не журнал прогона: свежая дата
  колонки, сколько ячеек заполнено, сколько ошибок сбора. Отставшая книга
  видна, даже если workflow отчитался зелёным (и наоборот).

Что шлётся:
    • строка на бренд по полкам и по ценам: дата колонки, заполнено, ошибки;
    • отдельным блоком ⚠ книги, где свежая дата не сегодняшняя.

Запуск (после ежедневного прогона, шагом в shelves.yml):
    python digest.py --send
Без `--send` печатает сводку в лог и ничего не отправляет.
"""

from __future__ import annotations

import argparse
import statistics
from datetime import datetime

import gspread

import brand_shelves as bs
import notify
import prices_to_sheets as pts
import shelf_positions as sp
import to_sheets as ts
import wb_config
from vexor_shelves import install_retries

SHEET_SHELVES = "Полки"
SHEET_PRICES = "Цены"


def _fresh_column(ws) -> tuple[str | None, list[str]]:
    """(дата самой свежей колонки, её значения). Свежая дата стоит СЛЕВА."""
    head_rows = ws.get_values("A1:CZ1")
    if not head_rows:
        return None, []
    head = [str(x).strip() for x in head_rows[0]]
    lay = bs.read_layout(head)
    if not lay["date_at"]:
        return None, []
    col = min(lay["date_at"].values())
    date = next(d for d, i in lay["date_at"].items() if i == col)
    letter = bs.col_letter(col + 1)
    values = ws.get_values(f"{letter}2:{letter}{ws.row_count}")
    return date, [str(r[0]).strip() if r else "" for r in values]


def _days_ago(date: str, now: datetime) -> int | None:
    """«07.08» → сколько дней назад. Год берём текущий, декабрьский переход — минус год."""
    try:
        day, month = (int(x) for x in date.split("."))
    except ValueError:
        return None
    when = datetime(now.year, month, day)
    if (when - now.replace(tzinfo=None)).days > 30:
        when = when.replace(year=now.year - 1)
    return (now.replace(tzinfo=None).date() - when.date()).days


def shelves_line(ws, now: datetime) -> tuple[str, int | None]:
    date, values = _fresh_column(ws)
    if not date:
        return "полки — колонок с датой нет", None
    nums = [int(v) for v in values if v.isdigit()]
    fails = sum(1 for v in values if v == bs.STATE_FAIL)
    none_here = sum(1 for v in values if v == bs.NOT_IN_SHELF)
    no_product = sum(1 for v in values if v == bs.STATE_NO_PRODUCT)
    med = int(statistics.median(nums)) if nums else 0
    return (f"полки {date} · позиций {len(nums)} · нет в полке {none_here} · "
            f"нет товара {no_product} · ошибок {fails} · медиана {med}",
            _days_ago(date, now))


def prices_line(ws, now: datetime) -> tuple[str, int | None]:
    date, values = _fresh_column(ws)
    if not date:
        return "цены — колонок с датой нет", None
    nums = [v for v in values if v.replace(" ", "").replace("\xa0", "").isdigit()]
    fails = sum(1 for v in values if v == pts.STATE_FAIL)
    none_stock = sum(1 for v in values if v == pts.STATE_NONE)
    gone = sum(1 for v in values if v == pts.STATE_GONE)
    # Считаем ЯЧЕЙКИ колонки, а не уникальные карточки: одна карточка стоит в
    # нескольких группах, и вопрос сводки — «колонка заполнена целиком?».
    return (f"цены {date} · заполнено {len(nums)} · нет в наличии {none_stock} · "
            f"нет карточки {gone} · ошибок {fails}",
            _days_ago(date, now))


def collect(client) -> tuple[list[str], list[str]]:
    """(строки сводки, предупреждения об отставших книгах)."""
    now = datetime.now(sp.MSK)
    lines, warn = [], []
    for brand, cfg in bs.BRANDS.items():
        lines.append(brand)
        try:
            ws = client.open_by_key(cfg["sheet_id"]).worksheet(SHEET_SHELVES)
            text, ago = shelves_line(ws, now)
            lines.append("  " + text)
            if ago:
                warn.append(f"{brand}: полки отстали на {ago} дн. (последняя {text.split()[1]})")
        except Exception as exc:                        # noqa: BLE001
            lines.append(f"  полки — НЕ ПРОЧЛИСЬ: {exc.__class__.__name__}")
            warn.append(f"{brand}: книга полок не прочлась — {exc.__class__.__name__}")
        if not cfg["prices"]:
            continue
        try:
            ws = client.open_by_key(cfg["prices_book"]).worksheet(SHEET_PRICES)
            text, ago = prices_line(ws, now)
            lines.append("  " + text)
            if ago:
                warn.append(f"{brand}: цены отстали на {ago} дн. (последняя {text.split()[1]})")
        except gspread.WorksheetNotFound:
            lines.append("  цены — листа «Цены» в книге цен ещё нет")
        except Exception as exc:                        # noqa: BLE001
            lines.append(f"  цены — НЕ ПРОЧЛИСЬ: {exc.__class__.__name__}")
            warn.append(f"{brand}: книга цен не прочлась — {exc.__class__.__name__}")
    return lines, warn


def main() -> None:
    ap = argparse.ArgumentParser(description="Сводка по книгам брендов в Telegram")
    ap.add_argument("--creds", default=None)
    ap.add_argument("--send", action="store_true", help="отправить в Telegram")
    ap.add_argument("--title", default="Полки и цены WB")
    args = ap.parse_args()

    install_retries()
    client = ts.get_client(args.creds)
    lines, warn = collect(client)

    now = datetime.now(sp.MSK)
    head = f"{'⚠️' if warn else '📊'} {args.title} — {now.strftime('%d.%m %H:%M')} МСК"
    text = head + "\n\n" + "\n".join(lines)
    if warn:
        text += "\n\nОтстают:\n" + "\n".join("• " + w for w in warn)
    print(text)

    if args.send and not notify.send(text):
        print("Сводка НЕ отправлена: нет секретов Telegram или ошибка отправки")


if __name__ == "__main__":
    main()
