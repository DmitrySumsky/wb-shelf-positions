#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сбор полок и цен из браузера менеджера — облачная половина (v2.4, 22.09.2026).

ЗАЧЕМ. С ночи на 22.09.2026 WB закрыл антиботом публичные адреса витрины
(`recom.wb.ru`, `card.wb.ru`, `search.wb.ru`): из облака и с любого «голого»
клиента — `403 Forbidden`. Сайт теперь ходит за теми же данными через
`www.wildberries.ru/__internal/…`, и пускают туда только браузер, прошедший
проверку WB (кука `x_wbaas_token`). Проверку проходит человек, а не мы:
обходить её роботом не делаем.

КАК ТЕПЕРЬ. Сбор идёт в Chrome менеджера расширением (модуль
`mp-core/modules/wb-shelf-browser`), по кнопке в книге бренда. Одна кнопка
собирает весь КОНТУР — набор книг из `brands.json → contours`, — поэтому
неважно, кто нажал первым: следующий зайдёт в уже обновлённую книгу.

Эта программа — облачная половина:

    plan   — прочитать книги контура и собрать план для расширения:
             какие полки листать, какие наши карточки в каждой искать, по каким
             артикулам снять цены. План кладётся в хаб (веб-приложение Apps
             Script), откуда его забирает расширение.
    fetch  — забрать из хаба запись расширения (итог сбора) в файл; дальше
             её разносят `brand_shelves.py --from-browser` и
             `brand_prices.py --from-browser` тем же кодом, что и раньше.
    probe  — открыта ли витрина для облака. Код выхода 0 — открыта (старый
             сбор из Actions работает), 1 — закрыта (облачные прогоны сбор
             пропускают, чтобы не писать «ошибку сбора» и не жечь 25 минут).

Хаб: `SHELF_HUB_URL` (адрес веб-приложения) и `SHELF_HUB_KEY` (общий ключ) —
секреты репозитория; локально — `21.orders-cloud/api_keys_extra.txt`, строки
`SHELF_HUB_URL` / `SHELF_HUB_KEY`.

Запуск:
    python browser_job.py plan --contour brands --push
    python browser_job.py plan --contour brands --out plan.json     # без хаба
    python browser_job.py fetch --contour brands --out result.json
    python browser_job.py probe
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime

import requests

import brand_prices as bp
import brand_shelves as bs
import shelf_positions as sp
import to_sheets as ts
import wb_config
from vexor_shelves import install_retries

PLAN_VERSION = 1
# Источник групп общей книги («Аналитика цен WB»); в облаке — секрет SOURCE_ID.
COMMON_SOURCE_ID = "1YWWqK2Fh9NTJCVW7j7xocChUf2Pu9oNnvi1_oqMPi_Q"
MAX_POSITIONS = 600        # глубже 600-й позиции полку не листаем (как в облаке)


def contour_brands(name: str) -> list[str]:
    contours = wb_config.CFG.get("contours") or {}
    if name not in contours:
        raise SystemExit(f"Неизвестный контур {name!r}; есть: {list(contours)}")
    return list(contours[name]["brands"])


def build_plan(client, contour: str) -> dict:
    """План контура: {полка: [наши карточки, которые в ней искать]} + артикулы цен.

    Наши карточки задаются по полке, а не общим списком: расширение бросает
    листать полку, как только нашло всех «своих», — при медиане позиции 30–40
    это одна страница вместо шести.
    """
    shelves: dict[int, set[int]] = {}
    cards: set[int] = set()
    books: list[str] = []
    for brand in contour_brands(contour):
        cfg = bs.BRANDS[brand]
        try:
            ws = client.open_by_key(cfg["sheet_id"]).worksheet(bs.SHEET_DST)
            rows, _lay, _values, no_card = bs.read_sheet(ws, cfg["aliases"])
            groups = bs.build_groups(rows, no_card)
        except Exception as exc:                             # noqa: BLE001
            sp.log(f"ОШИБКА чтения полок {brand}: {exc.__class__.__name__}: {exc}")
            continue
        for g in groups:
            for comp in g["competitors"]:
                shelves.setdefault(comp, set()).update(g["ours"])
            cards.update(g["ours"])
            cards.update(g["competitors"])
        books.append(f"{brand}: полки")
        if cfg.get("prices") and cfg.get("prices_book"):
            try:
                _book, pws = bp.open_prices_ws(client, brand)
                prow, _l, _v = bp.read_price_rows(pws)
                cards.update(int(r["nm"]) for r in prow if r["nm"])
                books.append(f"{brand}: цены")
            except BaseException as exc:                     # noqa: BLE001
                sp.log(f"ОШИБКА чтения цен {brand}: {exc.__class__.__name__}: {exc}")

    # v2.5. Книги вне брендов: VEXOR (лист «Сводная» чужой книги) и общая книга
    # (списки групп — таблица «Аналитика цен WB»). Полки и карточки тех же видов.
    extra = (wb_config.CFG.get("contours") or {})[contour].get("extra", [])
    extra_groups: list[tuple[str, list[dict]]] = []
    if "vexor" in extra:
        try:
            import vexor_shelves as vs
            vid = os.environ.get("VEXOR_SHEET_ID") or vs.VEXOR_SHEET_ID
            vgroups, _rows = vs.read_blocks(client.open_by_key(vid))
            extra_groups.append(("VEXOR: полки", vgroups))
        except Exception as exc:                             # noqa: BLE001
            sp.log(f"ОШИБКА чтения книги VEXOR: {exc.__class__.__name__}: {exc}")
    if "common" in extra:
        try:
            sid = os.environ.get("SOURCE_ID") or COMMON_SOURCE_ID
            cgroups, _skipped = ts.read_groups(client, sid, ts.SOURCE_TAB)
            extra_groups.append(("Общая книга: полки и цены", cgroups))
        except Exception as exc:                             # noqa: BLE001
            sp.log(f"ОШИБКА чтения общей книги: {exc.__class__.__name__}: {exc}")
    for label, groups in extra_groups:
        for g in groups:
            for comp in g["competitors"]:
                shelves.setdefault(comp, set()).update(g["ours"])
            cards.update(g["ours"])
            cards.update(g["competitors"])
        books.append(label)

    plan = {
        "v": PLAN_VERSION,
        "contour": contour,
        "title": (wb_config.CFG.get("contours") or {})[contour].get("title", contour),
        "built_at": datetime.now(sp.MSK).isoformat(timespec="seconds"),
        "dest": wb_config.DEST,
        "max_positions": MAX_POSITIONS,
        "books": books,
        # v2.5: планы поиска, которые расширение пройдёт после полок и цен.
        "also": list((wb_config.CFG.get("contours") or {})[contour].get("also", [])),
        "shelves": {str(c): sorted(t) for c, t in sorted(shelves.items())},
        "cards": sorted(cards),
    }
    sp.log(f"План «{contour}»: книг {len(books)}, полок {len(shelves)}, "
           f"пар полка×наша {sum(len(t) for t in shelves.values())}, "
           f"карточек под цены {len(cards)}")
    return plan


# ---------------------------------------------------------------------- хаб

def hub() -> tuple[str, str]:
    url, key = os.environ.get("SHELF_HUB_URL", ""), os.environ.get("SHELF_HUB_KEY", "")
    if not url or not key:
        raise SystemExit("Нет SHELF_HUB_URL / SHELF_HUB_KEY — адрес хаба и ключ "
                         "(секреты репозитория, локально api_keys_extra.txt)")
    return url, key


def hub_call(action: str, contour: str, body: dict | None = None) -> dict:
    """Запрос в хаб. Ответ Apps Script приходит после 302 на googleusercontent —
    requests проходит его сам (POST при этом штатно становится GET)."""
    url, key = hub()
    params = {"action": action, "contour": contour, "key": key}
    if body is None:
        r = requests.get(url, params=params, timeout=120)
    else:
        # v2.4.2. Редиректы POST проходим сами: из GitHub Actions script.google.com
        # иногда сначала уводит на другой адрес того же скрипта, requests
        # превращал POST в GET, и хаб отвечал «неизвестное действие put_plan».
        # Ответ «уже обработано» — 302 на googleusercontent (его читаем GET),
        # любой другой 30x — POST повторяется по новому адресу.
        data = json.dumps(body, ensure_ascii=False).encode("utf-8")
        target, send_params = url, params
        for _ in range(5):
            r = requests.post(target, params=send_params, data=data,
                              headers={"Content-Type": "text/plain"},
                              timeout=120, allow_redirects=False)
            loc = r.headers.get("Location", "")
            if r.status_code not in (301, 302, 303, 307, 308) or not loc:
                break
            if "googleusercontent.com" in loc:
                r = requests.get(loc, timeout=120)
                break
            target, send_params = loc, None
    try:
        data = r.json()
    except ValueError:
        raise SystemExit(f"Хаб ответил не JSON ({r.status_code}): {r.text[:300]}")
    if not data.get("ok"):
        raise SystemExit(f"Хаб отказал: {data.get('error') or data}")
    return data


# --------------------------------------------------------------------- probe

def probe() -> bool:
    """Открыта ли витрина для этого клиента: один запрос карточки."""
    try:
        r = requests.get(sp.CARD_URL, params={"appType": "1", "curr": "rub",
                                               "dest": str(wb_config.DEST), "spp": "30",
                                               "nm": "219503618"},
                         headers=sp.HEADERS, timeout=25)
    except requests.RequestException as exc:
        print(f"Витрина WB не отвечает: {exc.__class__.__name__}")
        return False
    ok = r.status_code == 200
    print(f"Витрина WB для облака: {'открыта' if ok else 'ЗАКРЫТА'} (HTTP {r.status_code})")
    return ok


def main() -> None:
    ap = argparse.ArgumentParser(description="Сбор полок из браузера — облачная половина")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("plan")
    p.add_argument("--contour", default="all")
    p.add_argument("--out", default=None)
    p.add_argument("--push", action="store_true", help="положить план в хаб")
    p.add_argument("--creds", default=None)
    f = sub.add_parser("fetch")
    f.add_argument("--contour", default="all")
    f.add_argument("--out", default="browser_result.json")
    sub.add_parser("probe")
    args = ap.parse_args()

    if args.cmd == "probe":
        sys.exit(0 if probe() else 1)

    if args.cmd == "plan":
        install_retries()
        plan = build_plan(ts.get_client(args.creds), args.contour)
        if args.out:
            with open(args.out, "w", encoding="utf-8") as fh:
                json.dump(plan, fh, ensure_ascii=False)
            sp.log(f"План сохранён: {args.out}")
        if args.push:
            res = hub_call("put_plan", args.contour, plan)
            sp.log(f"План в хабе: {res.get('size')} символов")
        return

    if args.cmd == "fetch":
        res = hub_call("get_result", args.contour)
        rec = res.get("result")
        if not rec or not rec.get("shelves"):
            raise SystemExit("В хабе нет записи сбора по контуру " + args.contour)
        with open(args.out, "w", encoding="utf-8") as fh:
            json.dump(rec, fh, ensure_ascii=False)
        sp.log(f"Запись сбора от {rec.get('at')} ({rec.get('who') or 'кто — не указано'}): "
               f"полок {len(rec.get('shelves') or {})}, карточек {len(rec.get('cards') or {})}"
               f" → {args.out}")


if __name__ == "__main__":
    main()
