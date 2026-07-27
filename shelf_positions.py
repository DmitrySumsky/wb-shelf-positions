#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Позиции наших карточек в рекомендательных полках конкурентов на Wildberries.

Механизм восстановлен разбором исходников расширения «ТурбоМинимум RM» v74
(27.07.2026), фича `campaign-edit-competitor-recom-shelves-positions`:

    GET https://recom.wb.ru/recom/ru/common/v8/search
        ?query=<nmID КОНКУРЕНТА>      <- артикул подставляется в поле запроса
        &resultset=catalog&curr=rub&spp=30
        &ab_testing=false&suppressSpellcheck=false
        &page=N&dest=<регион>&appType=1

Ответ — упорядоченный список товаров полки. Позиция нашей карточки = индекс+1
(в расширении это функция qC).

Отличия от расширения — сознательные:
  * расширение читает только page=1 (топ-100) и всё остальное помечает «>300»;
    здесь идёт постраничный обход по `total`, поэтому позиции 101+ видны реально;
  * расширение делает запрос на каждую пару «наш артикул × конкурент»;
    здесь полка конкурента забирается ОДИН раз, а в ней ищутся сразу все наши
    артикулы — объём работы не зависит от числа наших карточек.

Авторизация не нужна: recom.wb.ru и card.wb.ru отвечают без кук и токенов.
Антибот WBaaS стоит на www.wildberries.ru, но не на этих API-хостах.
"""

from __future__ import annotations

import argparse
import json
import os
import random
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta, timezone

import requests

RECOM_URL = "https://recom.wb.ru/recom/ru/common/v8/search"
CARD_URL = "https://card.wb.ru/cards/v4/detail"

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
)
HEADERS = {
    "User-Agent": UA,
    "Accept": "*/*",
    "Origin": "https://www.wildberries.ru",
    "Referer": "https://www.wildberries.ru/",
}

# Параметры полки — ровно те, что зашиты в расширении (объект OFt).
SHELF_PARAMS = {
    "ab_testing": "false",
    "curr": "rub",
    "resultset": "catalog",
    "spp": "30",
    "suppressSpellcheck": "false",
    "appType": "1",
}

DEST_MOSCOW = -1257786
PAGE_SIZE = 100          # WB иногда отдаёт 99 — по этому признаку останавливаться НЕЛЬЗЯ
MAX_PAGES = 8            # страховка от бесконечной пагинации
MSK = timezone(timedelta(hours=3))

_print_lock = threading.Lock()


def log(msg: str) -> None:
    with _print_lock:
        print(f"[{datetime.now(MSK):%H:%M:%S}] {msg}", flush=True)


def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    return s


# HTTP 200 с пустым телом — это штатный ответ WB на несуществующий артикул,
# а не сбой. Ретраить его бессмысленно, и мешать с сетевой ошибкой нельзя.
EMPTY = object()


def _get_json(session: requests.Session, url: str, params: dict, tries: int = 4):
    """GET с экспоненциальным backoff. Возвращает dict, EMPTY или None (сбой)."""
    delay = 1.0
    for attempt in range(1, tries + 1):
        try:
            r = session.get(url, params=params, timeout=25)
            if r.status_code == 200:
                if not r.content.strip():
                    return EMPTY
                return r.json()
            # 429/5xx — ждём дольше, остальное тоже ретраим, но без надежды
            if attempt == tries:
                log(f"  HTTP {r.status_code} после {tries} попыток: {url}")
                return None
        except Exception as exc:
            if attempt == tries:
                log(f"  сеть отвалилась после {tries} попыток ({exc.__class__.__name__}): {url}")
                return None
        time.sleep(delay + random.uniform(0, 0.4))
        delay *= 2
    return None


def collect_shelf(session: requests.Session, competitor_nm: int, dest: int,
                  max_positions: int = 600) -> dict:
    """
    Забирает полку конкурента целиком (постранично) и возвращает
    {"ids": [nmID по порядку], "total": сколько всего заявил WB, "pages": N}.
    """
    ids: list[int] = []
    total = None
    pages = 0
    status = "ok"

    for page in range(1, MAX_PAGES + 1):
        params = dict(SHELF_PARAMS, query=str(competitor_nm), page=str(page), dest=str(dest))
        data = _get_json(session, RECOM_URL, params)
        if data is EMPTY:
            # Пустое тело на первой же странице = артикула на WB нет.
            status = "missing" if page == 1 else "ok"
            break
        if data is None:
            status = "failed" if page == 1 else "ok"
            break
        pages += 1
        if total is None:
            total = data.get("total")
        products = data.get("products") or []
        if not products:
            break
        ids.extend(p["id"] for p in products if p.get("id") is not None)

        # Остановка ТОЛЬКО по total или лимиту. По len(products) < 100 останавливаться
        # нельзя: WB штатно отдаёт 99 товаров на первой странице.
        if total and len(ids) >= total:
            break
        if len(ids) >= max_positions:
            break
        time.sleep(0.25)

    return {"ids": ids, "total": total, "pages": pages, "status": status}


def fetch_card_info(session: requests.Session, nm_ids: list[int], dest: int) -> dict:
    """Бренд и название по списку артикулов (батчами по 100)."""
    info: dict[int, dict] = {}
    for i in range(0, len(nm_ids), 100):
        chunk = nm_ids[i:i + 100]
        params = {
            "appType": "1",
            "curr": "rub",
            "dest": str(dest),
            "nm": ";".join(str(x) for x in chunk),
        }
        data = _get_json(session, CARD_URL, params)
        for p in (data or {}).get("products") or []:
            info[p["id"]] = {
                "brand": p.get("brand") or "",
                "name": p.get("name") or "",
                "supplierId": p.get("supplierId"),
            }
        time.sleep(0.2)
    return info


def run(our_nm: list[int], competitors: list[int], dest: int = DEST_MOSCOW,
        workers: int = 4, max_positions: int = 600) -> dict:
    """
    Основной проход. Возвращает снапшот:
      positions[наш_nm][конкурент] = позиция (int) или None, если карточки в полке нет.
    """
    started = time.time()
    our_set = set(our_nm)
    shelves: dict[int, dict] = {}

    log(f"Полок к обходу: {len(competitors)}; наших артикулов: {len(our_nm)}; dest={dest}")

    def work(comp: int) -> tuple[int, dict]:
        s = _session()
        res = collect_shelf(s, comp, dest, max_positions)
        log(f"  полка {comp}: собрано {len(res['ids'])} из {res['total']} за {res['pages']} стр.")
        return comp, res

    with ThreadPoolExecutor(max_workers=workers) as pool:
        for comp, res in pool.map(work, competitors):
            shelves[comp] = res

    # Позиции: один проход по каждой полке закрывает сразу все наши артикулы.
    positions: dict[int, dict[int, int | None]] = {nm: {} for nm in our_nm}
    for comp, res in shelves.items():
        seen: dict[int, int] = {}
        for idx, nm in enumerate(res["ids"], start=1):
            if nm in our_set and nm not in seen:
                seen[nm] = idx
        for nm in our_nm:
            positions[nm][comp] = seen.get(nm)

    session = _session()
    log("Тяну бренды и названия…")
    cards = fetch_card_info(session, sorted(set(our_nm) | set(competitors)), dest)

    snapshot = {
        "snapshot_at": datetime.now(MSK).isoformat(timespec="seconds"),
        "dest": dest,
        "source": "recom.wb.ru/recom/ru/common/v8/search",
        "our_nm": our_nm,
        "competitors": competitors,
        "cards": {str(k): v for k, v in cards.items()},
        "shelves": {
            str(c): {"total": r["total"], "collected": len(r["ids"]), "pages": r["pages"]}
            for c, r in shelves.items()
        },
        "positions": {str(nm): {str(c): p for c, p in row.items()} for nm, row in positions.items()},
        "elapsed_sec": round(time.time() - started, 1),
    }
    found = sum(1 for row in positions.values() for p in row.values() if p)
    total_cells = len(our_nm) * len(competitors)
    log(f"Готово за {snapshot['elapsed_sec']} с. Найдено {found} из {total_cells} вхождений.")
    return snapshot


def run_groups(groups: list[dict], dest: int = DEST_MOSCOW, workers: int = 4,
               max_positions: int = 600) -> dict:
    """
    Проход по товарным группам.

    groups = [{"product": "...", "ours": [nmID...], "competitors": [nmID...]}, ...]

    Позиции считаются только внутри группы: наш «Магний» ищется в полках
    конкурентов по магнию, а не по коллагену. Полка каждого конкурента при этом
    забирается ОДИН раз, даже если он встречается в нескольких группах.
    """
    started = time.time()
    all_comps = sorted({c for g in groups for c in g["competitors"]})
    all_ours = sorted({o for g in groups for o in g["ours"]})

    log(f"Групп: {len(groups)}; наших артикулов: {len(all_ours)}; "
        f"уникальных полок к обходу: {len(all_comps)}; dest={dest}")

    shelves: dict[int, dict] = {}
    done = 0

    def work(comp: int) -> tuple[int, dict]:
        s = _session()
        return comp, collect_shelf(s, comp, dest, max_positions)

    with ThreadPoolExecutor(max_workers=workers) as pool:
        for comp, res in pool.map(work, all_comps):
            shelves[comp] = res
            done += 1
            if done % 25 == 0 or done == len(all_comps):
                log(f"  полок обойдено: {done}/{len(all_comps)}")

    # Полка, не отдавшаяся по сети, — это НЕ «нас там нет». Добиваем отдельным
    # проходом. Мёртвый артикул (пустое тело) не ретраим — там нечего добивать.
    failed = [c for c, r in shelves.items() if r["status"] == "failed"]
    if failed:
        log(f"Повторный проход по {len(failed)} не отдавшимся полкам…")
        s = _session()
        for comp in failed:
            time.sleep(2.0)
            res = collect_shelf(s, comp, dest, max_positions)
            if res["status"] != "failed":
                shelves[comp] = res
                log(f"  полка {comp}: добита ({res['status']}, {len(res['ids'])} позиций)")

    failed = [c for c, r in shelves.items() if r["status"] == "failed"]
    missing = [c for c, r in shelves.items() if r["status"] == "missing"]
    if failed:
        log(f"ВНИМАНИЕ: {len(failed)} полок не отдались по сети: {failed}")
    if missing:
        log(f"Артикулов нет на WB (карточка удалена/скрыта): {len(missing)} — {missing}")

    # Индекс «nmID -> позиция» по каждой полке, один раз.
    index: dict[int, dict[int, int]] = {}
    for comp, res in shelves.items():
        pos_of: dict[int, int] = {}
        for idx, nm in enumerate(res["ids"], start=1):
            pos_of.setdefault(nm, idx)
        index[comp] = pos_of

    positions: dict[str, dict[str, int | None]] = {}
    pairs = 0
    for g in groups:
        for our in g["ours"]:
            row = positions.setdefault(str(our), {})
            for comp in g["competitors"]:
                row[str(comp)] = index.get(comp, {}).get(our)
                pairs += 1

    session = _session()
    log("Тяну бренды и названия…")
    cards = fetch_card_info(session, sorted(set(all_ours) | set(all_comps)), dest)

    found = sum(1 for r in positions.values() for p in r.values() if p)
    snapshot = {
        "snapshot_at": datetime.now(MSK).isoformat(timespec="seconds"),
        "dest": dest,
        "source": "recom.wb.ru/recom/ru/common/v8/search",
        "groups": groups,
        "our_nm": all_ours,
        "competitors": all_comps,
        "failed_shelves": [str(c) for c in failed],
        "missing_shelves": [str(c) for c in missing],
        "cards": {str(k): v for k, v in cards.items()},
        "shelves": {
            str(c): {"total": r["total"], "collected": len(r["ids"]),
                     "pages": r["pages"], "status": r["status"]}
            for c, r in shelves.items()
        },
        "positions": positions,
        "elapsed_sec": round(time.time() - started, 1),
    }
    log(f"Готово за {snapshot['elapsed_sec']} с. Найдено {found} из {pairs} пар.")
    return snapshot


def load_config(path: str) -> dict:
    with open(path, encoding="utf-8") as f:
        cfg = json.load(f)
    if not cfg.get("our_nm"):
        raise SystemExit("В конфиге пустой our_nm — нечего искать.")
    if not cfg.get("competitors"):
        raise SystemExit("В конфиге пустой competitors — негде искать.")
    return cfg


def to_csv(snapshot: dict, path: str) -> None:
    import csv
    cards = snapshot["cards"]
    comps = snapshot["competitors"]
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        head = ["Наш артикул", "Бренд", "Название"]
        head += [f"{c} ({cards.get(str(c), {}).get('brand', '')})" for c in comps]
        w.writerow(head)
        for nm in snapshot["our_nm"]:
            info = cards.get(str(nm), {})
            row = [nm, info.get("brand", ""), info.get("name", "")]
            row += [snapshot["positions"][str(nm)].get(str(c)) or "" for c in comps]
            w.writerow(row)


def main() -> None:
    ap = argparse.ArgumentParser(description="Позиции наших карточек в полках конкурентов на WB")
    ap.add_argument("--config", default="config.json", help="JSON со списками артикулов")
    ap.add_argument("--out", default="snapshot.json", help="куда сохранить снапшот")
    ap.add_argument("--csv", default=None, help="дополнительно выгрузить матрицу в CSV")
    ap.add_argument("--dest", type=int, default=None, help="регион WB (по умолчанию Москва)")
    ap.add_argument("--workers", type=int, default=None)
    args = ap.parse_args()

    cfg = load_config(args.config)
    snapshot = run(
        our_nm=[int(x) for x in cfg["our_nm"]],
        competitors=[int(x) for x in cfg["competitors"]],
        dest=args.dest or int(cfg.get("dest", DEST_MOSCOW)),
        workers=args.workers or int(cfg.get("workers", 4)),
        max_positions=int(cfg.get("max_positions", 600)),
    )
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(snapshot, f, ensure_ascii=False, indent=2)
    log(f"Снапшот записан: {args.out}")
    if args.csv:
        to_csv(snapshot, args.csv)
        log(f"CSV записан: {args.csv}")


if __name__ == "__main__":
    main()
