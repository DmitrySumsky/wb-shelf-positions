#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Конкуренты из ПОИСКОВОЙ ВЫДАЧИ WB вместо ручного списка.

Экспериментальная ветка задачи Артура от 29.07.2026: расширение Вячеслава Малых
берёт «токен» из поиска ЛК и строит по нему таблицу конкурентов. Токен не нужен —
поисковая выдача отдаётся публично, теми же условиями, что recom.wb.ru и card.wb.ru:

    GET https://search.wb.ru/exactmatch/ru/common/v12/search
        ?query=<слово>&resultset=catalog&curr=rub&spp=30
        &ab_testing=false&suppressSpellcheck=false
        &page=N&dest=<регион>&appType=1&sort=popular

Ответ: {"metadata": {...}, "products": [...], "total": N}. В каждом товаре уже есть
`id` (nmID), `brand`, `supplierId`, `subjectId`, `name` — отдельный запрос к карточке
не нужен. 100 товаров на страницу, глубина ровно 100 страниц (10 000 позиций),
дальше `products` пуст.

Модуль только СОБИРАЕТ группы в формате `shelf_positions.run_groups()`; сам обход
полок и запись в таблицу — там же, где и у боевого контура. Ничего из работающего
кода не переопределяется.
"""

from __future__ import annotations

import collections
import time

import shelf_positions as sp

SEARCH_URL = "https://search.wb.ru/exactmatch/ru/common/{ver}/search"

# Версию WB меняет молча: 28.07.2026 живы v4…v12, а v13 отдаёт битое тело
# (метаданные без products). Поэтому не хардкодим одну, а берём первую живую.
SEARCH_VERSIONS = ("v12", "v11", "v10", "v9", "v8", "v5", "v4")

PAGE_SIZE = 100
MAX_SEARCH_PAGES = 100      # потолок самого WB: со 101-й страницы products пуст

# Насколько глубоко смотрим выдачу в поисках НАШИХ карточек, даже если конкурентов
# берём только из топа: наш Health Form по «магний» стоит на 283-й позиции, при
# top_n=200 он бы просто не нашёлся.
OUR_SCAN_DEPTH = 1000

# Одного продавца в выборке — не больше стольких карточек. Иначе топ по узкому
# запросу забивается одним и тем же поставщиком, и мы платим по запросу за полку
# за каждый его дубль.
MAX_PER_SUPPLIER = 3

_version: str | None = None


def resolve_version(session, probe: str = "магний") -> str:
    """Первая версия эндпоинта, реально отдающая товары. Определяется один раз."""
    global _version
    if _version:
        return _version
    for ver in SEARCH_VERSIONS:
        data = search_page(session, probe, page=1, dest=sp.DEST_MOSCOW, ver=ver)
        if isinstance(data, dict) and data.get("products"):
            _version = ver
            sp.log(f"Поиск WB: работает {ver}")
            return ver
        time.sleep(0.5)
    raise SystemExit("Ни одна версия search.wb.ru не отдала товары — эндпоинт переехал.")


def search_page(session, query: str, page: int, dest: int, ver: str, tries: int = 4) -> dict | None:
    """
    Одна страница выдачи. None = страницу получить не удалось (это НЕ «конец выдачи»).

    Грабля (29.07.2026): примерно каждый десятый ответ приходит в чужом конверте —
    `{"metadata", "state", "version", "params", "data"}` вместо `{"metadata",
    "products", "total"}`, внутри `data.products` лежит один посторонний товар
    (ловил Apple по запросу «магний») и `params.dest` не тот, что просили. Похоже
    на промах кэша на стороне WB. Читать `data.products` НЕЛЬЗЯ — в конкуренты
    приедет мусор; считать концом выдачи тоже нельзя — сбор молча обрежется.
    Единственно верное — перезапросить ту же страницу.
    """
    params = dict(sp.SHELF_PARAMS, query=query, page=str(page), dest=str(dest), sort="popular")
    url = SEARCH_URL.format(ver=ver)
    for attempt in range(1, tries + 1):
        data = sp._get_json(session, url, params)
        if data is sp.EMPTY or data is None:
            return None
        if isinstance(data, dict) and "products" in data:
            return data
        if attempt == tries:
            sp.log(f"  «{query}» стр. {page}: WB {tries} раза подряд отдал чужой ответ")
            return None
        time.sleep(0.8 * attempt)
    return None


def collect_query(session, query: str, depth: int, dest: int) -> tuple[list[dict], int | None, bool]:
    """
    Выдача по запросу до `depth` позиций.

    Возвращает (товары по порядку, сколько всего заявил WB, дошли ли до конца).
    Порядок сохраняется: индекс+1 = позиция в выдаче. Третий элемент — честный
    признак полноты: False означает «страница не отдалась», а не «выдача кончилась».
    """
    ver = resolve_version(session)
    products: list[dict] = []
    total = None
    complete = True

    pages = min(MAX_SEARCH_PAGES, -(-depth // PAGE_SIZE))
    for page in range(1, pages + 1):
        data = search_page(session, query, page, dest, ver)
        if not data:
            sp.log(f"  «{query}»: страница {page} не отдалась, остановился на {len(products)}")
            complete = False
            break
        if total is None:
            total = data.get("total")
        chunk = data.get("products") or []
        if not chunk:
            break                       # выдача кончилась по-настоящему
        products.extend(chunk)
        if len(products) >= depth:
            break
        time.sleep(0.3)

    return products[:depth], total, complete


def _norm(s: str) -> str:
    return (s or "").strip().lower()


def build_group(query: str, products: list[dict], our_brands: list[str],
                top_n: int, our_nm: list[int] | None = None,
                subject_filter: bool = True) -> tuple[dict, dict]:
    """
    Товарная группа из выдачи: наши карточки + конкуренты из топ-N.

    Формат группы — тот же, что читает `run_groups()` и `to_sheets.build_entries()`:
    {"product", "ours", "competitors", "brand_of"}.
    """
    our_norm = {_norm(b): b for b in our_brands}
    stats = {"query": query, "scanned": len(products)}

    # Категория запроса = самый частый subjectId в первой сотне. По «магний» это
    # 1524 (БАДы) — 215 товаров из 300; остальные 28% это витаминные комплексы,
    # спортпит и просто мусор, полки которых сравнивать не с чем.
    subject = None
    if subject_filter and products:
        subject = collections.Counter(
            p.get("subjectId") for p in products[:PAGE_SIZE] if p.get("subjectId")
        ).most_common(1)[0][0]
    stats["subject"] = subject

    brand_of: dict[int, str] = {}
    ours: list[int] = []
    ours_seen: dict[str, int] = {}          # бренд -> лучшая позиция, по одной карточке на бренд
    ours_pos: dict[int, int] = {}

    for pos, p in enumerate(products, start=1):
        nm, brand = p.get("id"), p.get("brand") or ""
        if nm is None:
            continue
        key = _norm(brand)
        if key in our_norm and key not in ours_seen:
            ours_seen[key] = pos
            ours.append(int(nm))
            ours_pos[int(nm)] = pos
            brand_of[int(nm)] = our_norm[key]

    # Явно заданные артикулы важнее автоподбора: менеджер может следить за
    # конкретной карточкой, а не за той, что сегодня выше в выдаче.
    if our_nm:
        ours = [int(x) for x in our_nm]
        want = set(ours)
        ours_pos = {}
        for pos, p in enumerate(products, start=1):
            nm = p.get("id")
            if nm is not None and int(nm) in want and int(nm) not in ours_pos:
                ours_pos[int(nm)] = pos
                brand_of[int(nm)] = p.get("brand") or ""
        # Наша карточка может вообще не попасть в просмотренную глубину выдачи —
        # это не повод её терять, позиции в полках считаются всё равно.
        for nm in ours:
            brand_of.setdefault(nm, "")

    stats["ours"] = {nm: ours_pos.get(nm) for nm in ours}

    our_ids = set(ours)
    competitors: list[int] = []
    per_supplier: collections.Counter = collections.Counter()
    dropped = {"наши": 0, "другая категория": 0, "лимит на продавца": 0, "дубли": 0}
    seen: set[int] = set()

    for p in products:
        if len(competitors) >= top_n:
            break
        nm = p.get("id")
        if nm is None:
            continue
        nm = int(nm)
        # Отсекаем НАШИ БРЕНДЫ ЦЕЛИКОМ, а не только выбранную карточку: у одного
        # бренда в выдаче по «магний» стоит 5 своих артикулов, и четыре лишних
        # уехали бы в таблицу строками «Бренд конкурента: NATURI».
        if nm in our_ids or _norm(p.get("brand")) in our_norm:
            dropped["наши"] += 1
            continue
        if nm in seen:
            dropped["дубли"] += 1
            continue
        if subject and p.get("subjectId") != subject:
            dropped["другая категория"] += 1
            continue
        sup = p.get("supplierId")
        if sup is not None and per_supplier[sup] >= MAX_PER_SUPPLIER:
            dropped["лимит на продавца"] += 1
            continue
        seen.add(nm)
        per_supplier[sup] += 1
        competitors.append(nm)
        brand_of[nm] = p.get("brand") or ""

    stats["competitors"] = len(competitors)
    stats["dropped"] = dropped
    # Выдача может кончиться раньше, чем наберётся top_n — это не ошибка, но знать надо.
    stats["short"] = len(competitors) < top_n

    group = {
        "product": query,
        "ours": ours,
        "competitors": competitors,
        # ключи строками — так же, как их потом читает brand_of() из to_sheets
        "brand_of": {str(k): v for k, v in brand_of.items()},
        "source": "search",
    }
    return group, stats


def build_groups(queries: list[dict], our_brands: list[str], dest: int = sp.DEST_MOSCOW,
                 max_shelves: int | None = None) -> tuple[list[dict], list[str]]:
    """
    queries = [{"query": "магний", "top_n": 200, "our_nm": [...], "subject_filter": True}, ...]

    Возвращает (группы, отчёт строками). Отчёт печатается и уходит в снапшот —
    в нём видно, сколько выдачи отброшено и почему, чтобы «взяли топ-200» не
    выглядело как «обошли всех конкурентов».
    """
    session = sp._session()
    groups: list[dict] = []
    report: list[str] = []

    for q in queries:
        query = q["query"].strip()
        top_n = int(q.get("top_n") or 200)
        depth = max(top_n, OUR_SCAN_DEPTH)
        sp.log(f"Выдача по «{query}»: беру до {depth} позиций, конкурентов до {top_n}")

        products, total, complete = collect_query(session, query, depth, dest)
        group, st = build_group(query, products, our_brands, top_n,
                                our_nm=q.get("our_nm"),
                                subject_filter=q.get("subject_filter", True))

        found = ", ".join(f"{brand_of_nm(group, nm)} {nm} (поз. {p or '—'})"
                          for nm, p in st["ours"].items()) or "НЕ НАЙДЕНЫ"
        # Бренд, которого нет в просмотренной глубине, — это не пустая колонка в
        # матрице, а факт: по этому запросу нас там не видно. Говорим об этом прямо.
        absent = [b for b in our_brands
                  if _norm(b) not in {_norm(brand_of_nm(group, nm)) for nm in group["ours"]}]
        if absent:
            found += f"; не найдены в первых {depth} позициях: {', '.join(absent)}"
        line = (f"«{query}»: всего в выдаче {total}, просмотрено {st['scanned']}"
                + ("" if complete else " ⚠ выдача недобрана, WB не отдал страницу")
                + f", конкурентов взято {st['competitors']}"
                + (" (выдача кончилась раньше top_n)" if st["short"] else "")
                + f"; отброшено: другая категория {st['dropped']['другая категория']}, "
                f"лимит на продавца {st['dropped']['лимит на продавца']}, "
                f"наши {st['dropped']['наши']}; наши карточки: {found}")
        sp.log("  " + line)
        report.append(line)

        if not group["ours"]:
            report.append(f"«{query}»: пропущен — в выдаче нет ни одной нашей карточки "
                          f"(задайте артикулы вручную в колонке «Наши артикулы»)")
            continue
        if not group["competitors"]:
            report.append(f"«{query}»: пропущен — конкурентов не набралось")
            continue
        groups.append(group)

    if max_shelves:
        uniq = sorted({c for g in groups for c in g["competitors"]})
        if len(uniq) > max_shelves:
            # Режем поровну по запросам и с головы выдачи: терять надо хвост
            # топа, а не целый запрос.
            keep: set[int] = set()
            per_query = max(1, max_shelves // len(groups))
            for g in groups:
                keep.update(g["competitors"][:per_query])
            for g in groups:                       # остаток бюджета — по кругу
                for c in g["competitors"]:
                    if len(keep) >= max_shelves:
                        break
                    keep.add(c)
            cut = 0
            for g in groups:
                before = len(g["competitors"])
                g["competitors"] = [c for c in g["competitors"] if c in keep]
                cut += before - len(g["competitors"])
            msg = (f"Ограничение --max-shelves={max_shelves}: из {len(uniq)} полок "
                   f"отброшено {cut} строк выдачи (равномерно по запросам)")
            sp.log(msg)
            report.append(msg)

    return groups, report


def brand_of_nm(group: dict, nm: int) -> str:
    return (group.get("brand_of") or {}).get(str(nm), "")
