#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Центральные настройки книг брендов — `brands.json`.

Зачем: до 10.08.2026 ID книг и глубина истории были зашиты константами в
`brand_shelves.py`, а книга цен вообще не существовала как понятие (лист «Цены»
лежал внутри книги полок). Пятый бренд или переезд таблицы требовали правки
кода в трёх местах. Теперь всё в одном json, а модули берут его отсюда.

Что даёт:
    BRANDS            — {бренд: {"aliases", "sheet_id" (книга полок),
                        "prices_book", "prices"}}. Ключ `sheet_id` сохранён,
                        чтобы не переписывать работающий разбор полок.
    brand_cfg(name)    — настройки бренда по неточному имени ("4me" → "4ME").
    books_with_prices()— бренды, у которых включён сбор цен.
    CFG                — весь json целиком (dest, keep_dates, telegram, …).

Переопределение ID книги окружением (для отладки на копии таблицы):
    SHEET_ID_<БРЕНД>          — книга полок,   напр. SHEET_ID_4ME
    PRICES_ID_<БРЕНД>         — книга цен,     напр. PRICES_ID_HEALTH_FORM
Путь к самому файлу настроек — BRANDS_CONFIG (по умолчанию рядом со скриптом).
"""

from __future__ import annotations

import json
import os

CONFIG_PATH = os.environ.get(
    "BRANDS_CONFIG",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "brands.json"))


def _env_key(brand: str) -> str:
    return brand.upper().replace(" ", "_").replace("-", "_")


def _load() -> dict:
    with open(CONFIG_PATH, encoding="utf-8") as f:
        cfg = json.load(f)
    if not cfg.get("brands"):
        raise SystemExit(f"В {CONFIG_PATH} нет ни одного бренда")
    return cfg


CFG: dict = _load()

DEST: int = int(CFG.get("dest", -1257786))
KEEP_DATES: int = int(CFG.get("keep_dates", 90))
GRADIENT_DATES: int = int(CFG.get("gradient_dates", KEEP_DATES))
TEMPLATE_SHEET_ID: str = CFG.get("template", {}).get("sheet_id", "")
TEMPLATE_TAB: str = CFG.get("template", {}).get("tab", "")
SYNC_FROM: str = CFG.get("sync", {}).get("from", "NATURI")
SYNC_SKIP: list[str] = list(CFG.get("sync", {}).get("skip", []))
TELEGRAM: dict = CFG.get("telegram", {})

# Книга полок держит исторический ключ `sheet_id`: на него завязан разбор полок,
# и переименование ничего бы не улучшило.
BRANDS: dict[str, dict] = {}
for _name, _b in CFG["brands"].items():
    _k = _env_key(_name)
    BRANDS[_name] = {
        "aliases": list(_b.get("aliases") or [_name]),
        "sheet_id": os.environ.get(f"SHEET_ID_{_k}") or _b.get("shelves_book", ""),
        "prices_book": os.environ.get(f"PRICES_ID_{_k}") or _b.get("prices_book", ""),
        "prices": bool(_b.get("prices", False)) and bool(
            os.environ.get(f"PRICES_ID_{_k}") or _b.get("prices_book")),
        # v2.2.0. Бренд с собственным списком товаров (Upgrade) не должен
        # принимать группы из книги-донора: у него свои 19 групп и свои
        # конкуренты, а `--sync-groups --brand all` залил бы ему чужие 43.
        "sync": bool(_b.get("sync", True)),
    }


def brand_cfg(name: str) -> tuple[str, dict]:
    """Неточное имя бренда → (каноничное имя, настройки). Иначе SystemExit."""
    want = (name or "").strip().lower()
    for brand, cfg in BRANDS.items():
        if brand.lower() == want:
            return brand, cfg
    raise SystemExit(f"Неизвестный бренд {name!r}; известные: {', '.join(BRANDS)}")


def books_with_prices() -> list[str]:
    """Бренды, у которых заведена книга цен."""
    return [b for b, cfg in BRANDS.items() if cfg["prices"]]


def prices_book(brand: str) -> str:
    _, cfg = brand_cfg(brand)
    if not cfg["prices_book"]:
        raise SystemExit(f"У бренда {brand} в brands.json не заведена книга цен "
                         "(поле prices_book)")
    return cfg["prices_book"]
