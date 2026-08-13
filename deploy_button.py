#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Заливка кнопки в книги брендов (Apps Script через clasp).

Две цели заливки (`--target`):

    shelves  книга ПОЛОК, меню «Полки WB» → «Обновить полки» / «Обновить цены» /
             «Обновить всё» (`apps_script/Code.gs`);
    prices   книга ЦЕН, меню «Цены WB» → «Обновить цены»
             (`apps_script/CodePrices.gs`, с 13.08.2026).

Обе кнопки дёргают один workflow `brand-shelves.yml` со своим брендом и входом
`task`; книга цен умеет только `prices` — полок в ней нет.

Токен GitHub в репозиторий не попадает: в файлах `apps_script/*.gs` лежит
плейсхолдер `__GH_TOKEN__`, значение подставляется ЗДЕСЬ из локального файла
(по умолчанию `21.orders-cloud/github_token.txt`) во временную папку, откуда и
идёт `clasp push`.

Запуск (нужен установленный и залогиненный clasp):
    python deploy_button.py                      # книги полок, все бренды
    python deploy_button.py --target prices      # книги цен, все бренды
    python deploy_button.py --brand SUNSHINE
    python deploy_button.py --update             # перезалить код в уже созданные

Соответствие «бренд → scriptId» пишется в `apps_script/deployed.json` (полки) и
`apps_script/deployed_prices.json` (цены); обе карты в .gitignore — это карта
чужих проектов, не код. Скрипт, заведённый человеком прямо в таблице
(Расширения → Apps Script), достаточно вписать в карту руками: заливка пойдёт
в него, нового проекта не создастся.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

from brand_shelves import BRANDS

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "apps_script")
DEFAULT_TOKEN_FILE = os.path.join(
    HERE, "..", "21.orders-cloud", "github_token.txt")

# Цель заливки: своя книга, свой файл кода, своя карта scriptId и свой заголовок
# проекта. Книга цен и книга полок — разные таблицы одного бренда.
TARGETS = {
    "shelves": {"code": "Code.gs", "state": "deployed.json",
                "book": "sheet_id", "title": "Полки {brand}"},
    "prices": {"code": "CodePrices.gs", "state": "deployed_prices.json",
               "book": "prices_book", "title": "Цены {brand}"},
}


def read_token(path: str) -> str:
    with open(path, encoding="utf-8") as f:
        raw = f.read().strip()
    # В файле токен лежит в виде «Bearer github_pat_…» — берём само значение.
    return raw.split()[-1]


def load_state(target: dict) -> dict:
    path = os.path.join(SRC, target["state"])
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_state(state: dict, target: dict) -> None:
    with open(os.path.join(SRC, target["state"]), "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def clasp(args: list[str], cwd: str) -> str:
    cmd = ["clasp"] + args
    res = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True,
                         encoding="utf-8", errors="replace",
                         shell=(os.name == "nt"))
    out = (res.stdout or "") + (res.stderr or "")
    if res.returncode != 0:
        raise SystemExit(f"clasp {' '.join(args)} → код {res.returncode}\n{out}")
    return out


def deploy(brand: str, token: str, state: dict, update_only: bool,
           target: dict) -> None:
    tmp = tempfile.mkdtemp(prefix=f"shelfbtn-{brand.replace(' ', '_')}-")
    try:
        code = open(os.path.join(SRC, target["code"]), encoding="utf-8").read()
        code = code.replace("__BRAND__", brand).replace("__GH_TOKEN__", token)
        # В проекте книги файл всегда один и называется Code.gs: имя файла в
        # репозитории (Code.gs / CodePrices.gs) — это выбор цели, а не имя,
        # под которым код лежит в таблице.
        with open(os.path.join(tmp, "Code.gs"), "w", encoding="utf-8") as f:
            f.write(code)
        shutil.copy(os.path.join(SRC, "appsscript.json"), tmp)

        script_id = state.get(brand)
        if script_id:
            with open(os.path.join(tmp, ".clasp.json"), "w", encoding="utf-8") as f:
                json.dump({"scriptId": script_id, "rootDir": tmp}, f)
        else:
            if update_only:
                print(f"{brand}: скрипта ещё нет, пропускаю (--update)")
                return
            parent = BRANDS[brand][target["book"]]
            if not parent:
                raise SystemExit(f"{brand}: в brands.json не заведена книга "
                                 f"({target['book']}) — некуда ставить кнопку")
            out = clasp(["create-script", "--type", "sheets",
                         "--title", target["title"].format(brand=brand),
                         "--parentId", parent,
                         "--rootDir", tmp], cwd=tmp)
            m = re.search(r"[-\w]{25,}", out.replace("\n", " "))
            cfg = os.path.join(tmp, ".clasp.json")
            if os.path.exists(cfg):
                script_id = json.load(open(cfg, encoding="utf-8")).get("scriptId")
            script_id = script_id or (m.group(0) if m else None)
            if not script_id:
                raise SystemExit(f"{brand}: не понял scriptId из ответа clasp:\n{out}")
            state[brand] = script_id
            save_state(state, target)

        clasp(["push", "--force"], cwd=tmp)
        print(f"{brand}: код залит, scriptId {state[brand]}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main() -> None:
    ap = argparse.ArgumentParser(description="Кнопка обновления в книги брендов")
    ap.add_argument("--brand", default="all")
    ap.add_argument("--target", default="shelves", choices=sorted(TARGETS),
                    help="shelves — книга полок, prices — книга цен")
    ap.add_argument("--token-file", default=DEFAULT_TOKEN_FILE)
    ap.add_argument("--update", action="store_true",
                    help="только перезалить код в уже созданные проекты")
    args = ap.parse_args()
    target = TARGETS[args.target]

    token = read_token(args.token_file)
    if not token.startswith(("github_pat_", "ghp_", "ghs_")):
        raise SystemExit(f"В {args.token_file} не похоже на токен GitHub")

    brands = list(BRANDS) if args.brand.lower() == "all" else [
        b for b in BRANDS if b.lower() == args.brand.strip().lower()]
    if not brands:
        raise SystemExit(f"Неизвестный бренд {args.brand!r}")

    state = load_state(target)
    for brand in brands:
        deploy(brand, token, state, args.update, target)
    menu = "Полки WB" if args.target == "shelves" else "Цены WB"
    # Без стрелок и многоточий: консоль Windows живёт в cp1251 и падает на них
    # уже после того, как код залит (13.08.2026).
    print(f"\nДальше - один раз в каждой книге: обновить страницу, меню "
          f"«{menu}», пункт обновления, разрешить скрипту доступ.")


if __name__ == "__main__":
    main()
