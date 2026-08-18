#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Разовая заводка листа «Полки» книги бренда из листа конкурентов юнитки кабинета.

v1.0.0 — 18.08.2026
  ШЕСТОЙ БРЕНД ORZAX (кабинет ИП1 Орзакс) НЕ ЛОЖИТСЯ НА ОБРАЗЕЦ «Натури пример»:
  у первых пяти книг конкурент — это ДРУГОЙ БРЕНД в той же нише, а у Orzax
  конкурент — ДРУГОЙ ПРОДАВЕЦ того же самого товара (перекупы возят ту же
  турецкую банку). Общего списка конкурентов с остальными книгами нет вообще,
  поэтому и донор товаров (`--sync-groups`) ему не указ.

  Источник строк — лист «Конкуренты» юнитки кабинета
  (`Ссылка | Название товара/ | Артикул | Продавец`), он же единственное место,
  где менеджер ведёт состав перекупов. Отсюда собирается лист «Полки» в обычной
  раскладке проекта 43, дальше книгу ведёт человек, а `brand_shelves.py`
  дописывает колонку за день ровно так же, как в остальных книгах.

  Наша карточка группы — строка НАШЕГО продавца (`our_seller` в `brands.json`).
  Она поднимается первой строкой группы и подписывается меткой бренда
  (`our_label`), иначе в колонке «Бренд конкурента», где у всех фамилии
  продавцов, свою строку глазами не найти. Заливка тёмно-зелёным — как у
  контрольных карточек в остальных книгах.

Скрипт разовый, в расписание не ставится: заведённый лист он не трогает
(кроме `--sync-rows`), чтобы повтор не затирал работу человека.

Запуск:
    python shelves_from_competitors.py --brand ORZAX --creds ключ.json
    python shelves_from_competitors.py --brand ORZAX --dry-run
    python shelves_from_competitors.py --brand ORZAX --sync-rows
"""

from __future__ import annotations

import argparse

import brand_shelves as bs
import shelf_positions as sp
import to_sheets as ts
import wb_config
from vexor_shelves import install_retries

SHEET_DST = bs.SHEET_DST                 # «Полки»

# Как названы колонки в листе конкурентов юнитки. Ищутся по шапке, а не по
# буквам: у кабинетов лист заведён руками, и порядок колонок может отличаться.
SRC_PRODUCT = "Название товара/"
SRC_ART = "Артикул"
SRC_SELLER = "Продавец"


def read_source(ws, our_seller: str) -> list[dict]:
    """Лист конкурентов → группы: [{"product", "our", "rows": [(артикул, продавец)]}].

    Порядок товаров и порядок продавцов внутри товара сохраняется — так менеджер
    ведёт список, и переставлять его самовольно нельзя. Наша строка из группы
    вынимается: в лист она встанет первой.
    """
    values = ws.get_all_values()
    if not values:
        raise SystemExit(f"Лист «{ws.title}» пуст")
    head = [str(x).strip() for x in values[0]]

    def col_of(title: str) -> int:
        for i, cell in enumerate(head):
            if cell.lower() == title.lower():
                return i
        raise SystemExit(f"В листе «{ws.title}» нет колонки «{title}»; "
                         f"шапка: {', '.join(h for h in head if h)}")

    ip, ia, isell = col_of(SRC_PRODUCT), col_of(SRC_ART), col_of(SRC_SELLER)

    groups: list[dict] = []
    index: dict[str, dict] = {}
    for raw in values[1:]:
        cells = [str(x).strip() for x in raw]
        if len(cells) <= max(ip, ia, isell):
            cells += [""] * (max(ip, ia, isell) + 1 - len(cells))
        product, art, seller = cells[ip], cells[ia], cells[isell]
        if not product or not art:
            continue
        g = index.get(product)
        if g is None:
            g = {"product": product, "our": None, "rows": []}
            index[product] = g
            groups.append(g)
        if seller.strip().lower() == our_seller.strip().lower() and g["our"] is None:
            g["our"] = (art, seller)
        else:
            g["rows"].append((art, seller))
    return groups


def build_body(groups: list[dict], our_label: str) -> tuple[list[list[str]], list[int]]:
    """Группы → строки листа «Полки» (+ номера строк наших карточек, 1-based)."""
    body: list[list[str]] = []
    ours: list[int] = []
    for n, g in enumerate(groups):
        if n:
            body.append([""] * bs.NFIX)          # пустая строка = граница группы
        # Карточки у нас нет — в артикуле «нет»: группа видна, но не мерится
        # (правило «чужой карточкой не мерим», решение 6 в СОСТОЯНИЕ.md `43`).
        art = g["our"][0] if g["our"] else "нет"
        body.append([g["product"], art, our_label, "", "", "", ""])
        ours.append(len(body) + 1)               # +1 на шапку
        for art, seller in g["rows"]:
            body.append([g["product"], art, seller, "", "", "", ""])
    return body, ours


def write_sheet(book, body: list[list[str]], ours: list[int], our_seller: str):
    """Завести лист «Полки»: шапка, ширины, заморозка, заливка наших строк."""
    ws = book.add_worksheet(title=SHEET_DST, rows=len(body) + 200, cols=bs.NFIX + 10)
    ws.update_index(0)
    ws.update(values=[bs.FIX_COLS] + body,
              range_name=f"A1:{bs.col_letter(bs.NFIX)}{len(body) + 1}",
              value_input_option="USER_ENTERED")

    reqs = [
        {"updateSheetProperties": {
            "properties": {"sheetId": ws.id,
                           "gridProperties": {"frozenRowCount": 1,
                                              "frozenColumnCount": 3}},
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}},
        {"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": bs.NFIX},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True},
                                           "wrapStrategy": "WRAP",
                                           "horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat(textFormat,wrapStrategy,horizontalAlignment)"}},
        # Колонка C в этой книге держит ПРОДАВЦА, а не бренд: пояснение висит
        # примечанием к шапке, чтобы через полгода никто не переписал её «как у
        # NATURI» и не сломал поиск нашей карточки.
        {"updateCells": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 2, "endColumnIndex": 3},
            "rows": [{"values": [{"note":
                      "Кабинет ИП1 Орзакс: конкурент — другой ПРОДАВЕЦ того же "
                      "товара Orzax, поэтому здесь фамилия продавца. Наша "
                      f"карточка (продавец {our_seller}) подписана меткой бренда "
                      "и выделена тёмно-зелёным — по ней скрипт и ищет позиции."}]}],
            "fields": "note"}},
    ]
    for i, w in enumerate(bs.COL_WIDTHS):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": i, "endIndex": i + 1},
            "properties": {"pixelSize": w}, "fields": "pixelSize"}})
    for row in ours:
        reqs.append({"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": row - 1, "endRowIndex": row,
                      "startColumnIndex": 0, "endColumnIndex": bs.NFIX},
            "cell": {"userEnteredFormat": {"backgroundColor": bs.COLOR_CONTROL}},
            "fields": "userEnteredFormat.backgroundColor"}})

    for i in range(0, len(reqs), 300):
        book.batch_update({"requests": reqs[i:i + 300]})
    return ws


def sync_rows(book, groups: list[dict], our_label: str) -> str:
    """Дописать в конец листа товары, которых в нём ещё нет. Прежние не трогать."""
    ws = book.worksheet(SHEET_DST)
    values = ws.get_all_values()
    head = [str(x).strip() for x in values[0]]
    lay = bs.read_layout(head)
    have = {r[lay["product"]].strip() for r in values[1:]
            if len(r) > lay["product"] and r[lay["product"]].strip()}

    fresh = [g for g in groups if g["product"] not in have]
    if not fresh:
        return "новых товаров в источнике нет"

    body, _ = build_body(fresh, our_label)
    start = len(values) + 2                       # пустая строка-разделитель
    ws.update(values=body,
              range_name=f"A{start}:{bs.col_letter(bs.NFIX)}{start + len(body) - 1}",
              value_input_option="USER_ENTERED")
    return (f"дописано товаров {len(fresh)}, строк {len(body)} "
            f"({', '.join(g['product'] for g in fresh[:3])}"
            f"{'…' if len(fresh) > 3 else ''})")


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Лист «Полки» книги бренда из листа конкурентов юнитки кабинета")
    ap.add_argument("--brand", required=True)
    ap.add_argument("--creds", default=None)
    ap.add_argument("--dry-run", action="store_true", help="в таблицу не писать")
    ap.add_argument("--sync-rows", action="store_true",
                    help="дописать в существующий лист недостающие товары")
    args = ap.parse_args()

    brand, cfg = wb_config.brand_cfg(args.brand)
    src = cfg.get("source") or {}
    if not src.get("sheet_id"):
        raise SystemExit(f"У бренда {brand} в brands.json нет блока source "
                         "(sheet_id / tab / our_seller) — неоткуда брать строки")

    install_retries()
    client = ts.get_client(args.creds)

    src_ws = client.open_by_key(src["sheet_id"]).worksheet(src.get("tab", "Конкуренты"))
    groups = read_source(src_ws, src["our_seller"])
    no_card = [g["product"] for g in groups if not g["our"]]
    sp.log(f"Источник «{src_ws.title}»: товаров {len(groups)}, "
           f"строк {sum(1 + len(g['rows']) for g in groups)}, "
           f"полок к обходу {sum(len(g['rows']) for g in groups)}")
    if no_card:
        sp.log(f"Без нашей карточки (мериться не будут): {'; '.join(no_card)}")

    if args.dry_run:
        for g in groups:
            print(f"  {g['product']}: наша {g['our'][0] if g['our'] else 'нет'}, "
                  f"полок {len(g['rows'])}")
        return

    book = client.open_by_key(cfg["sheet_id"])
    if SHEET_DST in [w.title for w in book.worksheets()]:
        if args.sync_rows:
            print("Готово:", f"{brand}: " + sync_rows(book, groups, src["our_label"]))
            return
        raise SystemExit(f"{brand}: лист «{SHEET_DST}» уже есть — трогать не буду "
                         "(--sync-rows дополнит его недостающими товарами)")

    body, ours = build_body(groups, src["our_label"])
    write_sheet(book, body, ours, src["our_seller"])
    # Пустой «Лист1», который Google создаёт вместе с книгой, только мешает.
    for ws in book.worksheets():
        if ws.title in ("Лист1", "Sheet1") and len(book.worksheets()) > 1:
            if not any(any(c.strip() for c in row) for row in ws.get_all_values()):
                book.del_worksheet(ws)
                sp.log(f"Удалён пустой лист «{ws.title}»")
    print("Готово:", f"{brand}: лист «{SHEET_DST}» заведён — строк {len(body)}, "
          f"групп {len(groups)}, наших карточек {len(ours) - len(no_card)}")


if __name__ == "__main__":
    main()
