#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Оповещения в Telegram.

Получатель — секреты репозитория `TELEGRAM_BOT_TOKEN` и `TELEGRAM_CHAT_ID`
(формат `chat_id` либо `chat_id:топик`, несколько через запятую). Нет секретов —
модуль молча ничего не шлёт и возвращает False: локальный прогон и отладка не
должны падать из-за отсутствия бота.

Договор проекта (10.08.2026): итог раз в сутки после ежедневного прогона
(`digest.py`) плюс ошибка — сразу отдельным сообщением. Интрадей-прогоны цен
молчат, пока всё хорошо.

Грабли, уже оплаченные в других проектах (см. `_memory/PATTERNS.md`):
    • топик «General» форум-группы имеет thread_id=1, но передавать эту 1 в API
      НЕЛЬЗЯ — sendMessage отвечает 400 «message thread not found». Трактуем
      `:1` как «без топика»;
    • сообщение длиннее 4096 символов Telegram не принимает — режем по строкам;
    • падение отправки не должно ронять прогон, который уже сделал работу.
"""

from __future__ import annotations

import json
import os
import urllib.error
import urllib.request

API = "https://api.telegram.org/bot{token}/sendMessage"
LIMIT = 3900              # запас к лимиту 4096 на служебные хвосты


def _targets() -> list[tuple[str, int | None]]:
    raw = (os.environ.get("TELEGRAM_CHAT_ID") or "").strip()
    out: list[tuple[str, int | None]] = []
    for part in raw.split(","):
        part = part.strip()
        if not part:
            continue
        chat, _, topic = part.partition(":")
        thread = None
        if topic.strip().isdigit() and topic.strip() != "1":
            thread = int(topic.strip())
        out.append((chat.strip(), thread))
    return out


def _chunks(text: str) -> list[str]:
    if len(text) <= LIMIT:
        return [text]
    parts, cur = [], ""
    for line in text.split("\n"):
        if len(cur) + len(line) + 1 > LIMIT:
            parts.append(cur.rstrip())
            cur = ""
        cur += line + "\n"
    if cur.strip():
        parts.append(cur.rstrip())
    return parts


def send(text: str, silent: bool = False) -> bool:
    """Отправить сообщение. Вернёт True, если ушло хотя бы одному получателю."""
    token = (os.environ.get("TELEGRAM_BOT_TOKEN") or "").strip()
    targets = _targets()
    if not token or not targets:
        return False

    ok = False
    for chat, thread in targets:
        for chunk in _chunks(text):
            payload: dict = {"chat_id": chat, "text": chunk,
                             "disable_web_page_preview": True,
                             "disable_notification": silent}
            if thread:
                payload["message_thread_id"] = thread
            body = json.dumps(payload).encode("utf-8")
            req = urllib.request.Request(
                API.format(token=token), data=body,
                headers={"Content-Type": "application/json"})
            try:
                with urllib.request.urlopen(req, timeout=30) as resp:
                    ok = ok or json.loads(resp.read().decode("utf-8")).get("ok", False)
            except urllib.error.HTTPError as exc:
                # Тело ошибки Telegram полезнее её кода — печатаем, но не падаем.
                detail = exc.read().decode("utf-8", "replace")[:200]
                print(f"Telegram HTTP {exc.code} для {chat}: {detail}")
            except Exception as exc:                    # noqa: BLE001
                print(f"Telegram недоступен для {chat}: "
                      f"{exc.__class__.__name__}: {exc}")
    return ok


def fail(title: str, detail: str = "") -> bool:
    """Сообщение о падении — уходит сразу, со звуком."""
    text = f"❌ {title}"
    if detail:
        text += f"\n\n{detail.strip()[:1500]}"
    return send(text)


def _main() -> None:
    import argparse
    ap = argparse.ArgumentParser(description="Отправить сообщение в Telegram")
    ap.add_argument("text", nargs="*", help="текст; пусто — читаю stdin")
    ap.add_argument("--fail-title", default=None,
                    help="оформить как сообщение о падении с этим заголовком")
    args = ap.parse_args()

    import sys
    text = " ".join(args.text) if args.text else sys.stdin.read()
    ok = fail(args.fail_title, text) if args.fail_title else send(text)
    print("Отправлено" if ok else "Не отправлено (нет секретов или ошибка)")


if __name__ == "__main__":
    _main()
