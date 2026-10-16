"""JSON storage with atomic writes (same pattern as desktop)."""
from __future__ import annotations

import json
import os
import shutil
import tempfile
from datetime import datetime, timedelta, timezone
from typing import Any

MSK = timezone(timedelta(hours=3))


def load_json(path: str, fallback: Any) -> Any:
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        bak = path + ".bak"
        try:
            if os.path.exists(bak):
                with open(path + ".bak", "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            pass
    return fallback


def save_json(path: str, data: Any) -> None:
    directory = os.path.dirname(path) or "."
    os.makedirs(directory, exist_ok=True)
    if os.path.exists(path):
        try:
            shutil.copy2(path, path + ".bak")
        except Exception:
            pass
    fd, tmp = tempfile.mkstemp(prefix=".wrs_", suffix=".tmp", dir=directory)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
    except Exception:
        try:
            if os.path.exists(tmp):
                os.unlink(tmp)
        except OSError:
            pass
        raise


def default_settings() -> dict:
    return {
        "version": 9,
        "ignore_city_contains": ["москва"],
        "managers": [
            "Артамонова Виктория",
            "Давыдова Оксана",
            "Морозова Мария",
            "Сизенов Роман",
            "Ольга Старых",
            "Найденова Наталия",
            "Спицына Екатерина",
            "Стародубцева Евгения",
            "Филимонова Юлия",
        ],
        "regions": {
            "Орёл": [
                "Артамонова Виктория",
                "Давыдова Оксана",
                "Морозова Мария",
                "Сизенов Роман",
                "Ольга Старых",
            ],
            "Курск": [
                "Найденова Наталия",
                "Спицына Екатерина",
                "Стародубцева Евгения",
            ],
            "Липецк": ["Филимонова Юлия"],
        },
    }


def now_str() -> str:
    """Московское время (UTC+3)."""
    return datetime.now(MSK).strftime("%d.%m.%Y %H:%M")
