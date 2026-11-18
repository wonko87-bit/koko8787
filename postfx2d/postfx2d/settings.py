"""사용자 설정 저장 (프로그램을 다시 켜도 유지).

Windows: %APPDATA%\\PostFX2D\\settings.json   그 외: ~/.postfx2d/settings.json
"""
from __future__ import annotations

import json
import os


def path() -> str:
    base = os.environ.get("APPDATA")
    d = os.path.join(base, "PostFX2D") if base else os.path.join(os.path.expanduser("~"), ".postfx2d")
    return os.path.join(d, "settings.json")


def load() -> dict:
    try:
        with open(path(), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def save(data: dict) -> None:
    p = path()
    try:
        os.makedirs(os.path.dirname(p), exist_ok=True)
        tmp = p + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        os.replace(tmp, p)
    except OSError:
        pass


def update(**kw) -> None:
    d = load()
    for k, v in kw.items():
        if v is None:
            d.pop(k, None)
        else:
            d[k] = v
    save(d)
