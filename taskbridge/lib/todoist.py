"""
Todoist API v1 client.
Docs: https://developer.todoist.com/api/v1
"""

import json
import re
import requests
from datetime import datetime, timezone, timedelta
from typing import Optional

from lib.google_cal import strip_datetime, _parse_datetime


TODOIST_API_BASE = "https://api.todoist.com/api/v1"


def _extract_due_date(text: str) -> Optional[str]:
    """
    날짜 표현을 추출해 ISO date string (YYYY-MM-DD) 반환.
    google_cal의 _parse_datetime을 재사용.
    """
    today = datetime.now(timezone.utc).date()
    timing = _parse_datetime(text)
    if not timing:
        return None

    start = timing.get("start", {})
    if "date" in start:
        return start["date"]
    if "dateTime" in start:
        return start["dateTime"][:10]  # YYYY-MM-DD 부분만
    return None


def create_task(token: str, content: str) -> dict:
    """Create a Todoist task. Returns the created task dict."""
    clean_content = strip_datetime(content) or content  # 날짜 제거 후 빈 문자열이면 원본 사용
    payload: dict = {"content": clean_content}

    due_date = _extract_due_date(content)
    if due_date:
        payload["due_date"] = due_date

    resp = requests.post(
        f"{TODOIST_API_BASE}/tasks",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        data=json.dumps(payload),
        timeout=10,
    )
    if not resp.ok:
        print(f"[Todoist API error] status={resp.status_code} body={resp.text}")
        resp.raise_for_status()
    return resp.json()
