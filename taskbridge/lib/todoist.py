"""
Todoist REST API v2 client.
Docs: https://developer.todoist.com/rest/v2/
"""

import json
import re
import requests
from datetime import datetime, timezone
from typing import Optional


TODOIST_API_BASE = "https://api.todoist.com/rest/v2"


def _extract_due_date(text: str) -> Optional[str]:
    """
    Naively extract a due date string from text.
    Returns ISO date string (YYYY-MM-DD) if found, else None.
    More sophisticated NLP can be added later.
    """
    # Match patterns like "내일", "오늘", or explicit dates (YYYY-MM-DD / MM/DD)
    today = datetime.now(timezone.utc).date()

    if "내일" in text:
        from datetime import timedelta
        return str(today + timedelta(days=1))
    if "오늘" in text:
        return str(today)

    # Match YYYY-MM-DD
    m = re.search(r"\d{4}-\d{2}-\d{2}", text)
    if m:
        return m.group(0)

    # Match MM/DD or MM-DD
    m = re.search(r"(\d{1,2})[/-](\d{1,2})", text)
    if m:
        month, day = m.group(1), m.group(2)
        return f"{today.year}-{int(month):02d}-{int(day):02d}"

    return None


def create_task(token: str, content: str) -> dict:
    """Create a Todoist task. Returns the created task dict."""
    payload: dict = {"content": content}

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
    resp.raise_for_status()
    return resp.json()
