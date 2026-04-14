"""
Google Calendar API v3 client.
Docs: https://developers.google.com/calendar/api/v3/reference
"""

import json
import re
import requests
from datetime import datetime, timezone, timedelta
from typing import Optional


GCAL_API_BASE = "https://www.googleapis.com/calendar/v3"
CALENDAR_ID = "primary"


# -----------------------------------------------------------------
# Datetime parsing helpers
# -----------------------------------------------------------------

def _parse_datetime(text: str) -> Optional[dict]:
    """
    Try to extract a datetime (or date) from Korean natural language text.
    Returns a Google Calendar 'start'/'end' dict, or None if not found.

    Examples handled:
      "내일 오후 3시"       → tomorrow 15:00
      "오늘 오전 10시 30분" → today 10:30
      "4월 20일 오후 2시"   → April 20 14:00
      "2026-04-20 14:00"   → explicit datetime
    """
    now = datetime.now(timezone.utc).astimezone()
    today = now.date()

    # --- Resolve base date ---
    base_date = today

    if "내일" in text:
        base_date = today + timedelta(days=1)
    elif "모레" in text:
        base_date = today + timedelta(days=2)
    elif "오늘" in text:
        base_date = today
    else:
        # Try "N월 M일"
        m = re.search(r"(\d{1,2})월\s*(\d{1,2})일", text)
        if m:
            month, day = int(m.group(1)), int(m.group(2))
            base_date = today.replace(month=month, day=day)
        else:
            # Try YYYY-MM-DD
            m = re.search(r"(\d{4})-(\d{2})-(\d{2})", text)
            if m:
                base_date = datetime.strptime(m.group(0), "%Y-%m-%d").date()

    # --- Resolve time ---
    hour: Optional[int] = None
    minute: int = 0

    # "오후 N시 [M분]" or "오전 N시 [M분]"
    m = re.search(r"(오전|오후)\s*(\d{1,2})시(?:\s*(\d{1,2})분)?", text)
    if m:
        am_pm, h_str, min_str = m.group(1), m.group(2), m.group(3)
        hour = int(h_str)
        if am_pm == "오후" and hour != 12:
            hour += 12
        elif am_pm == "오전" and hour == 12:
            hour = 0
        if min_str:
            minute = int(min_str)
    else:
        # Try bare "HH:MM" or "HH시"
        m = re.search(r"(\d{1,2}):(\d{2})", text)
        if m:
            hour, minute = int(m.group(1)), int(m.group(2))
        else:
            m = re.search(r"(\d{1,2})시", text)
            if m:
                hour = int(m.group(1))

    if hour is not None:
        start_dt = datetime(
            base_date.year, base_date.month, base_date.day,
            hour, minute, 0,
        )
        end_dt = start_dt + timedelta(hours=1)
        tz = now.strftime("%z")           # e.g. "+0900"
        tz_fmt = f"{tz[:3]}:{tz[3:]}"    # "+09:00"
        return {
            "start": {"dateTime": start_dt.strftime("%Y-%m-%dT%H:%M:%S") + tz_fmt},
            "end":   {"dateTime": end_dt.strftime("%Y-%m-%dT%H:%M:%S") + tz_fmt},
        }

    # All-day event if only a date was found
    if base_date != today or any(kw in text for kw in ("내일", "모레", "오늘")):
        return {
            "start": {"date": str(base_date)},
            "end":   {"date": str(base_date + timedelta(days=1))},
        }

    return None


# -----------------------------------------------------------------
# API call
# -----------------------------------------------------------------

def create_event(token: str, summary: str) -> dict:
    """Create a Google Calendar event. Returns the created event dict."""
    timing = _parse_datetime(summary)

    if timing is None:
        # Fallback: all-day event today
        today_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        timing = {
            "start": {"date": today_str},
            "end":   {"date": today_str},
        }

    payload = {
        "summary": summary,
        **timing,
    }

    resp = requests.post(
        f"{GCAL_API_BASE}/calendars/{CALENDAR_ID}/events",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        data=json.dumps(payload),
        timeout=10,
    )
    resp.raise_for_status()
    return resp.json()
