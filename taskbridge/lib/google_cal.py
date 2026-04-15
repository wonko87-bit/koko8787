"""
Google Calendar API v3 client.
Docs: https://developers.google.com/calendar/api/v3/reference

지원하는 날짜/시간 표현:
  - 오늘 / 내일 / 모레
  - N월 M일 / M일 (월 생략 시 당월)
  - 오전/오후 N시 [M분] / N시 [M분] / HH:MM
  - 범위: [시작] ~ [끝]  (끝 부분에서 날짜 생략 시 시작 날짜 사용)

예시:
  "오늘 오후 3시 ~ 5시"           → 오늘 15:00 ~ 17:00
  "5월 7일 오전 10시 ~ 오후 1시"  → 5/7 10:00 ~ 13:00
  "내일 오후 2시 ~ 3월 8일 오후 4시" → 내일 14:00 ~ 3/8 16:00
  "5월 7일 테스트"                → 5/7 종일 이벤트
  "보고서 초안 작성"               → 날짜 없음 → None 반환
"""

import json
import re
import requests
from datetime import datetime, timezone, timedelta, date as date_type
from typing import Optional


GCAL_API_BASE = "https://www.googleapis.com/calendar/v3"
CALENDAR_ID   = "primary"


# ---------------------------------------------------------------
# Internal: single datetime parser
# ---------------------------------------------------------------

def _parse_single(
    text: str,
    default_date: Optional[date_type] = None,
) -> tuple[date_type, Optional[int], int]:
    """
    텍스트에서 날짜 + 시각을 파싱한다.
    반환: (date, hour_or_None, minute)
    - hour가 None이면 종일(all-day) 의미
    - default_date: 날짜 표현이 없을 때 사용할 기본 날짜
    """
    now   = datetime.now(timezone.utc).astimezone()
    today = now.date()
    base  = default_date or today

    # ---- 날짜 파싱 ----
    if "내일" in text:
        base = today + timedelta(days=1)
    elif "모레" in text:
        base = today + timedelta(days=2)
    elif "오늘" in text:
        base = today
    else:
        # N월 M일
        m = re.search(r"(\d{1,2})월\s*(\d{1,2})일", text)
        if m:
            month, day = int(m.group(1)), int(m.group(2))
            try:
                base = today.replace(month=month, day=day)
            except ValueError:
                pass
        else:
            # M일  (월 생략 → 당월)
            m = re.search(r"(?<!\d)(\d{1,2})일(?!\w)", text)
            if m:
                day = int(m.group(1))
                try:
                    base = today.replace(day=day)
                except ValueError:
                    pass
            else:
                # YYYY-MM-DD
                m = re.search(r"(\d{4})-(\d{2})-(\d{2})", text)
                if m:
                    base = date_type(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                # else: default_date 유지

    # ---- 시각 파싱 ----
    hour: Optional[int] = None
    minute: int = 0

    # 오전/오후 N시 [M분]
    m = re.search(r"(오전|오후)\s*(\d{1,2})시(?:\s*(\d{1,2})분)?", text)
    if m:
        am_pm   = m.group(1)
        hour    = int(m.group(2))
        minute  = int(m.group(3)) if m.group(3) else 0
        if am_pm == "오후" and hour != 12:
            hour += 12
        elif am_pm == "오전" and hour == 12:
            hour = 0
    else:
        # HH:MM
        m = re.search(r"(\d{1,2}):(\d{2})", text)
        if m:
            hour   = int(m.group(1))
            minute = int(m.group(2))
        else:
            # N시 [M분]  (오전/오후 없이)
            m = re.search(r"(?<!\d)(\d{1,2})시(?:\s*(\d{1,2})분)?", text)
            if m:
                hour   = int(m.group(1))
                minute = int(m.group(2)) if m.group(2) else 0

    return base, hour, minute


def _has_date_expr(text: str) -> bool:
    """텍스트에 날짜 표현이 있는지 확인."""
    return bool(
        re.search(r"오늘|내일|모레|\d{1,2}월\s*\d{1,2}일|(?<!\d)\d{1,2}일(?!\w)|\d{4}-\d{2}-\d{2}", text)
    )


def _tz_suffix() -> str:
    now = datetime.now(timezone.utc).astimezone()
    tz  = now.strftime("%z")          # "+0900"
    return f"{tz[:3]}:{tz[3:]}"       # "+09:00"


# ---------------------------------------------------------------
# Public: parse_datetime  (범위 포함)
# ---------------------------------------------------------------

def _parse_datetime(text: str) -> Optional[dict]:
    """
    텍스트에서 시작/종료 datetime을 파싱해 Google Calendar 형식으로 반환.
    날짜/시간 표현이 없으면 None 반환.
    """
    tz_fmt = _tz_suffix()

    # ---- 범위 처리: ~ 구분 ----
    if "~" in text:
        start_raw, _, end_raw = text.partition("~")
        start_raw = start_raw.strip()
        end_raw   = end_raw.strip()

        start_date, start_hour, start_min = _parse_single(start_raw)
        # 끝 부분 날짜 생략 → 시작 날짜를 기본값으로
        end_date,   end_hour,   end_min   = _parse_single(end_raw, default_date=start_date)

        if start_hour is not None and end_hour is not None:
            # 오후 맥락 전파: 시작이 오후이고 끝에 오전/오후 명시 없을 때
            # 끝 시각이 시작보다 작으면 PM으로 간주
            end_has_ampm = bool(re.search(r"오전|오후", end_raw))
            if not end_has_ampm and start_hour >= 12 and end_hour < start_hour and end_hour < 12:
                end_hour += 12

            s = datetime(start_date.year, start_date.month, start_date.day, start_hour, start_min)
            e = datetime(end_date.year,   end_date.month,   end_date.day,   end_hour,   end_min)
            return {
                "start": {"dateTime": s.strftime("%Y-%m-%dT%H:%M:%S") + tz_fmt},
                "end":   {"dateTime": e.strftime("%Y-%m-%dT%H:%M:%S") + tz_fmt},
            }

        if start_hour is not None:
            # 끝 시각 없음 → 시작+1시간
            s = datetime(start_date.year, start_date.month, start_date.day, start_hour, start_min)
            e = s + timedelta(hours=1)
            return {
                "start": {"dateTime": s.strftime("%Y-%m-%dT%H:%M:%S") + tz_fmt},
                "end":   {"dateTime": e.strftime("%Y-%m-%dT%H:%M:%S") + tz_fmt},
            }

        # 종일 범위
        if _has_date_expr(start_raw) or _has_date_expr(end_raw):
            return {
                "start": {"date": str(start_date)},
                "end":   {"date": str(end_date + timedelta(days=1))},
            }
        return None

    # ---- 단일 datetime ----
    base, hour, minute = _parse_single(text)

    if hour is not None:
        s = datetime(base.year, base.month, base.day, hour, minute)
        e = s + timedelta(hours=1)
        return {
            "start": {"dateTime": s.strftime("%Y-%m-%dT%H:%M:%S") + tz_fmt},
            "end":   {"dateTime": e.strftime("%Y-%m-%dT%H:%M:%S") + tz_fmt},
        }

    if _has_date_expr(text):
        return {
            "start": {"date": str(base)},
            "end":   {"date": str(base + timedelta(days=1))},
        }

    return None


# ---------------------------------------------------------------
# Public: strip_datetime
# ---------------------------------------------------------------

def strip_datetime(text: str) -> str:
    """날짜/시간/범위 표현을 텍스트에서 제거하고 깨끗한 내용만 반환."""
    result = text

    # ~ 이후 끝 시각 부분 제거  (예: "~ 5시", "~ 오후 5시", "~ 17:00")
    result = re.sub(
        r'\s*~\s*(?:(?:\d{1,2}월\s*\d{1,2}일\s*)?(?:오전|오후)?\s*\d{1,2}시(?:\s*\d{1,2}분)?'
        r'|(?:\d{1,2}월\s*\d{1,2}일\s*)?\d{1,2}:\d{2})',
        '', result
    )

    # 오전/오후 N시 [M분]
    result = re.sub(r'(오전|오후)\s*\d{1,2}시(?:\s*\d{1,2}분)?', '', result)
    # HH:MM
    result = re.sub(r'\d{1,2}:\d{2}', '', result)
    # N시 [M분]
    result = re.sub(r'(?<!\d)\d{1,2}시(?:\s*\d{1,2}분)?', '', result)

    # YYYY-MM-DD
    result = re.sub(r'\d{4}-\d{2}-\d{2}', '', result)
    # N월 M일
    result = re.sub(r'\d{1,2}월\s*\d{1,2}일', '', result)
    # M일 (단독)
    result = re.sub(r'(?<!\d)\d{1,2}일(?!\w)', '', result)

    # 한국어 날짜 키워드
    result = re.sub(r'오늘|내일|모레', '', result)

    # 공백 정리
    result = re.sub(r'\s+', ' ', result).strip()
    return result


# ---------------------------------------------------------------
# Public: create_event
# ---------------------------------------------------------------

def create_event(token: str, summary: str) -> dict:
    """Create a Google Calendar event. Returns the created event dict."""
    timing        = _parse_datetime(summary)
    clean_summary = strip_datetime(summary) or summary

    if timing is None:
        today_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        timing = {
            "start": {"date": today_str},
            "end":   {"date": today_str},
        }

    payload = {"summary": clean_summary, **timing}

    print(f"[Google Cal] payload={json.dumps(payload, ensure_ascii=False)}")
    resp = requests.post(
        f"{GCAL_API_BASE}/calendars/{CALENDAR_ID}/events",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type":  "application/json",
        },
        data=json.dumps(payload),
        timeout=10,
    )
    if not resp.ok:
        print(f"[Google Cal error] status={resp.status_code} body={resp.text}")
        resp.raise_for_status()
    result = resp.json()
    print(f"[Google Cal success] id={result.get('id')}")
    return result
