"""
Todoist API v1 client.
Docs: https://developer.todoist.com/api/v1

인라인 메타데이터 문법:
  #프로젝트명   → 해당 프로젝트에 저장
  @레이블       → 레이블 지정 (여러 개 가능)
  !p1 ~ !p4    → 우선순위 (p1=긴급, p4=보통)
"""

import json
import re
import requests
from datetime import datetime, timezone
from typing import Optional

from lib.google_cal import strip_datetime, _parse_datetime


TODOIST_API_BASE = "https://api.todoist.com/api/v1"


# ---------------------------------------------------------------
# 메타데이터 파서
# ---------------------------------------------------------------

def parse_meta(text: str) -> dict:
    """
    텍스트에서 Todoist 메타데이터를 파싱한다.
    반환: {
        'content' : 메타데이터/날짜 제거 후 깨끗한 텍스트,
        'project' : 프로젝트 이름 (str or None),
        'labels'  : 레이블 이름 리스트,
        'priority': Todoist 우선순위 정수 (4=긴급 … 1=보통) or None,
    }
    """
    content = text

    # 우선순위: !p1~!p4 또는 !1~!4
    priority = None
    m = re.search(r'!p?([1-4])\b', content, re.IGNORECASE)
    if m:
        user_p   = int(m.group(1))
        priority = 5 - user_p      # p1→4(긴급), p2→3, p3→2, p4→1(보통)
        content  = content[:m.start()] + content[m.end():]

    # 프로젝트: #이름 (공백 없는 단어)
    project = None
    m = re.search(r'#(\S+)', content)
    if m:
        project = m.group(1)
        content = content[:m.start()] + content[m.end():]

    # 레이블: @이름 (여러 개 가능)
    labels = re.findall(r'@(\S+)', content)
    content = re.sub(r'@\S+', '', content)

    # 날짜/시간 제거
    content = strip_datetime(content)

    # 공백 정리
    content = re.sub(r'\s+', ' ', content).strip()

    return {
        'content' : content,
        'project' : project,
        'labels'  : labels,
        'priority': priority,
    }


# ---------------------------------------------------------------
# 프로젝트 이름 → ID 조회
# ---------------------------------------------------------------

def _get_project_id(token: str, name: str) -> Optional[str]:
    """프로젝트 이름으로 ID를 조회한다. 없으면 None."""
    resp = requests.get(
        f"{TODOIST_API_BASE}/projects",
        headers={"Authorization": f"Bearer {token}"},
        timeout=10,
    )
    if not resp.ok:
        print(f"[Todoist] 프로젝트 목록 조회 실패: {resp.status_code}")
        return None

    for project in resp.json():
        if project.get("name", "").lower() == name.lower():
            return project["id"]

    print(f"[Todoist] 프로젝트 '{name}' 없음")
    return None


# ---------------------------------------------------------------
# 날짜 추출
# ---------------------------------------------------------------

def _extract_due_date(text: str) -> Optional[str]:
    timing = _parse_datetime(text)
    if not timing:
        return None
    start = timing.get("start", {})
    if "date" in start:
        return start["date"]
    if "dateTime" in start:
        return start["dateTime"][:10]
    return None


# ---------------------------------------------------------------
# 태스크 생성
# ---------------------------------------------------------------

def create_task(token: str, content: str) -> dict:
    """Create a Todoist task. Returns the created task dict."""
    meta = parse_meta(content)

    payload: dict = {"content": meta["content"] or content}

    # 날짜
    due_date = _extract_due_date(content)
    if due_date:
        payload["due_date"] = due_date

    # 우선순위
    if meta["priority"] is not None:
        payload["priority"] = meta["priority"]

    # 레이블
    if meta["labels"]:
        payload["labels"] = meta["labels"]

    # 프로젝트
    if meta["project"]:
        project_id = _get_project_id(token, meta["project"])
        if project_id:
            payload["project_id"] = project_id

    print(f"[Todoist] payload={json.dumps(payload, ensure_ascii=False)}")
    resp = requests.post(
        f"{TODOIST_API_BASE}/tasks",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type":  "application/json",
        },
        data=json.dumps(payload),
        timeout=10,
    )
    if not resp.ok:
        print(f"[Todoist API error] status={resp.status_code} body={resp.text}")
        resp.raise_for_status()
    return resp.json()
