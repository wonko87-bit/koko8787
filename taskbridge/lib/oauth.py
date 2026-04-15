"""
OAuth 2.0 helpers for Google and Todoist.

Both flows use the Authorization Code Grant.
Tokens are stored server-side in a simple in-memory dict keyed by session_id.
(For production use, replace with a persistent store.)
"""

import hashlib
import json
import os
import secrets
import time
import requests
from typing import Optional
from urllib.parse import urlencode


# -----------------------------------------------------------------
# In-memory stores
# -----------------------------------------------------------------

_sessions: dict[str, dict] = {}   # session_id → { google_token, todoist_token, ... }
_states:   dict[str, str]  = {}   # oauth state → session_id


# -----------------------------------------------------------------
# Session helpers
# -----------------------------------------------------------------

def new_session() -> str:
    sid = secrets.token_urlsafe(32)
    _sessions[sid] = {}
    return sid


def get_session(sid: str) -> Optional[dict]:
    return _sessions.get(sid)


def set_token(sid: str, provider: str, token_data: dict) -> None:
    if sid not in _sessions:
        _sessions[sid] = {}
    _sessions[sid][f"{provider}_token"] = token_data
    _sessions[sid][f"{provider}_token_expiry"] = time.time() + token_data.get("expires_in", 3600)


def get_access_token(sid: str, provider: str) -> Optional[str]:
    sess = _sessions.get(sid, {})
    td = sess.get(f"{provider}_token")
    if not td:
        return None
    expiry = sess.get(f"{provider}_token_expiry", 0)
    if time.time() > expiry - 60:
        # Token expired; try refresh
        refreshed = _refresh_token(provider, td.get("refresh_token"))
        if refreshed:
            set_token(sid, provider, refreshed)
            return refreshed["access_token"]
        return None
    return td.get("access_token")


def is_authenticated(sid: str) -> dict:
    """Return which providers are authenticated for this session."""
    sess = _sessions.get(sid, {})
    return {
        "google":   bool(sess.get("google_token")),
        "todoist":  bool(sess.get("todoist_token")),
    }


# -----------------------------------------------------------------
# Google OAuth 2.0
# -----------------------------------------------------------------

GOOGLE_AUTH_URL    = "https://accounts.google.com/o/oauth2/v2/auth"
GOOGLE_TOKEN_URL   = "https://oauth2.googleapis.com/token"
GOOGLE_SCOPES      = "https://www.googleapis.com/auth/calendar.events"


def google_auth_url(callback_base: str, session_id: str) -> str:
    state = secrets.token_urlsafe(16)
    _states[state] = session_id
    params = {
        "client_id":     os.environ["GOOGLE_CLIENT_ID"],
        "redirect_uri":  f"{callback_base}/auth/google/callback",
        "response_type": "code",
        "scope":         GOOGLE_SCOPES,
        "access_type":   "offline",
        "prompt":        "consent",
        "state":         state,
    }
    return f"{GOOGLE_AUTH_URL}?{urlencode(params)}"


def google_exchange_code(code: str, state: str, callback_base: str) -> Optional[str]:
    """Exchange authorization code for tokens. Returns session_id on success."""
    sid = _states.pop(state, None)
    if not sid:
        return None

    resp = requests.post(GOOGLE_TOKEN_URL, data={
        "code":          code,
        "client_id":     os.environ["GOOGLE_CLIENT_ID"],
        "client_secret": os.environ["GOOGLE_CLIENT_SECRET"],
        "redirect_uri":  f"{callback_base}/auth/google/callback",
        "grant_type":    "authorization_code",
    }, timeout=10)
    resp.raise_for_status()
    set_token(sid, "google", resp.json())
    return sid


# -----------------------------------------------------------------
# Todoist OAuth 2.0
# -----------------------------------------------------------------

TODOIST_AUTH_URL  = "https://todoist.com/oauth/authorize"
TODOIST_TOKEN_URL = "https://todoist.com/oauth/access_token"
TODOIST_SCOPES    = "task:add,data:read_write"


def todoist_auth_url(callback_base: str, session_id: str) -> str:
    state = secrets.token_urlsafe(16)
    _states[state] = session_id
    params = {
        "client_id":     os.environ["TODOIST_CLIENT_ID"],
        "scope":         TODOIST_SCOPES,
        "state":         state,
    }
    return f"{TODOIST_AUTH_URL}?{urlencode(params)}"


def todoist_exchange_code(code: str, state: str) -> Optional[str]:
    """Exchange authorization code for tokens. Returns session_id on success."""
    sid = _states.pop(state, None)
    if not sid:
        return None

    resp = requests.post(TODOIST_TOKEN_URL, data={
        "code":          code,
        "client_id":     os.environ["TODOIST_CLIENT_ID"],
        "client_secret": os.environ["TODOIST_CLIENT_SECRET"],
    }, timeout=10)
    if not resp.ok:
        print(f"[Todoist OAuth error] status={resp.status_code} body={resp.text}")
        resp.raise_for_status()
    token_data = resp.json()
    # Todoist token response doesn't include expires_in; treat as long-lived
    token_data.setdefault("expires_in", 365 * 24 * 3600)
    set_token(sid, "todoist", token_data)
    return sid


# -----------------------------------------------------------------
# Token refresh
# -----------------------------------------------------------------

def _refresh_token(provider: str, refresh_token: Optional[str]) -> Optional[dict]:
    if not refresh_token:
        return None

    if provider == "google":
        resp = requests.post(GOOGLE_TOKEN_URL, data={
            "client_id":     os.environ["GOOGLE_CLIENT_ID"],
            "client_secret": os.environ["GOOGLE_CLIENT_SECRET"],
            "refresh_token": refresh_token,
            "grant_type":    "refresh_token",
        }, timeout=10)
        if resp.ok:
            data = resp.json()
            data.setdefault("refresh_token", refresh_token)
            return data

    return None
