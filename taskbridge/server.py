#!/usr/bin/env python3
"""
TaskBridge – main HTTP server.

Routes:
  GET  /                        → main UI
  GET  /static/<file>           → static assets
  GET  /api/status              → auth status for current session
  POST /api/save                → parse input & save to Todoist / Google Calendar
  GET  /auth/google             → start Google OAuth flow
  GET  /auth/google/callback    → Google OAuth callback
  GET  /auth/todoist            → start Todoist OAuth flow
  GET  /auth/todoist/callback   → Todoist OAuth callback
  GET  /auth/logout             → clear session

Run:
  python server.py              (default port 8000)
  PORT=3000 python server.py
"""

import json
import mimetypes
import os
import sys
import traceback
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

# Allow importing from lib/
sys.path.insert(0, str(Path(__file__).parent))

from lib import oauth
from lib.parser import parse_input
from lib.todoist import create_task
from lib.google_cal import create_event

# Load .env if present
_env_path = Path(__file__).parent / ".env"
if _env_path.exists():
    for line in _env_path.read_text().splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, _, v = line.partition("=")
            os.environ.setdefault(k.strip(), v.strip())

TEMPLATES_DIR = Path(__file__).parent / "templates"
STATIC_DIR    = Path(__file__).parent / "static"

SESSION_COOKIE = "tb_session"


# -----------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------

def _get_session_id(handler: "TaskBridgeHandler") -> str:
    """Return existing session id from cookie, or create a new one."""
    cookie_header = handler.headers.get("Cookie", "")
    for part in cookie_header.split(";"):
        part = part.strip()
        if part.startswith(f"{SESSION_COOKIE}="):
            return part[len(SESSION_COOKIE) + 1:]
    return oauth.new_session()


def _json(handler: "TaskBridgeHandler", status: int, data: dict,
          extra_headers: list[tuple[str, str]] | None = None) -> None:
    body = json.dumps(data).encode()
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    for k, v in (extra_headers or []):
        handler.send_header(k, v)
    handler.end_headers()
    handler.wfile.write(body)


def _redirect(handler: "TaskBridgeHandler", location: str,
              extra_headers: list[tuple[str, str]] | None = None) -> None:
    handler.send_response(302)
    handler.send_header("Location", location)
    for k, v in (extra_headers or []):
        handler.send_header(k, v)
    handler.end_headers()


def _base_url(handler: "TaskBridgeHandler") -> str:
    host = handler.headers.get("Host", "localhost:8000")
    scheme = "https" if handler.headers.get("X-Forwarded-Proto") == "https" else "http"
    return f"{scheme}://{host}"


def _set_session_cookie(sid: str) -> tuple[str, str]:
    return ("Set-Cookie", f"{SESSION_COOKIE}={sid}; Path=/; HttpOnly; SameSite=Lax")


# -----------------------------------------------------------------
# Request handler
# -----------------------------------------------------------------

class TaskBridgeHandler(BaseHTTPRequestHandler):

    def log_message(self, fmt, *args):  # suppress default access log noise
        pass

    # ---- GET ----

    def do_GET(self):
        parsed = urlparse(self.path)
        path   = parsed.path
        qs     = parse_qs(parsed.query)

        try:
            if path == "/" or path == "":
                self._serve_file(TEMPLATES_DIR / "index.html", "text/html")

            elif path.startswith("/static/"):
                rel = path[len("/static/"):]
                self._serve_file(STATIC_DIR / rel)

            elif path == "/api/status":
                self._handle_status()

            elif path == "/auth/google":
                self._handle_google_auth()

            elif path == "/auth/google/callback":
                self._handle_google_callback(qs)

            elif path == "/auth/todoist":
                self._handle_todoist_auth()

            elif path == "/auth/todoist/callback":
                self._handle_todoist_callback(qs)

            elif path == "/auth/logout":
                self._handle_logout()

            else:
                self._not_found()

        except Exception:
            traceback.print_exc()
            self._error(500, "Internal server error")

    # ---- POST ----

    def do_POST(self):
        parsed = urlparse(self.path)
        path   = parsed.path

        try:
            if path == "/api/save":
                self._handle_save()
            else:
                self._not_found()
        except Exception:
            traceback.print_exc()
            self._error(500, "Internal server error")

    # ---- Route handlers ----

    def _handle_status(self):
        sid    = _get_session_id(self)
        status = oauth.is_authenticated(sid)
        _json(self, 200, status, [_set_session_cookie(sid)])

    def _handle_google_auth(self):
        sid      = _get_session_id(self)
        auth_url = oauth.google_auth_url(_base_url(self), sid)
        _redirect(self, auth_url, [_set_session_cookie(sid)])

    def _handle_google_callback(self, qs: dict):
        code  = qs.get("code",  [""])[0]
        state = qs.get("state", [""])[0]
        if not code or not state:
            self._error(400, "Missing code or state")
            return
        sid = oauth.google_exchange_code(code, state, _base_url(self))
        if not sid:
            self._error(400, "Invalid OAuth state")
            return
        _redirect(self, "/?connected=google", [_set_session_cookie(sid)])

    def _handle_todoist_auth(self):
        sid      = _get_session_id(self)
        auth_url = oauth.todoist_auth_url(_base_url(self), sid)
        _redirect(self, auth_url, [_set_session_cookie(sid)])

    def _handle_todoist_callback(self, qs: dict):
        code  = qs.get("code",  [""])[0]
        state = qs.get("state", [""])[0]
        error = qs.get("error", [""])[0]
        if error:
            _redirect(self, "/?error=todoist_denied")
            return
        if not code or not state:
            self._error(400, "Missing code or state")
            return
        sid = oauth.todoist_exchange_code(code, state)
        if not sid:
            self._error(400, "Invalid OAuth state")
            return
        _redirect(self, "/?connected=todoist", [_set_session_cookie(sid)])

    def _handle_logout(self):
        sid = _get_session_id(self)
        # Invalidate session
        oauth._sessions.pop(sid, None)
        new_sid = oauth.new_session()
        _redirect(self, "/", [_set_session_cookie(new_sid)])

    def _handle_save(self):
        sid = _get_session_id(self)

        length = int(self.headers.get("Content-Length", 0))
        body   = self.rfile.read(length)
        try:
            data = json.loads(body)
        except json.JSONDecodeError:
            _json(self, 400, {"error": "Invalid JSON"}, [_set_session_cookie(sid)])
            return

        text = (data.get("text") or "").strip()
        if not text:
            _json(self, 400, {"error": "텍스트를 입력해주세요."}, [_set_session_cookie(sid)])
            return

        parsed   = parse_input(text)
        target   = parsed.target
        content  = parsed.text

        if not content:
            _json(self, 400, {"error": "내용이 없습니다."}, [_set_session_cookie(sid)])
            return

        auth = oauth.is_authenticated(sid)
        results: dict = {}
        errors:  list = []

        # --- Save to Todoist ---
        if target in ("both", "todoist"):
            if not auth.get("todoist"):
                errors.append("Todoist 연결이 필요합니다.")
            else:
                t_token = oauth.get_access_token(sid, "todoist")
                try:
                    task = create_task(t_token, content)
                    results["todoist"] = {"id": task.get("id"), "url": task.get("url")}
                except Exception as e:
                    errors.append(f"Todoist 저장 실패: {e}")

        # --- Save to Google Calendar ---
        if target in ("both", "calendar"):
            if not auth.get("google"):
                errors.append("Google Calendar 연결이 필요합니다.")
            else:
                g_token = oauth.get_access_token(sid, "google")
                try:
                    event = create_event(g_token, content)
                    results["calendar"] = {
                        "id":      event.get("id"),
                        "htmlLink": event.get("htmlLink"),
                    }
                except Exception as e:
                    errors.append(f"Google Calendar 저장 실패: {e}")

        if errors and not results:
            _json(self, 400, {"error": " / ".join(errors)}, [_set_session_cookie(sid)])
        else:
            _json(self, 200, {
                "ok":      True,
                "target":  target,
                "results": results,
                "errors":  errors,
            }, [_set_session_cookie(sid)])

    # ---- Static file serving ----

    def _serve_file(self, path: Path, content_type: str | None = None):
        if not path.exists() or not path.is_file():
            self._not_found()
            return
        data = path.read_bytes()
        ct   = content_type or mimetypes.guess_type(str(path))[0] or "application/octet-stream"
        self.send_response(200)
        self.send_header("Content-Type", ct)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _not_found(self):
        self._error(404, "Not found")

    def _error(self, status: int, message: str):
        body = message.encode()
        self.send_response(status)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


# -----------------------------------------------------------------
# Entry point
# -----------------------------------------------------------------

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    server = HTTPServer(("0.0.0.0", port), TaskBridgeHandler)
    print(f"TaskBridge running → http://localhost:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
