"""POST /api/queue — hand a spreadsheet job to the self-hosted "last resort" worker
(e.g. a Raspberry Pi running a local LLM) and return immediately.

This exists to get around the platform's ~60s request limit: the worker can take
minutes, so we DON'T wait for it. We just forward the job (the worker acknowledges
in a split second and processes in the background), then the worker emails the
finished .xlsx to the address the user gave. Configured via two env vars:

  PI_WORKER_URL     base URL of the worker (its public tunnel), e.g.
                    https://sheetgenie-pi.example.com   (we POST /generate-async)
  PI_WORKER_SECRET  shared secret sent as X-Worker-Secret so only we can enqueue

If PI_WORKER_URL is unset the endpoint returns 503 and the frontend never offers it.
"""
import json
import os
import re
import urllib.error
import urllib.request
from http.server import BaseHTTPRequestHandler

MAX_BODY_BYTES = 1_000_000          # the queue payload is text only (no attachments)
FORWARD_TIMEOUT = 15                # the worker must ACK fast; it processes async

PI_WORKER_URL = os.environ.get("PI_WORKER_URL")
PI_WORKER_SECRET = os.environ.get("PI_WORKER_SECRET") or ""

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

_CORS_HEADERS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
}


class handler(BaseHTTPRequestHandler):
    """Vercel serverless handler. Class name MUST be `handler`."""

    def _send_json(self, status, obj):
        body = json.dumps(obj).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        for key, value in _CORS_HEADERS.items():
            self.send_header(key, value)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        for key, value in _CORS_HEADERS.items():
            self.send_header(key, value)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_POST(self):
        try:
            # Reject before reading if the body is too large.
            try:
                length = int(self.headers.get("Content-Length") or 0)
            except (TypeError, ValueError):
                length = 0
            if length > MAX_BODY_BYTES:
                self._send_json(413, {"error": "That request is too large."})
                return

            raw = self.rfile.read(length) if length else b""
            try:
                payload = json.loads(raw.decode("utf-8")) if raw else {}
            except (ValueError, UnicodeDecodeError):
                self._send_json(400, {"error": "Malformed request."})
                return
            if not isinstance(payload, dict):
                self._send_json(400, {"error": "Malformed request."})
                return

            prompt = payload.get("prompt")
            if not isinstance(prompt, str) or not prompt.strip():
                self._send_json(400, {"error": "Please describe the spreadsheet you want."})
                return

            email = payload.get("email")
            if not isinstance(email, str) or not _EMAIL_RE.match(email.strip()):
                self._send_json(400, {"error": "Please enter a valid email address so we can send the file."})
                return

            if not PI_WORKER_URL:
                self._send_json(503, {"error": "The backup server isn't set up yet."})
                return

            # Forward only what the worker needs (text — no attachments on this path).
            job = {
                "prompt": prompt.strip(),
                "data": payload.get("data") if isinstance(payload.get("data"), str) else None,
                "clarifications": payload.get("clarifications")
                    if isinstance(payload.get("clarifications"), list) else None,
                "chosenLayout": payload.get("chosenLayout")
                    if isinstance(payload.get("chosenLayout"), dict) else None,
                "locale": payload.get("locale") if isinstance(payload.get("locale"), str) else None,
                "email": email.strip(),
            }

            url = PI_WORKER_URL.rstrip("/") + "/generate-async"
            req = urllib.request.Request(
                url,
                data=json.dumps(job).encode("utf-8"),
                headers={"Content-Type": "application/json",
                         "X-Worker-Secret": PI_WORKER_SECRET},
                method="POST",
            )
            try:
                with urllib.request.urlopen(req, timeout=FORWARD_TIMEOUT) as resp:
                    ok = 200 <= resp.status < 300
            except urllib.error.HTTPError as exc:
                ok = 200 <= exc.code < 300  # some workers ACK with 202 via HTTPError paths
            except (urllib.error.URLError, TimeoutError, OSError):
                self._send_json(
                    504,
                    {"error": "Couldn't reach the backup server right now. Please try again later."},
                )
                return

            if not ok:
                self._send_json(
                    502,
                    {"error": "The backup server couldn't accept the job. Please try again later."},
                )
                return

            self._send_json(200, {"queued": True, "email": email.strip()})
        except Exception:  # noqa: BLE001 — never leak internals to the client
            self._send_json(500, {"error": "Something went wrong queueing your request."})

    def log_message(self, *args):
        return
