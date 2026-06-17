"""GET /api/health — lightweight status probe for monitoring/alerts.

Makes a tiny Gemini call to confirm the primary provider's key works and isn't
rate-limited, and reports a machine-readable reason. Always returns HTTP 200 with
a JSON body so a simple GET-based monitor can read it:

  { "ok": true,  "reason": "ok",        "model": "gemini-2.5-flash" }
  { "ok": false, "reason": "rate_limit" } # free-tier quota exhausted (transient)
  { "ok": false, "reason": "auth" }       # key missing/invalid -> needs attention
  { "ok": false, "reason": "no_key" | "error" }

Optional abuse guard: if HEALTH_TOKEN is set, the request must pass ?token=<it>.
The probe is intentionally tiny (Gemini's free tier, a few tokens) so it costs
effectively nothing.
"""
import json
import os
from http.server import BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

from google import genai
from google.genai import errors as gerrors
from google.genai import types as gtypes

GEMINI_MODEL = os.environ.get("GEMINI_MODEL") or "gemini-2.5-flash"


class handler(BaseHTTPRequestHandler):
    def _json(self, status, obj):
        body = json.dumps(obj).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        try:
            token = os.environ.get("HEALTH_TOKEN")
            if token:
                q = parse_qs(urlparse(self.path).query)
                if (q.get("token") or [""])[0] != token:
                    self._json(401, {"ok": False, "reason": "unauthorized"})
                    return

            key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
            if not key:
                self._json(200, {"ok": False, "reason": "no_key"})
                return

            try:
                client = genai.Client(api_key=key)
                client.models.generate_content(
                    model=GEMINI_MODEL,
                    contents="ping",
                    config=gtypes.GenerateContentConfig(max_output_tokens=5),
                )
                self._json(200, {"ok": True, "reason": "ok", "model": GEMINI_MODEL})
            except gerrors.APIError as exc:
                code = getattr(exc, "code", None) or getattr(exc, "status_code", None)
                msg = str(getattr(exc, "message", "") or exc).lower()
                if code == 429 or "resource_exhausted" in msg or "quota" in msg or "rate" in msg:
                    reason = "rate_limit"
                elif code in (401, 403) or "api key" in msg or "permission" in msg or "unauthenticated" in msg:
                    reason = "auth"
                else:
                    reason = "error"
                self._json(200, {"ok": False, "reason": reason})
        except Exception:  # noqa: BLE001 — never 500 a health probe
            self._json(200, {"ok": False, "reason": "error"})

    def log_message(self, *args):
        return
