"""GET /api/health — lightweight status probe for monitoring/alerts.

Makes a tiny (max_tokens=1) Anthropic call to confirm the API key still works and
has credit, and reports a machine-readable reason. Always returns HTTP 200 with a
JSON body so a simple GET-based monitor can read it:

  { "ok": true,  "reason": "ok",        "model": "claude-opus-4-8" }
  { "ok": false, "reason": "credit" }     # out of credit -> owner should top up
  { "ok": false, "reason": "auth" }       # key missing/invalid
  { "ok": false, "reason": "rate_limit" } # transient, no action needed
  { "ok": false, "reason": "no_key" | "no_model" | "error" }

Optional abuse guard: if HEALTH_TOKEN is set, the request must pass ?token=<it>.
The probe is intentionally cheap (~a few tokens) so it costs effectively nothing.
"""
import json
import os
from http.server import BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

import anthropic

# Try the user's working models in order; any one succeeding means the key + credit
# are fine (credit is shared across models, so an out-of-credit account fails all).
_PROBE_CHAIN = ["claude-opus-4-8", "claude-sonnet-4-6", "claude-haiku-4-5"]


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

            if not os.environ.get("ANTHROPIC_API_KEY"):
                self._json(200, {"ok": False, "reason": "no_key"})
                return

            client = anthropic.Anthropic()
            last = None
            for model in _PROBE_CHAIN:
                try:
                    client.messages.create(
                        model=model, max_tokens=1,
                        messages=[{"role": "user", "content": "ping"}],
                    )
                    self._json(200, {"ok": True, "reason": "ok", "model": model})
                    return
                except (anthropic.NotFoundError, anthropic.PermissionDeniedError) as exc:
                    last = exc           # model not available to this key — try next
                    continue
                except anthropic.AuthenticationError:
                    self._json(200, {"ok": False, "reason": "auth"})
                    return
                except anthropic.RateLimitError:
                    self._json(200, {"ok": False, "reason": "rate_limit"})
                    return
                except anthropic.BadRequestError as exc:
                    m = str(getattr(exc, "message", "") or exc).lower()
                    reason = "credit" if ("credit" in m or "billing" in m or "quota" in m) else "error"
                    self._json(200, {"ok": False, "reason": reason})
                    return

            m = str(getattr(last, "message", "") or last).lower()
            if "credit" in m or "billing" in m or "quota" in m:
                self._json(200, {"ok": False, "reason": "credit"})
            else:
                self._json(200, {"ok": False, "reason": "no_model"})
        except Exception:  # noqa: BLE001 — never 500 a health probe
            self._json(200, {"ok": False, "reason": "error"})

    def log_message(self, *args):
        return
