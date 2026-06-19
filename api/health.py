"""GET /api/health — lightweight status probe for monitoring/alerts.

Pings a provider to confirm its key works and isn't rate-limited, and reports a
machine-readable reason. Always returns HTTP 200 so a simple GET monitor can read it:

  { "ok": true,  "reason": "ok",   "provider": "gemini", "model": "gemini-2.5-flash" }
  { "ok": false, "reason": "rate_limit" }   # quota exhausted (transient)
  { "ok": false, "reason": "auth" }         # key missing/invalid -> needs attention
  { "ok": false, "reason": "no_key" | "error" }

Which provider:  ?provider=gemini  (default)  or  ?provider=groq.
Optional abuse guard: if HEALTH_TOKEN is set, the request must pass ?token=<it>.
The probe is intentionally tiny (a few tokens) so it costs effectively nothing.
"""
import json
import os
from http.server import BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

GEMINI_MODEL = os.environ.get("GEMINI_MODEL") or "gemini-2.5-flash"
GROQ_MODEL = os.environ.get("GROQ_MODEL") or "llama-3.3-70b-versatile"
GROQ_BASE_URL = os.environ.get("GROQ_BASE_URL") or "https://api.groq.com/openai/v1"


def _probe_gemini():
    """Return (ok, reason, model) for a tiny Gemini generation call."""
    key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not key:
        return False, "no_key", GEMINI_MODEL
    from google import genai
    from google.genai import errors as gerrors
    from google.genai import types as gtypes
    try:
        genai.Client(api_key=key).models.generate_content(
            model=GEMINI_MODEL, contents="ping",
            config=gtypes.GenerateContentConfig(max_output_tokens=5))
        return True, "ok", GEMINI_MODEL
    except gerrors.APIError as exc:
        code = getattr(exc, "code", None) or getattr(exc, "status_code", None)
        msg = str(getattr(exc, "message", "") or exc).lower()
        if code == 429 or "resource_exhausted" in msg or "quota" in msg or "rate" in msg:
            return False, "rate_limit", GEMINI_MODEL
        if code in (401, 403) or "api key" in msg or "permission" in msg or "unauthenticated" in msg:
            return False, "auth", GEMINI_MODEL
        return False, "error", GEMINI_MODEL
    except Exception:  # noqa: BLE001
        return False, "error", GEMINI_MODEL


def _probe_groq():
    """Return (ok, reason, model) for a tiny Groq chat-completion call."""
    key = os.environ.get("GROQ_API_KEY")
    if not key:
        return False, "no_key", GROQ_MODEL
    import openai
    from openai import OpenAI
    try:
        OpenAI(api_key=key, base_url=GROQ_BASE_URL).chat.completions.create(
            model=GROQ_MODEL, max_tokens=5,
            messages=[{"role": "user", "content": "ping"}])
        return True, "ok", GROQ_MODEL
    except openai.RateLimitError:
        return False, "rate_limit", GROQ_MODEL
    except openai.AuthenticationError:
        return False, "auth", GROQ_MODEL
    except Exception:  # noqa: BLE001
        return False, "error", GROQ_MODEL


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
            q = parse_qs(urlparse(self.path).query)
            token = os.environ.get("HEALTH_TOKEN")
            if token and (q.get("token") or [""])[0] != token:
                self._json(401, {"ok": False, "reason": "unauthorized"})
                return

            provider = (q.get("provider") or ["gemini"])[0].lower()
            if provider == "groq":
                ok, reason, model = _probe_groq()
            else:
                provider = "gemini"
                ok, reason, model = _probe_gemini()

            out = {"ok": ok, "reason": reason, "provider": provider}
            if ok:
                out["model"] = model
            self._json(200, out)
        except Exception:  # noqa: BLE001 — never 500 a health probe
            self._json(200, {"ok": False, "reason": "error"})

    def log_message(self, *args):
        return
