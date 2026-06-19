#!/usr/bin/env python3
"""SheetGenie — Raspberry Pi "last resort" worker.

Receives a queued spreadsheet job from the cloud app (POST /generate-async),
acknowledges INSTANTLY (HTTP 202), then in a background thread:
  1. asks a LOCAL LLM (Ollama) to build the SpreadsheetSpec JSON,
  2. renders a real .xlsx with openpyxl (reusing the app's api/generate.py),
  3. emails the finished file to the address the user gave.

Because this runs on your own always-on hardware it has NO request-time limit —
that's the point: it's the fallback for when every cloud provider is busy. The
cloud only ever waits for the instant 202, never for the slow generation.

Expose it behind a public tunnel (e.g. Cloudflare Tunnel) and point the cloud
app's PI_WORKER_URL / PI_WORKER_SECRET at it. Full setup: pi/README.md.
"""
import hmac
import json
import os
import re
import smtplib
import sys
import threading
import traceback
import urllib.request
from email.message import EmailMessage
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

# --- config (environment) --------------------------------------------------
PORT = int(os.environ.get("PORT") or 8080)
WORKER_SECRET = os.environ.get("WORKER_SECRET") or ""
OLLAMA_URL = (os.environ.get("OLLAMA_URL") or "http://localhost:11434").rstrip("/")
OLLAMA_MODEL = os.environ.get("OLLAMA_MODEL") or "qwen2.5:3b"
MAX_TOKENS = int(os.environ.get("MAX_TOKENS") or 4096)
GEN_TIMEOUT = int(os.environ.get("OLLAMA_TIMEOUT") or 900)  # seconds; the Pi is slow

SMTP_HOST = os.environ.get("SMTP_HOST") or "smtp.gmail.com"
SMTP_PORT = int(os.environ.get("SMTP_PORT") or 587)
SMTP_USER = os.environ.get("SMTP_USER") or ""
SMTP_PASS = os.environ.get("SMTP_PASS") or ""
EMAIL_FROM = os.environ.get("EMAIL_FROM") or SMTP_USER
EMAIL_FROM_NAME = os.environ.get("EMAIL_FROM_NAME") or "SheetGenie"

# The app's deterministic renderer is pure openpyxl (no cloud SDKs) — reuse it so
# the Pi produces byte-identical .xlsx files to the cloud path.
_API_DIR = os.environ.get("SHEETGENIE_API_DIR") or os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "api")
sys.path.insert(0, os.path.abspath(_API_DIR))
import generate as _gen  # noqa: E402

# Trimmed, small-model-friendly prompt (Ollama is told to emit JSON via format=json).
SYSTEM_PROMPT = """You turn a request into ONE JSON object describing an Excel workbook, and output nothing else.

Shape:
{"status":"ready","improvedPrompt":"a clear restatement","notes":"one friendly sentence","spec":SPEC}
SPEC  = {"title":str,"sheets":[SHEET,...]}
SHEET = {"name":str,"columns":[COLUMN,...],"rows":[[cell,...],...],"totalsRow":bool(optional),"charts":[CHART,...](optional)}
COLUMN= {"header":str,"type":"text|number|currency|percent|date|formula","format":str(optional),"formula":str(only if type=formula)}
CHART = {"type":"bar|line|pie","title":str,"categoriesColumn":int(1-based),"valueColumns":[int,...]}

Rules:
- ALWAYS set status "ready" and include a spec. Never ask questions.
- rows is an array of arrays aligned to columns by position. A formula column's cell MUST be null (the renderer fills it).
- formula uses Excel syntax with {row} for the current row number, e.g. "=B{row}*C{row}". Prefer SUMIFS/MAXIFS over array formulas like MAX(IF(...)).
- Generate about 8-15 realistic sample rows, unless an empty template is requested (then rows: []).
- Money -> type "currency" with a "format" such as "\\"CHF\\" #,##0.00". Percentages -> type "percent" storing fractions (25% = 0.25). Dates -> type "date" with ISO yyyy-mm-dd strings.
- Never leave a total/summary cell blank — use a real formula. Do NOT set column widths.
- Output ONLY the single JSON object."""


# --- JSON envelope helpers (compact, standalone) ---------------------------
def _extract_json(text):
    if not isinstance(text, str):
        return None
    try:
        obj = json.loads(text.strip())
        if isinstance(obj, dict):
            return obj
    except ValueError:
        pass
    i, j = text.find("{"), text.rfind("}")
    if i != -1 and j > i:
        try:
            obj = json.loads(text[i:j + 1])
            if isinstance(obj, dict):
                return obj
        except ValueError:
            pass
    return None


def _sanitize_spec(spec):
    if not isinstance(spec, dict):
        return None
    sheets = spec.get("sheets")
    if not isinstance(sheets, list) or not sheets:
        return None
    clean = []
    for sheet in sheets:
        if not isinstance(sheet, dict):
            continue
        cols = sheet.get("columns")
        if not isinstance(cols, list) or not cols:
            continue
        rows = sheet.get("rows")
        if isinstance(rows, list):
            sheet["rows"] = [r for r in rows if isinstance(r, list)]
        elif rows is not None:
            sheet["rows"] = []
        clean.append(sheet)
    if not clean:
        return None
    spec["sheets"] = clean
    return spec


def _normalize_result(result):
    if not isinstance(result, dict):
        return None
    spec = _sanitize_spec(result.get("spec"))
    if spec is None:
        return None
    result["spec"] = spec
    return result


# --- LLM call --------------------------------------------------------------
def _ollama(system, prompt):
    body = json.dumps({
        "model": OLLAMA_MODEL,
        "system": system,
        "prompt": prompt,
        "stream": False,
        "format": "json",
        "options": {"temperature": 0.3, "num_predict": MAX_TOKENS, "num_ctx": 8192},
    }).encode("utf-8")
    req = urllib.request.Request(OLLAMA_URL + "/api/generate", data=body,
                                 headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=GEN_TIMEOUT) as resp:
        out = json.loads(resp.read().decode("utf-8"))
    return out.get("response") or ""


def _build_user_text(job):
    parts = ["Request:\n" + (job.get("prompt") or "").strip()]
    loc = job.get("locale")
    if isinstance(loc, str) and loc.strip():
        parts.append("User locale: " + loc.strip()[:32] + " (use a local currency/number format when relevant).")
    layout = job.get("chosenLayout")
    if isinstance(layout, dict) and isinstance(layout.get("sheets"), list):
        parts.append("Use EXACTLY this structure (sheet names + column headers):\n" + json.dumps(layout))
    data = job.get("data")
    if isinstance(data, str) and data.strip():
        parts.append("Data to fill in:\n" + data.strip())
    clar = job.get("clarifications")
    if isinstance(clar, list) and clar:
        qa = []
        for it in clar:
            if isinstance(it, dict):
                qa.append("Q: %s\nA: %s" % (str(it.get("question", "")).strip(),
                                            str(it.get("answer", "")).strip() or "(use a sensible default)"))
        if qa:
            parts.append("Answers to earlier questions:\n" + "\n".join(qa))
    parts.append("Build the full workbook now: status \"ready\" with a complete spec.")
    return "\n\n".join(parts)


def _slug(s):
    s = re.sub(r"[^A-Za-z0-9]+", "-", (s or "spreadsheet")).strip("-").lower()
    return (s or "spreadsheet")[:60]


# --- email -----------------------------------------------------------------
def _send_email(to_addr, subject, body, attachment=None, filename=None):
    msg = EmailMessage()
    msg["From"] = "%s <%s>" % (EMAIL_FROM_NAME, EMAIL_FROM)
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.set_content(body)
    if attachment is not None:
        msg.add_attachment(
            attachment, maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=filename or "spreadsheet.xlsx")
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=60) as s:
        s.starttls()
        if SMTP_USER:
            s.login(SMTP_USER, SMTP_PASS)
        s.send_message(msg)


# --- the slow job (runs in a background thread) ----------------------------
def _process(job):
    to_addr = (job.get("email") or "").strip()
    try:
        text = _ollama(SYSTEM_PROMPT, _build_user_text(job))
        result = _normalize_result(_extract_json(text))
        if result is None:
            # one stricter retry
            text = _ollama(SYSTEM_PROMPT, _build_user_text(job)
                           + "\n\nReturn ONE valid JSON object only; every row must be an array.")
            result = _normalize_result(_extract_json(text))
        if result is None:
            raise ValueError("model did not return a usable spec")

        spec = result["spec"]
        _gen._validate_spec(spec)
        xlsx = _gen._build_workbook(spec)
        title = spec.get("title") or "Spreadsheet"
        _send_email(
            to_addr,
            "Your SheetGenie spreadsheet: " + title,
            "Hi,\n\nHere's the spreadsheet you asked for: \"%s\".\n"
            "It was built on our backup server, so thanks for your patience.\n\n"
            "— SheetGenie" % title,
            attachment=xlsx, filename=_slug(title) + ".xlsx")
        print("[worker] emailed '%s' to %s" % (title, to_addr), flush=True)
    except Exception:  # noqa: BLE001
        traceback.print_exc()
        try:
            if to_addr:
                _send_email(
                    to_addr, "SheetGenie couldn't build your spreadsheet",
                    "Sorry — our backup server couldn't build that one. Please try again "
                    "on the website in a little while (the fast AI may be back).\n\n— SheetGenie")
        except Exception:  # noqa: BLE001
            traceback.print_exc()


# --- HTTP ------------------------------------------------------------------
class Handler(BaseHTTPRequestHandler):
    def _json(self, code, obj):
        body = json.dumps(obj).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path.split("?")[0] == "/health":
            self._json(200, {"ok": True, "model": OLLAMA_MODEL})
        else:
            self._json(404, {"error": "not found"})

    def do_POST(self):
        if self.path.split("?")[0] != "/generate-async":
            self._json(404, {"error": "not found"})
            return
        # Shared-secret auth (constant-time compare).
        if WORKER_SECRET and not hmac.compare_digest(
                self.headers.get("X-Worker-Secret", ""), WORKER_SECRET):
            self._json(401, {"error": "unauthorized"})
            return
        try:
            length = int(self.headers.get("Content-Length") or 0)
        except (TypeError, ValueError):
            length = 0
        if length <= 0 or length > 1_000_000:
            self._json(400, {"error": "bad request"})
            return
        try:
            job = json.loads(self.rfile.read(length).decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            self._json(400, {"error": "malformed json"})
            return
        if not isinstance(job, dict) or not str(job.get("prompt", "")).strip() \
                or not str(job.get("email", "")).strip():
            self._json(400, {"error": "prompt and email are required"})
            return
        # Acknowledge immediately; do the slow work in the background.
        threading.Thread(target=_process, args=(job,), daemon=True).start()
        self._json(202, {"accepted": True})

    def log_message(self, *args):
        return


if __name__ == "__main__":
    print("[worker] SheetGenie Pi worker on :%d  model=%s  ollama=%s"
          % (PORT, OLLAMA_MODEL, OLLAMA_URL), flush=True)
    if not WORKER_SECRET:
        print("[worker] WARNING: WORKER_SECRET is empty — set it so only your app can enqueue!", flush=True)
    ThreadingHTTPServer(("0.0.0.0", PORT), Handler).serve_forever()
