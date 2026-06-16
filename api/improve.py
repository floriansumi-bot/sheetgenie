"""POST /api/improve — turn a rough prompt into {improvedPrompt, notes, spec}.

Vercel Python serverless function (native http.server pattern). This is the only
endpoint that calls the Anthropic API. It returns a validated-shape JSON object
whose `spec` is a SpreadsheetSpec per docs/SPEC.md; the deterministic
/api/generate endpoint renders that spec into a real .xlsx.

Contract: docs/SPEC.md §1 (HTTP API) and §2 (SpreadsheetSpec) are the single
source of truth. Keep this file aligned with them.
"""

import json
import os
import sys
import traceback
from http.server import BaseHTTPRequestHandler

import anthropic


# ---------------------------------------------------------------------------
# SpreadsheetSpec as a JSON Schema for Anthropic structured outputs.
#
# Encodes the FULL SpreadsheetSpec from docs/SPEC.md §2. Structured outputs do
# NOT support minLength/maxLength/minimum/maximum/multipleOf or recursive $refs,
# so the hard numeric limits (<=8 sheets, <=50 cols, <=5000 rows, <=6 charts,
# percent-as-fraction, etc.) live in the system prompt instead and are enforced
# by /api/generate. Every object sets additionalProperties:false. Only genuinely
# required fields are marked required.
# ---------------------------------------------------------------------------

_COLUMN_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "header": {"type": "string"},
        "type": {
            "type": "string",
            "enum": ["text", "number", "currency", "percent", "date", "formula"],
        },
        "width": {"type": "number"},
        # format / formula may legitimately be null (only formula columns carry a
        # formula); allow string-or-null so the model can emit explicit nulls.
        "format": {"type": ["string", "null"]},
        "formula": {"type": ["string", "null"]},
    },
    "required": ["header", "type"],
}

_CHART_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "type": {"type": "string", "enum": ["bar", "line", "pie"]},
        "title": {"type": "string"},
        "categoriesColumn": {"type": "integer"},
        "valueColumns": {"type": "array", "items": {"type": "integer"}},
        "dataStartRow": {"type": ["integer", "null"]},
        "dataEndRow": {"type": ["integer", "null"]},
        "anchor": {"type": ["string", "null"]},
    },
    "required": ["type", "title", "categoriesColumn", "valueColumns"],
}

_SHEET_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "name": {"type": "string"},
        "description": {"type": ["string", "null"]},
        "columns": {"type": "array", "items": _COLUMN_SCHEMA},
        # Row = [cell, ...]; cell = string | number | boolean | null.
        "rows": {
            "type": "array",
            "items": {
                "type": "array",
                "items": {"type": ["string", "number", "boolean", "null"]},
            },
        },
        "freezeHeader": {"type": "boolean"},
        "autoFilter": {"type": "boolean"},
        "charts": {"type": "array", "items": _CHART_SCHEMA},
    },
    "required": ["name", "columns"],
}

_SPREADSHEET_SPEC_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "title": {"type": "string"},
        "sheets": {"type": "array", "items": _SHEET_SCHEMA},
    },
    "required": ["title", "sheets"],
}

_QUESTION_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "question": {"type": "string"},
        "hint": {"type": ["string", "null"]},
    },
    "required": ["question"],
}

# Two-mode response (see docs/SPEC.md §1): status "ready" carries improvedPrompt +
# spec; status "needs_input" carries questions. Only status + notes are always
# required; spec/improvedPrompt/questions are conditional, so they are optional
# here and the system prompt governs which appear. spec is nullable so the model
# can omit it cleanly when asking questions.
RESPONSE_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "status": {"type": "string", "enum": ["ready", "needs_input"]},
        "notes": {"type": "string"},
        "improvedPrompt": {"type": ["string", "null"]},
        "spec": {"anyOf": [_SPREADSHEET_SPEC_SCHEMA, {"type": "null"}]},
        "questions": {"type": "array", "items": _QUESTION_SCHEMA},
    },
    "required": ["status", "notes"],
}


SYSTEM_PROMPT = """\
You are an expert spreadsheet architect. The user gives you a rough request \
(and sometimes pasted tabular data). You design a complete, well-structured \
workbook blueprint that another component renders deterministically into a real \
.xlsx file. Return structured JSON.

STEP 1 — DECIDE IF YOU CAN BUILD THE RIGHT THING.
Judge whether the request (plus any pasted data and any answers already provided) \
is clear enough to build a CORRECT, useful spreadsheet. Ask for clarification \
ONLY when a wrong assumption would produce the WRONG spreadsheet — e.g. the \
goal/scope is too vague to know the columns; essential fields, time period, or \
grouping are missing; pasted data has ambiguous or unlabeled columns you cannot \
confidently map; or it is unclear whether they want an empty template or \
filled-in data. In that case set:
  status = "needs_input"; notes = one friendly line saying you need a couple of \
details; questions = 1 to 4 SHORT, plain-language questions, each answerable in a \
few words, each with a helpful "hint" example. Do NOT include a spec.
PREFER sensible defaults over asking — most clear requests need NO questions, so \
just build. Never ask more than 4 questions, and never ask about cosmetic \
formatting trivia.
IMPORTANT: if the user has ALREADY answered questions (they appear below under \
"Answers to your questions"), do NOT ask again — make reasonable assumptions and \
BUILD with status = "ready".

STEP 2 — WHEN YOU CAN BUILD, set status = "ready" and return exactly these three:

1. improvedPrompt — a clear, specific restatement of the workbook to build. \
Name the sheets, the columns, the kinds of data, and any calculations or charts. \
Write it as the polished version of what the user meant.

2. notes — one or two friendly, plain-language sentences for a non-technical \
user describing what you made. No jargon.

3. spec — a SpreadsheetSpec object following these rules EXACTLY:

WORKBOOK SHAPE
- spec.title: a short workbook title.
- spec.sheets: an array of sheets, left-to-right. Use multiple sheets only when \
the request clearly involves distinct datasets.
- Each sheet has: name, optional description, columns (left-to-right), optional \
rows, optional freezeHeader (default true), optional autoFilter (default true), \
optional charts.

COLUMNS
- Each column has a header and a type: one of text, number, currency, percent, \
date, formula.
- Choose the most sensible type per column. Money -> currency. Ratios/rates -> \
percent. Calendar dates -> date (use ISO yyyy-mm-dd strings in rows). Counts and \
quantities -> number. Labels/names/notes -> text.
- width and format are optional; omit them unless a specific width or explicit \
Excel number format is genuinely needed.

FORMULA COLUMNS
- For a calculated column, set type to "formula" and provide a "formula" that is \
an Excel formula string using the literal token {row} for the current 1-based \
row number, e.g. "=B{row}-C{row}" or "=B{row}*C{row}". Column letters refer to \
the FINAL left-to-right column order.
- In every row, the cell for a formula column MUST be null — the renderer fills \
it from the formula. Never put a computed value there.

PERCENT VALUES
- percent values are fractions, not whole numbers: 25% is 0.25, not 25. \
A 7.5% rate is 0.075.

ROWS
- rows is an array of arrays. Each inner array is one row, aligned to columns by \
index (same order as columns).
- If the user supplied pasted data, parse it carefully into rows aligned to your \
columns. Map their values to the right columns; convert percentages to fractions; \
keep dates as ISO strings.
- If no data was supplied, generate a small set of realistic, plausible sample \
rows (about 8 to 15) so the workbook is immediately useful — UNLESS the user \
clearly wants a blank/empty template, in which case set rows to [].
- Formula-column cells are always null in rows (see above).

CHARTS
- Add charts when the request implies visualization (e.g. "with a chart", \
"compare", "trend", "breakdown") or when a chart would clearly help.
- Each chart: type (bar, line, pie), title, categoriesColumn (1-based column \
index for category labels / pie slices), valueColumns (array of 1-based column \
indices for the numeric series). Optionally dataStartRow, dataEndRow, anchor — \
omit these (or use null) to let the renderer auto-place.
- All chart column indices must be valid 1-based indices into that sheet's \
columns, and value columns must point at numeric/currency/percent/formula \
columns.

HARD LIMITS (never exceed)
- At most 8 sheets. At most 50 columns per sheet. At most 5000 rows per sheet. \
At most 6 charts per sheet. At most 10 value columns per chart.

Keep the spec minimal and explicit: the renderer does exactly what the spec \
says and nothing more. Prefer clarity and correctness over cleverness."""


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Model selection. Default to the most capable model (Fable 5). If the API key
# cannot access a model (404 not_found / 403 permission), fall back to the next
# one at runtime — so a key without Fable access still works on Opus 4.8, then
# Sonnet 4.6. Override the primary via the MODEL env var (a comma-separated list
# is also accepted). See docs/DEPLOY.md / .env.example.
def _model_chain():
    primary = (os.environ.get("MODEL") or "claude-fable-5").strip()
    chain = [m.strip() for m in primary.split(",") if m.strip()]
    for fallback in ("claude-opus-4-8", "claude-sonnet-4-6"):
        if fallback not in chain:
            chain.append(fallback)
    return chain or ["claude-fable-5", "claude-opus-4-8", "claude-sonnet-4-6"]


MODEL_CHAIN = _model_chain()

# Models that accept adaptive thinking + the effort parameter. Haiku / older
# Sonnet 4.5 reject these (HTTP 400), so we only send them to supported models.
_THINKING_MODELS = (
    "claude-fable-5",
    "claude-opus-4-8",
    "claude-opus-4-7",
    "claude-opus-4-6",
    "claude-sonnet-4-6",
)


def _supports_thinking(model):
    return any(model == m or model.startswith(m + "[") for m in _THINKING_MODELS)


# Thinking effort: low | medium | high | max. Higher = more reasoning, slower.
EFFORT = os.environ.get("EFFORT") or "high"

# Output token ceiling. Kept moderate so generation comfortably finishes inside
# Vercel's 60s function limit; raise via MAX_TOKENS for very large data fills.
try:
    MAX_TOKENS = int(os.environ.get("MAX_TOKENS") or "10000")
except (TypeError, ValueError):
    MAX_TOKENS = 10000

# Hard cap on the request body (prompt + pasted data) to prevent abuse.
MAX_BODY_BYTES = 256 * 1024  # 256 KB

_CORS_HEADERS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
}


class handler(BaseHTTPRequestHandler):
    """Vercel serverless handler. Class name MUST be `handler` (SPEC.md §1)."""

    def _send_json(self, status, obj):
        """Write a JSON response with the given status code and CORS headers."""
        body = json.dumps(obj).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        for key, value in _CORS_HEADERS.items():
            self.send_header(key, value)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        """CORS preflight: 204 No Content with the permissive CORS headers."""
        self.send_response(204)
        for key, value in _CORS_HEADERS.items():
            self.send_header(key, value)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_POST(self):
        try:
            # --- Parse the request body -------------------------------------
            try:
                length = int(self.headers.get("Content-Length") or 0)
            except (TypeError, ValueError):
                length = 0

            if length > MAX_BODY_BYTES:
                self._send_json(
                    413,
                    {"error": "Your request is too large. Please shorten the "
                              "prompt or the data you pasted."},
                )
                return

            raw = self.rfile.read(min(length, MAX_BODY_BYTES)) if length > 0 else b""

            try:
                payload = json.loads(raw.decode("utf-8")) if raw else None
            except (ValueError, UnicodeDecodeError):
                payload = None

            if not isinstance(payload, dict):
                self._send_json(
                    400,
                    {"error": "Please enter a prompt describing the spreadsheet you want."},
                )
                return

            prompt = payload.get("prompt")
            if not isinstance(prompt, str) or not prompt.strip():
                self._send_json(
                    400,
                    {"error": "Please enter a prompt describing the spreadsheet you want."},
                )
                return

            # --- Check configuration ----------------------------------------
            api_key = os.environ.get("ANTHROPIC_API_KEY")
            if not api_key:
                self._send_json(
                    500,
                    {"error": "The server is not configured yet (missing API key)."},
                )
                return

            # --- Build the user message -------------------------------------
            data = payload.get("data")
            user_text = "Request:\n" + prompt.strip()
            if isinstance(data, str) and data.strip():
                user_text += (
                    "\n\nUser-provided data to fill in:\n" + data.strip()
                )

            # Fold in answers to any clarifying questions we asked previously, so
            # the model now has enough to build (and must not ask again).
            clar = payload.get("clarifications")
            if isinstance(clar, list) and clar:
                qa_lines = []
                for item in clar:
                    if not isinstance(item, dict):
                        continue
                    q = str(item.get("question", "")).strip()
                    a = str(item.get("answer", "")).strip()
                    if q or a:
                        qa_lines.append("Q: " + q + "\nA: " + (a or "(no answer given)"))
                if qa_lines:
                    user_text += (
                        "\n\nAnswers to your questions (do NOT ask again — build the "
                        "spreadsheet now):\n" + "\n".join(qa_lines)
                    )

            # --- Call the Anthropic API -------------------------------------
            # The SDK resolves ANTHROPIC_API_KEY from the environment. We never
            # pass temperature/top_p/top_k (removed on current models -> HTTP 400).
            # Adaptive thinking + effort are sent only to models that support them
            # (see _supports_thinking). We walk MODEL_CHAIN and fall back when a
            # model is unavailable to this key (404 not_found / 403 permission).
            client = anthropic.Anthropic()
            resp = None
            unavailable = None
            for model in MODEL_CHAIN:
                output_config = {
                    "format": {"type": "json_schema", "schema": RESPONSE_SCHEMA}
                }
                extra = {}
                if _supports_thinking(model):
                    extra["thinking"] = {"type": "adaptive"}
                    output_config["effort"] = EFFORT
                try:
                    resp = client.messages.create(
                        model=model,
                        max_tokens=MAX_TOKENS,
                        system=SYSTEM_PROMPT,
                        messages=[{"role": "user", "content": user_text}],
                        output_config=output_config,
                        **extra,
                    )
                    break
                except (anthropic.NotFoundError, anthropic.PermissionDeniedError) as exc:
                    # This key can't use that model — try the next in the chain.
                    unavailable = exc
                    continue

            if resp is None:
                # No model in the chain was accessible with this key.
                raise unavailable if unavailable else RuntimeError("no model available")

            # --- Parse the structured response ------------------------------
            if resp.stop_reason == "max_tokens":
                self._send_json(
                    500,
                    {"error": "That request produced a spreadsheet plan too large "
                              "to finish. Try a simpler request or less data."},
                )
                return

            # output_config.format guarantees a text block with valid JSON
            # matching RESPONSE_SCHEMA, but we still parse defensively.
            text = next((b.text for b in resp.content if b.type == "text"), None)
            if not text:
                self._send_json(
                    500,
                    {"error": "The model did not return a spreadsheet plan. "
                              "Please try again."},
                )
                return

            try:
                result = json.loads(text)
            except json.JSONDecodeError:
                self._send_json(
                    500,
                    {"error": "The spreadsheet plan came back malformed. "
                              "Please try again."},
                )
                return

            self._send_json(200, result)

        except Exception:  # noqa: BLE001 — sanitize all failures for the client
            # Log the full error server-side (Vercel captures stderr) but never
            # leak the API key, raw exception text, or a stack trace to the client.
            traceback.print_exc(file=sys.stderr)
            try:
                self._send_json(
                    500,
                    {
                        "error": "Something went wrong generating your spreadsheet plan. "
                        "Please try again."
                    },
                )
            except Exception:  # noqa: BLE001 — headers/body may already be sent
                traceback.print_exc(file=sys.stderr)

    def log_message(self, *args):  # noqa: D401 — silence default stderr access logs
        return
