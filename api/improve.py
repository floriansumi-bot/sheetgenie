"""POST /api/improve — turn a rough prompt into {improvedPrompt, notes, spec}.

Vercel Python serverless function (native http.server pattern). This is the only
endpoint that calls an AI provider — free Google Gemini first, falling back to xAI
Grok. It returns a validated-shape JSON object whose `spec` is a SpreadsheetSpec per
docs/SPEC.md; the deterministic /api/generate endpoint renders it into a real .xlsx.

Contract: docs/SPEC.md §1 (HTTP API) and §2 (SpreadsheetSpec) are the single
source of truth. Keep this file aligned with them.
"""

import base64
import json
import os
import re
import sys
import traceback
from http.server import BaseHTTPRequestHandler

from google import genai
from google.genai import errors as gerrors
from google.genai import types as gtypes
import openai
from openai import OpenAI


# ---------------------------------------------------------------------------
# SpreadsheetSpec as a JSON Schema (REFERENCE ONLY — not sent to the API).
#
# NOTE: the full schema below compiles to a grammar that exceeds Anthropic's
# structured-output size limit ("compiled grammar is too large"), so it is no
# longer passed as output_config.format. It is kept here as the canonical record
# of the JSON shape; at runtime improve.py pins the envelope in SYSTEM_PROMPT
# (see "OUTPUT FORMAT") and parses the reply defensively with _extract_json.
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

# Conditional-formatting rule on a column's data range (SPEC.md §2 CondFmt).
# value/value2 may be string-or-number-or-null; comparison rules use value (+
# value2 for "between"); top10/bottom10/colorScale omit them. color is a named
# colour or a 6-hex string, or null to let the renderer choose.
_CONDFMT_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "column": {"type": "integer"},
        "rule": {
            "type": "string",
            "enum": [
                "greaterThan",
                "greaterThanOrEqual",
                "lessThan",
                "lessThanOrEqual",
                "equal",
                "between",
                "top10",
                "bottom10",
                "colorScale",
            ],
        },
        "value": {"type": ["string", "number", "null"]},
        "value2": {"type": ["string", "number", "null"]},
        "color": {"type": ["string", "null"]},
    },
    "required": ["column", "rule"],
}

# Dropdown/list validation on a column's data range (SPEC.md §2 Validation).
_VALIDATION_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "column": {"type": "integer"},
        "values": {"type": "array", "items": {"type": "string"}},
    },
    "required": ["column", "values"],
}

# Workbook-level defined name (SPEC.md §2 NamedRange). Formula columns may
# reference `name` instead of a literal range; `ref` is an Excel reference such
# as "'Settings'!$B$1".
_NAMED_RANGE_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "name": {"type": "string"},
        "ref": {"type": "string"},
    },
    "required": ["name", "ref"],
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
        # Advanced features (all optional; see SPEC.md §2 "Advanced features").
        "totalsRow": {"type": "boolean"},
        "conditionalFormats": {"type": "array", "items": _CONDFMT_SCHEMA},
        "dataValidations": {"type": "array", "items": _VALIDATION_SCHEMA},
    },
    "required": ["name", "columns"],
}

_SPREADSHEET_SPEC_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "title": {"type": "string"},
        "sheets": {"type": "array", "items": _SHEET_SCHEMA},
        # Optional workbook-level defined names.
        "namedRanges": {"type": "array", "items": _NAMED_RANGE_SCHEMA},
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
confidently map; it is unclear whether they want an empty template or filled-in \
data; or any CONSEQUENTIAL choice the user did NOT specify is genuinely ambiguous \
— a unit, dimension, currency, region, time period, or granularity that changes \
the numbers or their meaning. GENERAL RULE: when you would otherwise pick a default \
for a consequential choice the user did not state, prefer to ASK and offer the \
options (you MAY put your recommended one first) rather than deciding silently. \
You MUST ask (never silently assume) for, e.g.: temperature with no unit (Celsius \
or Fahrenheit?); a duration / "how long" / "when" with no unit (years, months, or \
weeks?); weight (kg/lb), distance (km/mi), or volume; an ambiguous date (is 03/04 \
the 4th of March or 3rd of April?); and ESPECIALLY any prices / costs / money with \
NO stated currency — do NOT default to USD; ask which currency and recommend the \
user's local one (infer it from "User locale" below — e.g. CHF for Switzerland, \
EUR for the euro area, GBP for the UK). Honor a choice the user DID state (never \
re-ask something they already specified). Also run through this CHECKLIST and ask when any item is unstated and would change \
the data, numbers, or their meaning: TIME (unit/granularity years-months-weeks-days, \
the period or date range covered, fiscal-vs-calendar, time zone); SCOPE (which \
region/country/market, which subset, how many items, and the ranking/selection \
basis -- "top"/"popular"/"best" by what measure?); DATA REALISM (whether to fetch \
LIVE current data via web search or use approximate figures -- and NEVER present \
invented numbers as authoritative facts: verify them, or clearly mark them as \
estimates; matters for prices, statistics, anything "current/latest"); FORMULA \
ASSUMPTIONS not given (rates, tax %, growth/discount rates, compounding, rounding, \
inclusive/exclusive boundaries); and ambiguous DEFINITIONS or thresholds. Don't \
just assume -- if a thoughtful analyst would check it first, ask. \
Ask up to 4 such questions at once, each \
with a helpful "hint" example. (For a question/calculation you otherwise \
build a spreadsheet that COMPUTES the answer with live formulas, e.g. a growth \
table, and surfaces the result.) In that case set:
  status = "needs_input"; notes = one friendly line saying you need a couple of \
details; questions = 1 to 4 SHORT, plain-language questions, each answerable in a \
few words, each with a helpful "hint" example. Do NOT include a spec.
PREFER sensible defaults over asking — most clear requests need NO questions, so \
just build. Never ask more than 4 questions, and never ask about cosmetic \
formatting trivia.
IMPORTANT: if the user has ALREADY answered questions (they appear below under \
"Answers to your questions"), do NOT ask again — make reasonable assumptions and \
BUILD with status = "ready".

LIVE DATA — you can search the web. If the request needs CURRENT real-world values \
(stock / crypto prices, exchange or interest rates, weather, recent statistics, or \
anything phrased as "current / latest / today / now"), use web search and put the \
REAL fetched numbers into the spreadsheet; add a short note in "notes" stating the \
figure(s), their date, and the source. NEVER invent a live figure — if you cannot \
verify it, say so in "notes". Do NOT search for ordinary template or sample-data \
requests.

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
- For a currency SYMBOL, set an explicit "format", e.g. "\\"CHF\\" #,##0.00" or \
"#,##0.00 \\"CHF\\"" (or "$#,##0.00", "\\u20ac#,##0.00"). Use type "percent" for \
percentages (store values as fractions, see below).

FORMULA COLUMNS
- For a calculated column, set type to "formula" and provide a "formula" that is \
an Excel formula string using the literal token {row} for the current 1-based \
row number, e.g. "=B{row}-C{row}" or "=B{row}*C{row}". Column letters refer to \
the FINAL left-to-right column order.
- In every row, the cell for a formula column MUST be null — the renderer fills \
it from the formula. Never put a computed value there.
- Formulas can be rich: cross-sheet refs like ='Q1'!B{row}, IF, VLOOKUP/XLOOKUP, \
SUMIFS, absolute refs ($B$2), and named ranges. You KNOW how many rows you emit, \
so you may write fixed ranges like =SUM($B$2:$B$13) or =VLOOKUP(A{row},$E$2:$F$20,2,FALSE).

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

ATTACHMENTS (images and PDFs)
- The user may attach photos, screenshots, scans, or PDFs that CONTAIN the data. \
Read every attachment carefully and transcribe its tabular data into rows mapped to \
your columns. Preserve numbers, dates, currencies, and labels EXACTLY as shown, and \
infer the column headers from the attachment when the request doesn't name them. \
Combine attachment data with any pasted text. If an attachment is genuinely \
unreadable, or its structure is too ambiguous to map confidently, ask a clarifying \
question (status "needs_input") instead of guessing.

EDITING AN EXISTING SPREADSHEET
- If a CURRENT SPREADSHEET is provided, treat the request as an EDIT instruction. \
Modify that spec accordingly — add / remove / rename columns or sheets, add or change \
rows, sort, add formulas or charts, or merge in new data / attachments as instructed. \
PRESERVE the existing sheets, columns, and data unless the instruction clearly says to \
change or remove them. Always return the COMPLETE updated spec (every sheet), not just \
the delta. If the instruction is ambiguous (e.g. "add this data" without saying which \
sheet or how), ask a clarifying question (status "needs_input").

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

ADVANCED FEATURES (all optional — use ONLY when the request calls for them)
- totalsRow (sheet-level boolean): set true when the user wants a "total" or \
"sum" row — it appends a live SUM row under the data for every numeric column.
- conditionalFormats (sheet-level array): use when the user wants to \
"highlight"/"colour" cells — e.g. the top/bottom values, anything over/under a \
threshold, or a heatmap. Each: column (1-based), rule \
(greaterThan, greaterThanOrEqual, lessThan, lessThanOrEqual, equal, between, \
top10, bottom10, colorScale), value (the compared number/string; omit for \
top10/bottom10/colorScale), value2 (second bound for "between" only), and \
optional color (red, green, yellow, orange, blue, or a 6-hex like "FFC7CE").
- dataValidations (sheet-level array): use when the user wants a "dropdown" or a \
fixed "status"/category list. Each: column (1-based) and values (the allowed \
list, e.g. ["To Do","In Progress","Done"]).
- namedRanges (workbook-level array on spec): use when the user wants a "named \
range" or a "named constant" (e.g. a tax rate). Each: name and ref (an Excel \
reference like "'Settings'!$B$1"); formula columns may then reference name.

HARD LIMITS (never exceed)
- At most 8 sheets. At most 50 columns per sheet. At most 5000 rows per sheet. \
At most 6 charts per sheet. At most 10 value columns per chart.

Keep the spec minimal and explicit: the renderer does exactly what the spec \
says and nothing more. Prefer clarity and correctness over cleverness.

OUTPUT FORMAT — respond with EXACTLY ONE JSON object and NOTHING else: no markdown, \
no code fences, no text before or after. Shape: {"status": "ready" | "needs_input", \
"notes": string, "improvedPrompt": string (ready only), "spec": <SpreadsheetSpec> \
(ready only), "questions": [{"question": string, "hint": string}] (needs_input only)}."""


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Provider chain: free Gemini is primary, Grok (xAI) is the fallback. Reorder or
# limit via the PROVIDERS env var (comma-separated, e.g. "gemini,grok" or "gemini").
# Each provider is skipped if its key is missing; on a transient failure we fall
# through to the next. See docs/DEPLOY.md / .env.example.
def _provider_chain():
    raw = (os.environ.get("PROVIDERS") or "gemini,grok").lower()
    chain = [p.strip() for p in raw.split(",") if p.strip() in ("gemini", "grok")]
    return chain or ["gemini", "grok"]


PROVIDER_CHAIN = _provider_chain()

# Google Gemini (free tier): multimodal (image + PDF) and Google-Search grounding.
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
GEMINI_MODEL = os.environ.get("GEMINI_MODEL") or "gemini-2.5-flash"

# xAI Grok (OpenAI-compatible). Default to the current multimodal flagship so image
# uploads still work on the fallback path; override via XAI_MODEL. NOTE: Grok is a
# PAID fallback — an xAI account with credit is required, else the API returns 403.
XAI_API_KEY = os.environ.get("XAI_API_KEY") or os.environ.get("GROK_API_KEY")
XAI_MODEL = os.environ.get("XAI_MODEL") or "grok-4.3"
XAI_BASE_URL = os.environ.get("XAI_BASE_URL") or "https://api.x.ai/v1"

# Output token ceiling. Kept moderate so generation comfortably finishes inside
# Vercel's 60s function limit; raise via MAX_TOKENS for very large data fills.
try:
    MAX_TOKENS = int(os.environ.get("MAX_TOKENS") or "10000")
except (TypeError, ValueError):
    MAX_TOKENS = 10000

# Hard cap on the request body. Attachments (images / PDFs, base64) ride along, so
# this is larger than the text-only case — but kept UNDER Vercel's ~4.5 MB serverless
# request-body limit so our friendly 413 fires before the platform's opaque one.
MAX_BODY_BYTES = 4 * 1024 * 1024  # 4 MB

# Attachment limits (re-validated server-side; the client also caps + downscales).
MAX_FILES = 6
MAX_FILE_B64 = 3_700_000          # per-file base64 length (~2.8 MB)
MAX_FILES_B64_TOTAL = 3_700_000   # combined base64 length across all attachments
_IMAGE_MEDIA_TYPES = {"image/jpeg", "image/png", "image/webp", "image/gif"}

# Live web search: lets the model pull CURRENT real-world data (prices, rates,
# weather, stats) into the spreadsheet — via Gemini's Google-Search grounding
# (free). Toggle via the WEB_SEARCH env var. (The Grok fallback is text + image only.)
WEB_SEARCH_ENABLED = (os.environ.get("WEB_SEARCH") or "on").strip().lower() not in (
    "off", "0", "false", "no",
)

_CORS_HEADERS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
}


class _InputError(Exception):
    """Bad client input -> HTTP 400 with a safe, human-readable message."""


def _validate_files(files):
    """Validate request `files` and return a clean list of attachment dicts
    {kind, media_type, data} (base64). Each provider converts these to its own
    format. Raises _InputError (-> 400) on bad type / media_type / size / count.
    """
    if files is None:
        return []
    if not isinstance(files, list):
        raise _InputError("Attached files are malformed.")
    if len(files) > MAX_FILES:
        raise _InputError("Too many attachments (max %d)." % MAX_FILES)

    clean = []
    total = 0
    for f in files:
        if not isinstance(f, dict):
            raise _InputError("Attached files are malformed.")
        kind = f.get("type")
        media_type = f.get("media_type")
        data = f.get("data")
        if not isinstance(data, str) or not data:
            raise _InputError("An attachment is empty or unreadable.")
        if len(data) > MAX_FILE_B64:
            raise _InputError("An attachment is too large. Try a smaller or clearer file.")
        total += len(data)
        if total > MAX_FILES_B64_TOTAL:
            raise _InputError("The attachments are too large together. Use fewer or smaller files.")
        if kind == "image":
            if media_type not in _IMAGE_MEDIA_TYPES:
                raise _InputError("Unsupported image type.")
        elif kind == "pdf":
            if media_type != "application/pdf":
                raise _InputError("Unsupported document type.")
        else:
            raise _InputError("Unsupported attachment type.")
        clean.append({"kind": kind, "media_type": media_type, "data": data})
    return clean


class _ProviderError(Exception):
    """A provider call failed. `reason` is one of:
    rate_limit | auth | quota | error | no_key."""

    def __init__(self, reason, detail=""):
        super().__init__(reason)
        self.reason = reason
        self.detail = detail


def _gemini_reason(exc):
    code = getattr(exc, "code", None) or getattr(exc, "status_code", None)
    msg = str(getattr(exc, "message", "") or exc).lower()
    if code == 429 or "resource_exhausted" in msg or "quota" in msg or "rate limit" in msg:
        return "rate_limit"
    if code in (401, 403) or "permission" in msg or "api key" in msg or "unauthenticated" in msg:
        return "auth"
    return "error"


def _call_gemini(system, user_text, files):
    """Call Google Gemini (free tier). Returns the model's text. Raises
    _ProviderError. Supports images + PDFs and Google-Search grounding."""
    if not GEMINI_API_KEY:
        raise _ProviderError("no_key")

    parts = [gtypes.Part(text=user_text)]
    for f in files:
        try:
            raw = base64.b64decode(f["data"])
        except Exception:  # noqa: BLE001 — skip an undecodable attachment
            continue
        parts.append(gtypes.Part.from_bytes(data=raw, mime_type=f["media_type"]))

    cfg = {"system_instruction": system, "max_output_tokens": MAX_TOKENS, "temperature": 0.4}
    if WEB_SEARCH_ENABLED:
        cfg["tools"] = [gtypes.Tool(google_search=gtypes.GoogleSearch())]
    if hasattr(gtypes, "ThinkingConfig"):
        # Keep the full token budget for the JSON (this is structured output, not a
        # deep-reasoning task) and stay fast inside Vercel's 60s limit.
        cfg["thinking_config"] = gtypes.ThinkingConfig(thinking_budget=0)

    try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        resp = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=[gtypes.Content(role="user", parts=parts)],
            config=gtypes.GenerateContentConfig(**cfg),
        )
    except gerrors.APIError as exc:
        raise _ProviderError(_gemini_reason(exc), str(exc)[:200])
    except Exception as exc:  # noqa: BLE001
        raise _ProviderError("error", str(exc)[:200])

    text = getattr(resp, "text", None)
    if not text:
        try:
            text = "".join(
                p.text for c in (resp.candidates or []) for p in (c.content.parts or [])
                if getattr(p, "text", None)
            )
        except Exception:  # noqa: BLE001
            text = None
    if not text:
        raise _ProviderError("error", "empty response")
    return text


def _call_grok(system, user_text, files):
    """Call xAI Grok (OpenAI-compatible). Returns the model's text. Raises
    _ProviderError. Sends images inline; PDFs aren't supported by the chat API,
    so they're noted rather than read."""
    if not XAI_API_KEY:
        raise _ProviderError("no_key")

    content = [{"type": "text", "text": user_text}]
    skipped_pdf = False
    for f in files:
        if f["kind"] == "image":
            content.append({
                "type": "image_url",
                "image_url": {"url": "data:%s;base64,%s" % (f["media_type"], f["data"])},
            })
        else:
            skipped_pdf = True
    if skipped_pdf:
        content.append({"type": "text", "text": "(A PDF was attached but can't be read on "
                                                 "this fallback path; use the rest of the input.)"})

    try:
        client = OpenAI(api_key=XAI_API_KEY, base_url=XAI_BASE_URL)
        resp = client.chat.completions.create(
            model=XAI_MODEL,
            messages=[{"role": "system", "content": system},
                      {"role": "user", "content": content}],
            max_tokens=MAX_TOKENS,
        )
    except openai.RateLimitError as exc:
        raise _ProviderError("rate_limit", str(exc)[:200])
    except openai.AuthenticationError as exc:
        raise _ProviderError("auth", str(exc)[:200])
    except openai.APIError as exc:
        raise _ProviderError("error", str(exc)[:200])
    except Exception as exc:  # noqa: BLE001
        raise _ProviderError("error", str(exc)[:200])

    try:
        text = resp.choices[0].message.content
    except Exception:  # noqa: BLE001
        text = None
    if not text:
        raise _ProviderError("error", "empty response")
    return text


def _generate(system, user_text, files):
    """Try each provider in PROVIDER_CHAIN; return the first text response. Raises
    _ProviderError with the most actionable reason if every provider fails."""
    seen = []
    for provider in PROVIDER_CHAIN:
        try:
            if provider == "gemini":
                return _call_gemini(system, user_text, files)
            if provider == "grok":
                return _call_grok(system, user_text, files)
        except _ProviderError as pe:
            seen.append(pe.reason)
            continue
    for reason in ("auth", "quota", "rate_limit", "error", "no_key"):
        if reason in seen:
            raise _ProviderError(reason)
    raise _ProviderError("error")


def _balanced_objects(text):
    """Yield each balanced top-level {...} substring, respecting string literals
    and escapes (so braces inside strings don't confuse the scan)."""
    objs = []
    depth, start, in_str, esc = 0, -1, False, False
    for k, ch in enumerate(text):
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
            continue
        if ch == '"':
            in_str = True
        elif ch == "{":
            if depth == 0:
                start = k
            depth += 1
        elif ch == "}" and depth > 0:
            depth -= 1
            if depth == 0 and start != -1:
                objs.append(text[start:k + 1])
                start = -1
    return objs


def _extract_json(text):
    """Parse the model's JSON envelope, tolerating a stray code fence, a stray
    leading brace, or prose around it. Returns the envelope dict (preferring one
    with a "status" field), or None if nothing parseable is found."""
    if not isinstance(text, str):
        return None
    candidates = [text.strip()]
    fenced = re.search(r"```(?:json)?\s*(\{.*\})\s*```", text, re.DOTALL)
    if fenced:
        candidates.append(fenced.group(1))
    candidates.extend(_balanced_objects(text))
    i, j = text.find("{"), text.rfind("}")
    if i != -1 and j > i:
        candidates.append(text[i:j + 1])

    parsed = []
    for cand in candidates:
        try:
            obj = json.loads(cand)
        except (ValueError, TypeError):
            continue
        if isinstance(obj, dict):
            parsed.append(obj)
    if not parsed:
        return None
    for obj in parsed:
        if "status" in obj:
            return obj
    return parsed[0]


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
            if not (GEMINI_API_KEY or XAI_API_KEY):
                self._send_json(
                    500,
                    {"error": "The server is not configured yet (no AI provider key)."},
                )
                return

            # --- Build the user message -------------------------------------
            data = payload.get("data")
            user_text = "Request:\n" + prompt.strip()

            locale = payload.get("locale")
            if isinstance(locale, str) and locale.strip():
                user_text += (
                    "\n\nUser locale: " + locale.strip()[:32]
                    + " (use it to recommend a local currency / number format when relevant)."
                )

            base_spec = payload.get("baseSpec")
            if isinstance(base_spec, dict) and isinstance(base_spec.get("sheets"), list):
                user_text += (
                    "\n\nCURRENT SPREADSHEET to edit — apply the request to THIS and "
                    "return the COMPLETE updated spec (keep existing data and columns "
                    "unless the request says to change them):\n" + json.dumps(base_spec)
                )

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

            # --- Attachments (images / PDFs) --------------------------------
            try:
                files = _validate_files(payload.get("files"))
            except _InputError as exc:
                self._send_json(400, {"error": str(exc)})
                return
            if files:
                user_text += (
                    "\n\nThe data to use is in the attached file(s) — read them "
                    "carefully and transcribe the data into the rows."
                )

            # --- Call the AI provider (Gemini -> Grok) ----------------------
            try:
                text = _generate(SYSTEM_PROMPT, user_text, files)
            except _ProviderError as pe:
                traceback.print_exc(file=sys.stderr)
                if pe.reason == "rate_limit":
                    self._send_json(
                        503,
                        {"error": "The AI is over its usage limit right now "
                                  "(free-tier rate limit). Please try again in a few minutes."},
                    )
                elif pe.reason in ("auth", "quota", "no_key"):
                    self._send_json(
                        503,
                        {"error": "The AI is temporarily unavailable. Please try again "
                                  "later — your work isn't lost."},
                    )
                else:
                    self._send_json(
                        500,
                        {"error": "Something went wrong generating your spreadsheet plan. "
                                  "Please try again."},
                    )
                return

            # --- Parse the JSON envelope ------------------------------------
            # The model returns one JSON object (per the OUTPUT FORMAT instruction);
            # parse tolerantly (grounding citations may add surrounding prose).
            result = _extract_json(text)
            if not isinstance(result, dict):
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
