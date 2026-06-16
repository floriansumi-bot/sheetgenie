"""
SheetGenie — POST /api/generate

Pure, deterministic rendering of a SpreadsheetSpec (see docs/SPEC.md) into a real
.xlsx workbook using openpyxl. NO AI, NO network calls. Vercel Python serverless
function using the native http.server.BaseHTTPRequestHandler pattern.

Contract (docs/SPEC.md §1):
  Request body : { "spec": <SpreadsheetSpec>, "filename": "string|null" }
  Success 200  : raw .xlsx bytes, attachment download headers
  Error 4xx/5xx: JSON { "error": "<safe human-readable message>" }
"""

import json
import re
import sys
import datetime
import traceback
from io import BytesIO
from http.server import BaseHTTPRequestHandler

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment


# ---------------------------------------------------------------------------
# Limits (docs/SPEC.md §2 "Hard limits")
# ---------------------------------------------------------------------------
MAX_SHEETS = 8
MIN_SHEETS = 1
MAX_COLUMNS = 50
MIN_COLUMNS = 1
MAX_ROWS = 5000
MIN_ROWS = 0
MAX_CHARTS = 6
MIN_CHARTS = 0
MAX_VALUE_COLUMNS = 10
MIN_VALUE_COLUMNS = 1

# Abuse / resource caps (defense-in-depth; the body cap is the primary guard).
MAX_BODY_BYTES = 4 * 1024 * 1024   # 4 MB request body
MAX_CELL_CHARS = 32767             # Excel's hard per-cell character limit
MAX_FORMULA_CHARS = 8192           # Excel's formula length limit
MAX_HEADER_CHARS = 255
MAX_TOTAL_CELLS = 200_000          # across the whole workbook

VALID_COLUMN_TYPES = {"text", "number", "currency", "percent", "date", "formula"}
VALID_CHART_TYPES = {"bar", "line", "pie"}

# Column type -> Excel number format (docs/SPEC.md §2 table).
TYPE_FORMATS = {
    "text": "General",
    "number": "#,##0.00",
    "currency": "#,##0.00",
    "percent": "0.0%",
    "date": "yyyy-mm-dd",
    # "formula" inherits from neighbours -> no forced format.
}

HEADER_FILL = PatternFill(start_color="FFE8EEF7", end_color="FFE8EEF7", fill_type="solid")
HEADER_FONT = Font(bold=True)

MIN_COL_WIDTH = 10
MAX_COL_WIDTH = 40


class SpecError(ValueError):
    """Raised for any spec validation failure -> 400."""


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------
def _validate_spec(spec):
    """Validate a SpreadsheetSpec against the hard limits in SPEC.md.

    Raises SpecError (-> 400) with a clear reason on any failure.
    """
    if not isinstance(spec, dict):
        raise SpecError("spec must be an object.")

    sheets = spec.get("sheets")
    if not isinstance(sheets, list):
        raise SpecError("spec.sheets must be a list.")
    if not (MIN_SHEETS <= len(sheets) <= MAX_SHEETS):
        raise SpecError(
            "spec must have between %d and %d sheets (got %d)."
            % (MIN_SHEETS, MAX_SHEETS, len(sheets))
        )

    total_cells = 0

    for si, sheet in enumerate(sheets):
        where = "sheet %d" % (si + 1)
        if not isinstance(sheet, dict):
            raise SpecError("%s must be an object." % where)

        columns = sheet.get("columns")
        if not isinstance(columns, list):
            raise SpecError("%s: columns must be a list." % where)
        if not (MIN_COLUMNS <= len(columns) <= MAX_COLUMNS):
            raise SpecError(
                "%s: must have between %d and %d columns (got %d)."
                % (where, MIN_COLUMNS, MAX_COLUMNS, len(columns))
            )
        n_cols = len(columns)

        for ci, col in enumerate(columns):
            cwhere = "%s, column %d" % (where, ci + 1)
            if not isinstance(col, dict):
                raise SpecError("%s must be an object." % cwhere)
            ctype = col.get("type")
            if ctype not in VALID_COLUMN_TYPES:
                raise SpecError(
                    "%s: invalid type %r (allowed: %s)."
                    % (cwhere, ctype, ", ".join(sorted(VALID_COLUMN_TYPES)))
                )
            header = col.get("header")
            if isinstance(header, str) and len(header) > MAX_HEADER_CHARS:
                raise SpecError("%s: header is too long." % cwhere)
            if ctype == "formula":
                formula = col.get("formula")
                if not isinstance(formula, str) or not formula.strip():
                    raise SpecError(
                        "%s: type 'formula' requires a non-empty 'formula' string."
                        % cwhere
                    )
                if len(formula) > MAX_FORMULA_CHARS:
                    raise SpecError("%s: formula is too long." % cwhere)

        rows = sheet.get("rows")
        if rows is None:
            rows = []
        if not isinstance(rows, list):
            raise SpecError("%s: rows must be a list." % where)
        if not (MIN_ROWS <= len(rows) <= MAX_ROWS):
            raise SpecError(
                "%s: must have between %d and %d rows (got %d)."
                % (where, MIN_ROWS, MAX_ROWS, len(rows))
            )

        # Per-cell string-length cap (Excel's own limit) — bounds individual
        # cell size even within a shape-valid spec.
        for ri, row in enumerate(rows):
            if not isinstance(row, list):
                raise SpecError("%s: row %d must be a list." % (where, ri + 1))
            for cell in row:
                if isinstance(cell, str) and len(cell) > MAX_CELL_CHARS:
                    raise SpecError(
                        "%s: a cell value exceeds the %d-character limit."
                        % (where, MAX_CELL_CHARS)
                    )

        # Total-cell budget across the whole workbook (resource guard).
        total_cells += n_cols * len(rows)
        if total_cells > MAX_TOTAL_CELLS:
            raise SpecError(
                "The spreadsheet is too large (over %d cells). "
                "Please request fewer rows or columns." % MAX_TOTAL_CELLS
            )

        charts = sheet.get("charts")
        if charts is None:
            charts = []
        if not isinstance(charts, list):
            raise SpecError("%s: charts must be a list." % where)
        if not (MIN_CHARTS <= len(charts) <= MAX_CHARTS):
            raise SpecError(
                "%s: must have between %d and %d charts (got %d)."
                % (where, MIN_CHARTS, MAX_CHARTS, len(charts))
            )

        for chi, chart in enumerate(charts):
            chwhere = "%s, chart %d" % (where, chi + 1)
            if not isinstance(chart, dict):
                raise SpecError("%s must be an object." % chwhere)
            if chart.get("type") not in VALID_CHART_TYPES:
                raise SpecError(
                    "%s: invalid chart type %r (allowed: %s)."
                    % (chwhere, chart.get("type"), ", ".join(sorted(VALID_CHART_TYPES)))
                )

            cat_col = chart.get("categoriesColumn")
            if not isinstance(cat_col, int) or isinstance(cat_col, bool):
                raise SpecError("%s: categoriesColumn must be an integer." % chwhere)
            if not (1 <= cat_col <= n_cols):
                raise SpecError(
                    "%s: categoriesColumn %d out of range (1..%d)."
                    % (chwhere, cat_col, n_cols)
                )

            value_cols = chart.get("valueColumns")
            if not isinstance(value_cols, list):
                raise SpecError("%s: valueColumns must be a list." % chwhere)
            if not (MIN_VALUE_COLUMNS <= len(value_cols) <= MAX_VALUE_COLUMNS):
                raise SpecError(
                    "%s: valueColumns must have between %d and %d entries (got %d)."
                    % (chwhere, MIN_VALUE_COLUMNS, MAX_VALUE_COLUMNS, len(value_cols))
                )
            for vc in value_cols:
                if not isinstance(vc, int) or isinstance(vc, bool):
                    raise SpecError("%s: valueColumns entries must be integers." % chwhere)
                if not (1 <= vc <= n_cols):
                    raise SpecError(
                        "%s: valueColumns entry %d out of range (1..%d)."
                        % (chwhere, vc, n_cols)
                    )

            # Optional row bounds, if given, must be sane integers.
            for key in ("dataStartRow", "dataEndRow"):
                val = chart.get(key)
                if val is None:
                    continue
                if not isinstance(val, int) or isinstance(val, bool) or val < 1:
                    raise SpecError("%s: %s must be a positive integer." % (chwhere, key))


# ---------------------------------------------------------------------------
# Rendering helpers
# ---------------------------------------------------------------------------
def _unique_title(name, used):
    """Truncate a sheet name to 31 chars and de-duplicate against `used`."""
    if not isinstance(name, str) or not name.strip():
        base = "Sheet"
    else:
        base = name.strip()
    base = base[:31]
    title = base
    n = 2
    while title in used or not title:
        # Reserve room for the suffix so the result stays <= 31 chars.
        suffix = " (%d)" % n
        title = (base[: 31 - len(suffix)] + suffix) if base else ("Sheet%d" % n)
        n += 1
    used.add(title)
    return title


def _default_width(header):
    """Sensible default column width derived from the header length, clamped."""
    length = len(header) if isinstance(header, str) else MIN_COL_WIDTH
    return max(MIN_COL_WIDTH, min(MAX_COL_WIDTH, length + 2))


def _coerce_date(value):
    """Coerce an ISO string to a date/datetime object; otherwise return as-is."""
    if isinstance(value, str):
        s = value.strip()
        try:
            return datetime.date.fromisoformat(s)
        except ValueError:
            pass
        try:
            return datetime.datetime.fromisoformat(s)
        except ValueError:
            pass
    return value


def _render_sheet(ws, sheet):
    """Render one Sheet spec onto an openpyxl worksheet."""
    columns = sheet["columns"]
    rows = sheet.get("rows") or []
    n_cols = len(columns)

    # --- Headers (row 1) ---
    for ci, col in enumerate(columns, start=1):
        header = col.get("header")
        header = header if isinstance(header, str) else ""
        cell = ws.cell(row=1, column=ci, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=False)

        # Column width: explicit width else sensible default from header length.
        width = col.get("width")
        letter = get_column_letter(ci)
        if isinstance(width, (int, float)) and not isinstance(width, bool) and width > 0:
            ws.column_dimensions[letter].width = float(width)
        else:
            ws.column_dimensions[letter].width = _default_width(header)

    # --- Data rows (start at row 2) ---
    first_data_row = 2
    last_data_row = first_data_row + len(rows) - 1  # == 1 when rows is empty

    for ri, row in enumerate(rows):
        actual_row = first_data_row + ri
        # Pad/truncate the row to exactly len(columns).
        cells = list(row) if isinstance(row, list) else []
        if len(cells) < n_cols:
            cells = cells + [None] * (n_cols - len(cells))
        else:
            cells = cells[:n_cols]

        for ci, col in enumerate(columns, start=1):
            ctype = col.get("type")
            target = ws.cell(row=actual_row, column=ci)

            if ctype == "formula":
                # Formula columns ignore any value supplied in rows.
                formula = col.get("formula") or ""
                target.value = formula.replace("{row}", str(actual_row))
            else:
                value = cells[ci - 1]
                if ctype == "date":
                    value = _coerce_date(value)
                target.value = value
                # Formula-injection defense: openpyxl turns a leading "=" string
                # into a live formula. In a non-formula column, force any such
                # cell back to a literal string so user data can never execute.
                if target.data_type == "f":
                    target.data_type = "s"

            # Number format: explicit column.format overrides the type default.
            explicit = col.get("format")
            if isinstance(explicit, str) and explicit:
                target.number_format = explicit
            else:
                fmt = TYPE_FORMATS.get(ctype)
                if fmt:
                    target.number_format = fmt

            if ctype == "text":
                target.alignment = Alignment(horizontal="left")

    # --- Freeze header (default true) ---
    if sheet.get("freezeHeader", True):
        ws.freeze_panes = "A2"

    # --- Auto filter (default true) ---
    if sheet.get("autoFilter", True):
        end_col = get_column_letter(n_cols)
        end_row = last_data_row if rows else 1
        ws.auto_filter.ref = "A1:%s%d" % (end_col, end_row)

    # --- Charts (skip entirely if there are no data rows) ---
    if rows:
        _render_charts(ws, sheet, columns, first_data_row, last_data_row)


def _render_charts(ws, sheet, columns, first_data_row, last_data_row):
    n_cols = len(columns)
    charts = sheet.get("charts") or []
    for chart_spec in charts:
        ctype = chart_spec["type"]
        if ctype == "bar":
            chart = BarChart()
        elif ctype == "line":
            chart = LineChart()
        else:  # "pie"
            chart = PieChart()

        chart.title = chart_spec.get("title") or None

        cat_col = chart_spec["categoriesColumn"]
        data_start = chart_spec.get("dataStartRow") or first_data_row
        data_end = chart_spec.get("dataEndRow") or last_data_row
        # Clamp to the populated range so an over-specified bound can't extend
        # the series over empty cells (trailing zero/blank points).
        data_start = max(data_start, first_data_row)
        data_end = min(data_end, last_data_row)
        if data_end < data_start:
            continue  # nothing to plot

        categories = Reference(
            ws, min_col=cat_col, min_row=data_start, max_row=data_end
        )

        value_cols = chart_spec["valueColumns"]
        if ctype == "pie":
            value_cols = value_cols[:1]  # pie uses only the first value column

        for vc in value_cols:
            # min_row=1 so the header row becomes the series name.
            series_ref = Reference(ws, min_col=vc, min_row=1, max_row=data_end)
            chart.add_data(series_ref, titles_from_data=True)

        chart.set_categories(categories)

        anchor = chart_spec.get("anchor")
        if not (isinstance(anchor, str) and anchor):
            anchor = "%s2" % get_column_letter(n_cols + 2)
        ws.add_chart(chart, anchor)


def _slugify_filename(filename, spec):
    """Build a safe download filename: slug.xlsx."""
    candidate = None
    if isinstance(filename, str) and filename.strip():
        candidate = filename
    elif isinstance(spec, dict) and isinstance(spec.get("title"), str) and spec["title"].strip():
        candidate = spec["title"]
    else:
        candidate = "spreadsheet"

    slug = candidate.strip().lower()
    slug = re.sub(r"[^a-z0-9]+", "-", slug)
    slug = re.sub(r"-+", "-", slug).strip("-")
    if not slug:
        slug = "spreadsheet"
    return slug + ".xlsx"


def _build_workbook(spec):
    """Render a validated SpreadsheetSpec into .xlsx bytes."""
    wb = Workbook()
    used_titles = set()

    for idx, sheet in enumerate(spec["sheets"]):
        if idx == 0:
            ws = wb.active  # reuse the default sheet for the first
        else:
            ws = wb.create_sheet()
        ws.title = _unique_title(sheet.get("name"), used_titles)
        _render_sheet(ws, sheet)

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# HTTP handler
# ---------------------------------------------------------------------------
class handler(BaseHTTPRequestHandler):
    def _send_json(self, status, payload):
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_error(self, status, message):
        self._send_json(status, {"error": message})

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_POST(self):
        try:
            # --- Read + parse body ---
            try:
                length = int(self.headers.get("Content-Length") or 0)
            except (TypeError, ValueError):
                length = 0
            if length > MAX_BODY_BYTES:
                self._send_error(413, "Request body is too large.")
                return
            raw = self.rfile.read(min(length, MAX_BODY_BYTES)) if length > 0 else b""

            try:
                payload = json.loads(raw.decode("utf-8")) if raw else None
            except (ValueError, UnicodeDecodeError):
                self._send_error(400, "Request body is not valid JSON.")
                return

            if not isinstance(payload, dict):
                self._send_error(400, "Request body must be a JSON object.")
                return

            spec = payload.get("spec")
            filename = payload.get("filename")

            # --- Validate spec ---
            try:
                _validate_spec(spec)
            except SpecError as e:
                self._send_error(400, str(e))
                return

            # --- Render ---
            try:
                data = _build_workbook(spec)
            except SpecError as e:
                self._send_error(400, str(e))
                return

            safe_name = _slugify_filename(filename, spec)

            # --- Respond with the .xlsx bytes ---
            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            self.send_header(
                "Content-Disposition", 'attachment; filename="%s"' % safe_name
            )
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(data)

        except Exception:  # noqa: BLE001 — last-resort safety net
            traceback.print_exc(file=sys.stderr)
            try:
                self._send_error(
                    500, "Failed to generate the spreadsheet. Please try again."
                )
            except Exception:  # noqa: BLE001 — headers may already be sent
                traceback.print_exc(file=sys.stderr)
