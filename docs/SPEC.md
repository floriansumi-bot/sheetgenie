# SheetGenie — Shared Contract (SPEC)

> **This file is the single source of truth.** Every component — the frontend, the
> `/api/improve` endpoint, and the `/api/generate` endpoint — is built against the
> contracts below. If you change a contract here, every consumer must change with it.

There are two contracts:

1. **The HTTP API** — how the browser talks to the two serverless functions.
2. **The `SpreadsheetSpec`** — the JSON "blueprint" of a workbook. The AI produces it,
   and the generator renders it deterministically into a real `.xlsx`. This is the
   common language that lets the AI step and the Excel step evolve independently.

---

## 1. HTTP API contract

Both endpoints are Vercel Python serverless functions using the native
`http.server.BaseHTTPRequestHandler` pattern (class **must** be named `handler`).
Same-origin in production; permissive CORS headers are sent so local dev on a
different port also works. Both accept `OPTIONS` (CORS preflight) and `POST`.

### `POST /api/improve`  →  JSON

Turns a rough human prompt into a polished prompt **and** a complete `SpreadsheetSpec`.
This is the only endpoint that calls the Anthropic API.

**Request body**
```json
{
  "prompt": "string  — the raw user prompt (typed or dictated)",
  "hasData": true,
  "data": "string|null — raw pasted tabular text (CSV/TSV/lines), optional",
  "baseSpec": { "...": "an existing SpreadsheetSpec to EDIT (optional)" },
  "locale": "string|null — e.g. \"de-CH\"; lets the model recommend a local currency / number format",
  "files": [
    {
      "type": "image | pdf",
      "media_type": "image/jpeg | image/png | image/webp | image/gif | application/pdf",
      "data": "base64-encoded bytes (NO 'data:' prefix)",
      "name": "string — optional original filename"
    }
  ],
  "clarifications": [
    { "question": "string — a question we previously asked", "answer": "string — the user's answer" }
  ]
}
```
`files` is optional. The client may attach photos / screenshots / scans (sent to the
model as **image** blocks) and PDFs (sent as **document** blocks); the model reads them
to extract the data and fill the rows. Images are downscaled client-side and the **total
request body is kept under ~4 MB** (the serverless limit) — at most **6 attachments**.
Allowed media types: the four image types above plus `application/pdf`. The server
re-validates type, count, and size and returns `400`/`413` otherwise.

`baseSpec` is optional. When present, the request is an **edit**: the model modifies
that spec per `prompt` (the change / import instruction) plus any new `data`/`files`,
and returns the COMPLETE updated spec (`status: "ready"`) — or asks a question if the
instruction is ambiguous. This powers the in-app library, which stores past specs in
the browser (IndexedDB) so they can be re-opened, edited, and re-downloaded.

`clarifications` is optional / `null` on the first call. If the previous response was
`status: "needs_input"`, the client re-calls with the same `prompt`/`data` plus the
user's answers here.

**Success (200)** — one of two shapes, discriminated by `status`:

*Ready to build:*
```json
{
  "status": "ready",
  "notes": "string — one or two friendly sentences for the user (plain language)",
  "improvedPrompt": "string — a clear, specific restatement of the workbook to build",
  "spec": { "...": "a valid SpreadsheetSpec (see §2)" }
}
```

*Needs clarification (the prompt/data is genuinely ambiguous):*
```json
{
  "status": "needs_input",
  "notes": "string — one friendly line explaining why you're asking",
  "questions": [
    { "question": "string — short, plain-language", "hint": "string|null — example/placeholder" }
  ]
}
```
Rules: ask **only** when a wrong assumption would produce the wrong spreadsheet
(unclear scope, missing key fields, ambiguous data mapping, template-vs-filled
unclear). **Max 4 questions**, each answerable in a few words. Prefer sensible
defaults over asking. Once the client sends `clarifications`, the response **must**
be `status: "ready"` — never re-ask.

**Error (4xx/5xx)**
```json
{ "error": "human-readable message" }
```
- `400` — missing/empty `prompt`, or malformed JSON.
- `500` — missing `ANTHROPIC_API_KEY`, upstream API error, or the model returned
  something that is not a valid spec. The message must be safe to show a user
  (never leak the key or a raw stack trace).

### `POST /api/generate`  →  binary `.xlsx`

Pure, deterministic rendering. **No AI, no network** — just openpyxl. Fast and free.

**Request body**
```json
{
  "spec": { "...": "a SpreadsheetSpec (see §2)" },
  "filename": "string|null — optional base name, no extension"
}
```

**Success (200)** — the raw `.xlsx` bytes, with headers:
```
Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
Content-Disposition: attachment; filename="<safe-name>.xlsx"
```
`<safe-name>` is derived from `filename` or `spec.title`, slugified, `.xlsx` appended.

**Error (4xx/5xx)** — JSON `{ "error": "..." }` (`400` for an invalid spec).

---

## 2. `SpreadsheetSpec` schema

A spec describes one workbook. Keep it small and explicit — the generator does
exactly what the spec says and nothing more.

```jsonc
SpreadsheetSpec = {
  "title": "string",                 // workbook title; default filename if none given
  "sheets": [ Sheet, ... ],          // 1..8 sheets
  "namedRanges": [ NamedRange, ... ] // optional, default [] — workbook-level defined names
}

Sheet = {
  "name": "string",                  // Excel tab name; generator truncates to 31 chars, dedupes
  "description": "string|null",       // optional human note (not rendered into cells)
  "columns": [ Column, ... ],        // 1..50 columns, left-to-right order
  "rows": [ Row, ... ],              // optional; [] means an empty, labelled template
  "freezeHeader": true,              // optional, default true — freezes the header row
  "autoFilter": true,                // optional, default true — adds filter dropdowns
  "charts": [ Chart, ... ],          // optional, default []
  "totalsRow": false,                // optional — append a live SUM totals row (see below)
  "conditionalFormats": [ CondFmt, ... ],  // optional — colour rules on a column's data
  "dataValidations": [ Validation, ... ]   // optional — dropdown lists on a column's data
}

NamedRange = { "name": "TaxRate", "ref": "'Settings'!$B$1" }  // formulas may use the name

CondFmt = {
  "column": 2,                       // 1-based column the rule applies to (its data rows)
  "rule": "greaterThan|greaterThanOrEqual|lessThan|lessThanOrEqual|equal|between|top10|bottom10|colorScale",
  "value": 500,                      // compared-against number/string (omit for top10/bottom10/colorScale)
  "value2": null,                    // optional second bound for "between"
  "color": "red"                     // red|green|yellow|orange|blue, or a 6-hex like "FFC7CE"
}

Validation = {
  "column": 4,                       // 1-based column to attach a dropdown to (its data rows)
  "values": ["To Do", "In Progress", "Done"]   // allowed list; entries outside it are rejected
}

Row = [ cell, cell, ... ]            // aligned to `columns` by index; pad/truncate to len(columns)
cell = string | number | boolean | null

Column = {
  "header": "string",                // the column label written in row 1 (bold)
  "type": "text|number|currency|percent|date|formula",
  "width": 18,                       // optional char width; generator picks a sensible default
  "format": "string|null",           // optional explicit Excel number format, overrides `type`
  "formula": "string|null"           // required iff type=="formula"; see below
}

Chart = {
  "type": "bar|line|pie",
  "title": "string",
  "categoriesColumn": 1,             // 1-based column index used for category labels / pie slices
  "valueColumns": [2, 3],            // 1-based column indices for the numeric series
  "dataStartRow": 2,                 // optional, 1-based; default 2 (first row after header)
  "dataEndRow": null,                // optional, 1-based; default = last populated data row
  "anchor": "H2"                     // optional cell anchor; generator auto-places if omitted
}
```

### Column `type` → Excel number format (generator applies)
| type       | number_format            | notes |
|------------|--------------------------|-------|
| `text`     | `General`                | left-aligned |
| `number`   | `#,##0.00`               | |
| `currency` | `#,##0.00`               | header may carry a currency hint; no locale assumed |
| `percent`  | `0.0%`                   | values are fractions (0.25 → 25.0%) |
| `date`     | `yyyy-mm-dd`             | accepts ISO strings or Excel serials |
| `formula`  | inherits from neighbours | the cell value is the rendered `formula` |

### Formula templating
For `type=="formula"`, `formula` is an Excel formula string using the literal
token `{row}` for the current 1-based row number. The generator substitutes `{row}`
per data row. Example: `"=B{row}*C{row}"` on row 5 becomes `=B5*C5`.
Column letters refer to the final left-to-right column order.

### Advanced features (optional)
- **`totalsRow`** — when true, the generator appends one row below the data that
  `=SUM`s every numeric / currency / percent / formula column over the data range,
  labels the first text column "Total", and bolds the row with a top border. The SUMs
  are live formulas. From another sheet, reference a sheet's total with a formula
  column, e.g. `"=SUM('Q1'!B2:B100)"`.
- **`conditionalFormats`** — real Excel conditional-formatting rules on the column's
  data range (never one-off shading). `top10`/`bottom10` highlight the top/bottom 10%;
  `colorScale` applies a 3-colour scale; comparison rules use `value` (+ `value2` for
  `between`).
- **`dataValidations`** — a real dropdown/list validation on the column's data range;
  entries outside `values` are rejected by Excel.
- **`namedRanges`** — workbook-level defined names; formula columns may reference the
  name instead of a literal range.

### Number formats & formulas (guidance for the model)
- For a currency symbol use an explicit `format`, e.g. `"\"CHF\" #,##0.00"` or
  `"#,##0.00 \"CHF\""`. Use `type:"percent"` for percentages (store fractions).
- Formula columns may use cross-sheet refs (`='Q1'!B{row}`), `IF`, `VLOOKUP`/`XLOOKUP`,
  `SUMIFS`, absolute refs (`$B$2`), and named ranges. `{row}` expands per data row, and
  the model knows how many rows it emits, so it can write ranges like `$B$2:$B$13`.

### Hard limits (generator validates; reject with 400 if exceeded)
- `sheets`: 1–8 · `columns` per sheet: 1–50 · `rows` per sheet: 0–5000
- `charts` per sheet: 0–6 · chart `valueColumns`: 1–10
- `conditionalFormats` per sheet: 0–20 · `dataValidations` per sheet: 0–20 ·
  validation `values`: 1–200 · `namedRanges`: 0–50
- All column indices (charts, conditional formats, validations) must be in range.

### Worked example
Prompt: *"monthly budget tracker with categories, budgeted vs actual, and a bar chart"*
```json
{
  "title": "Monthly Budget Tracker",
  "sheets": [{
    "name": "Budget",
    "columns": [
      { "header": "Category",  "type": "text" },
      { "header": "Budgeted",  "type": "currency" },
      { "header": "Actual",    "type": "currency" },
      { "header": "Variance",  "type": "formula", "formula": "=B{row}-C{row}" }
    ],
    "rows": [
      ["Rent", 1500, 1500, null],
      ["Groceries", 400, 462.30, null],
      ["Transport", 120, 98.5, null]
    ],
    "freezeHeader": true,
    "autoFilter": true,
    "charts": [
      { "type": "bar", "title": "Budgeted vs Actual",
        "categoriesColumn": 1, "valueColumns": [2, 3] }
    ]
  }]
}
```
(`Variance` cells are `null` in `rows` because the generator fills them from the
`formula` — formula columns ignore any value supplied in `rows`.)
