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
  "data": "string|null — raw pasted tabular text (CSV/TSV/lines), optional"
}
```

**Success (200)**
```json
{
  "improvedPrompt": "string — a clear, specific restatement of the workbook to build",
  "notes": "string — one or two friendly sentences for the user (plain language)",
  "spec": { "...": "a valid SpreadsheetSpec (see §2)" }
}
```

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
  "sheets": [ Sheet, ... ]           // 1..8 sheets
}

Sheet = {
  "name": "string",                  // Excel tab name; generator truncates to 31 chars, dedupes
  "description": "string|null",       // optional human note (not rendered into cells)
  "columns": [ Column, ... ],        // 1..50 columns, left-to-right order
  "rows": [ Row, ... ],              // optional; [] means an empty, labelled template
  "freezeHeader": true,              // optional, default true — freezes the header row
  "autoFilter": true,                // optional, default true — adds filter dropdowns
  "charts": [ Chart, ... ]           // optional, default []
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

### Hard limits (generator validates; reject with 400 if exceeded)
- `sheets`: 1–8 · `columns` per sheet: 1–50 · `rows` per sheet: 0–5000
- `charts` per sheet: 0–6 · chart `valueColumns`: 1–10
- All column indices in a chart must be in range for that sheet.

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
