<h1 align="center">📊 SheetGenie</h1>
<p align="center"><em>Speak or type what you need — get a real Excel spreadsheet, instantly.</em></p>

SheetGenie turns a plain-language request into a downloadable `.xlsx` workbook with
clearly labelled columns, live formulas, totals rows, conditional formatting, dropdown
validation, named ranges, and native Excel charts — empty as a template, or filled with
data you type, paste, **photograph, or drop in as a PDF**. Past spreadsheets are saved in
a private in-browser library you can reopen, edit by prompt, rename, duplicate, and export.
It installs as a web app on **Android, iPhone, PC, and Mac**.

> "Make me a monthly budget tracker with categories, budgeted vs actual columns,
> a variance formula, and a bar chart." → a finished spreadsheet, one click later.

## How it works
1. **Capture** — type or dictate your request, and optionally add data by pasting,
   uploading a file, snapping a **photo** (camera), or attaching a **PDF** (the AI reads it).
2. **Improve** — an AI step rewrites it into a precise blueprint you can review. If your
   request is ambiguous, it first asks a couple of quick questions, then builds. You can
   also **edit an existing spreadsheet** by prompt and/or new data.
3. **Generate** — the blueprint is rendered into a real `.xlsx` and downloaded.

The expensive AI step runs once; the file build is free, instant, and reproducible.

## Tested
A repeatable stress harness ([tests/stress_suite.py](tests/stress_suite.py)) runs prompts
through the live pipeline and grades the `.xlsx` against an external QA checklist — the
current build scores **14/14** across totals rows, multi-sheet, currency, conditional
formatting, dropdown validation, charts, IF/lookup/cross-sheet formulas, named ranges,
the clarifying-question behavior, and data fidelity (verbatim, computed columns, dedup,
aggregation). Plus **77 offline unit tests** and several adversarial review passes.

## Tech at a glance
- **Frontend:** vanilla-JS PWA (installable, offline shell, voice + typed input)
- **AI:** Google Gemini (free tier) primary → xAI Grok fallback — multimodal (image + PDF) with live web-search, producing a structured spec
- **Backend:** Vercel Python serverless functions (key stays server-side)
- **Excel:** openpyxl (charts, formulas, totals rows, conditional formatting, dropdowns, named ranges)
- **Hosting:** Vercel — online 24/7, auto-deploy from GitHub

## Documentation
| Doc | What's in it |
|-----|--------------|
| [docs/SPEC.md](docs/SPEC.md) | **The shared contract** — HTTP API + `SpreadsheetSpec` (read first) |
| [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) | System design, stack rationale, security |
| [docs/ORCHESTRATION.md](docs/ORCHESTRATION.md) | How an AI agent team built it (methodology + research) |
| [docs/DEPLOY.md](docs/DEPLOY.md) | Turnkey deploy + local dev (non-developer friendly) |

## Quick start
See [docs/DEPLOY.md](docs/DEPLOY.md). Short version: add `GEMINI_API_KEY` (free) to a
local `.env`, run `vercel dev`, open `http://localhost:3000`.

## Cost
**Free.** The primary provider is Google Gemini's free tier (rate-limited, no card),
so the AI step costs nothing for normal/portfolio traffic. The optional xAI **Grok**
fallback uses a little credit only if Gemini is ever unavailable. Live web-search uses
Gemini's free Google-Search grounding.

---
<p align="center">Built as a portfolio project. The Excel generator runs no AI and makes no network calls — it cannot leak anything.</p>
