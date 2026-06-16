<h1 align="center">📊 SheetGenie</h1>
<p align="center"><em>Speak or type what you need — get a real Excel spreadsheet, instantly.</em></p>

SheetGenie turns a plain-language request into a downloadable `.xlsx` workbook with
clearly labelled columns, formulas, and native Excel charts — empty as a template, or
filled with data you provide. It runs as an installable web app on **Android, iPhone,
PC, and Mac**.

> "Make me a monthly budget tracker with categories, budgeted vs actual columns,
> a variance formula, and a bar chart." → a finished spreadsheet, one click later.

## How it works
1. **Capture** — type or dictate your request (and optionally paste your data).
2. **Improve** — an AI step rewrites it into a precise blueprint you can review.
3. **Generate** — the blueprint is rendered into a real `.xlsx` and downloaded.

The expensive AI step runs once; the file build is free, instant, and reproducible.

## Tech at a glance
- **Frontend:** vanilla-JS PWA (installable, offline shell, voice + typed input)
- **AI:** Anthropic API (`claude-fable-5`, with an Opus 4.8 fallback; swappable) producing a structured spec
- **Backend:** Vercel Python serverless functions (key stays server-side)
- **Excel:** openpyxl (charts, formulas, formatting)
- **Hosting:** Vercel — online 24/7, auto-deploy from GitHub

## Documentation
| Doc | What's in it |
|-----|--------------|
| [docs/SPEC.md](docs/SPEC.md) | **The shared contract** — HTTP API + `SpreadsheetSpec` (read first) |
| [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) | System design, stack rationale, security |
| [docs/ORCHESTRATION.md](docs/ORCHESTRATION.md) | How an AI agent team built it (methodology + research) |
| [docs/DEPLOY.md](docs/DEPLOY.md) | Turnkey deploy + local dev (non-developer friendly) |

## Quick start
See [docs/DEPLOY.md](docs/DEPLOY.md). Short version: add `ANTHROPIC_API_KEY` to a local
`.env`, run `vercel dev`, open `http://localhost:3000`.

## Cost
With the default **Fable 5** model (the most capable), expect roughly a few cents up to
~40¢ per spreadsheet. To spend less, set `MODEL=claude-opus-4-8` or lower `EFFORT` — the
app still produces great results. Set a spend limit in the Anthropic console for peace of mind.

---
<p align="center">Built as a portfolio project. The Excel generator runs no AI and makes no network calls — it cannot leak anything.</p>
