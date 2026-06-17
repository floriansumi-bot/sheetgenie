# SheetGenie — Architecture

## What it is
A prompt-to-spreadsheet web app. You **type or speak** what you want; an AI step
rewrites your prompt into a precise blueprint; a generator turns that blueprint into
a real `.xlsx` (labelled columns, formulas, native Excel charts) — empty/templated,
or filled with data you supply — and hands it back as an instant download.

Works on Android, iPhone, PC, and Mac from one codebase, installable as a PWA.

## The flow
```
                 ┌─────────────────────────── Browser (PWA) ───────────────────────────┐
   speak ──▶ 🎤  │  Web Speech API ─┐                                                   │
   type  ──▶ ⌨   │                  ├─▶ prompt (+ optional pasted data)                 │
                 │                  │                                                   │
                 │        [Improve] ─┼──── POST /api/improve ─────────────┐             │
                 │                  │                                     ▼             │
                 │   shows improved prompt + editable preview      ┌─────────────┐      │
                 │                  │                              │ /api/improve │──▶ Gemini → Grok
                 │        [Generate] ┼──── POST /api/generate ◀── spec ┘ (Gemini)     │
                 │                  ▼                              returns improvedPrompt│
                 │            downloads .xlsx ◀── binary ── ┌──────────────┐  + spec     │
                 │                                          │ /api/generate │ (openpyxl) │
                 └──────────────────────────────────────────┴──────────────┴────────────┘
```

1. **Capture** — user types or dictates a prompt; may paste tabular data.
2. **Improve** (`/api/improve`) — the AI provider (Gemini, then Grok) returns one of two things: a few
   short **clarifying questions** (when the request is genuinely ambiguous), or, when it
   has enough to go on, `{ status: "ready", improvedPrompt, notes, spec }`. If it asks,
   the user answers inline and the app re-calls `/api/improve` with those answers folded
   in; the model then builds (it won't ask twice). The `spec` is a
   [`SpreadsheetSpec`](SPEC.md), shown as a preview before download.
3. **Generate** (`/api/generate`) — openpyxl renders the spec to `.xlsx`, deterministic
   and free (no AI, no network). Streams back as a download.

Splitting "improve" (the AI step) from "generate" (free, deterministic) means the AI
step runs once and the file build is instant and reproducible.

## Stack & why
| Layer    | Choice | Why |
|----------|--------|-----|
| Frontend | Vanilla JS + CSS, PWA | No build step, installable everywhere, easy for the owner to maintain, fast |
| Voice    | Web Speech API (`webkitSpeechRecognition` fallback) | Native, free; works Chrome/Android + iOS Safari 14.5+; typed input always present |
| AI       | Google Gemini (free) → xAI Grok fallback (env `PROVIDERS`) | Free primary (multimodal: image + PDF; Google-Search grounding), paid fallback for resilience; JSON envelope parsed defensively |
| Backend  | Vercel Python serverless (`BaseHTTPRequestHandler`) | Always-on, free Hobby tier, real Excel charts via openpyxl, key stays server-side |
| Excel    | openpyxl | Native `.xlsx` with bar/line/pie charts, formulas, formatting |
| Hosting  | Vercel | 24/7, auto-deploy from GitHub, custom domain, env-var secret storage |

## Security posture
- `GEMINI_API_KEY` / `XAI_API_KEY` live only in Vercel env vars / local `.env` — never shipped to the browser, never committed (`.gitignore`).
- `/api/generate` runs no AI and makes no outbound calls — it cannot leak anything.
- Error messages are sanitized; raw stack traces and the key are never returned to the client.
- Input limits in [SPEC.md](SPEC.md) bound spec size to prevent resource abuse.

## Repository layout
```
sheet-genie/
├── api/
│   ├── improve.py        # POST /api/improve  — prompt → {improvedPrompt, notes, spec}
│   └── generate.py       # POST /api/generate — spec → .xlsx bytes
├── public/
│   ├── index.html        # app shell
│   ├── styles.css        # responsive, light/dark, mobile-first
│   ├── app.js            # capture → improve → preview → generate → download
│   ├── manifest.webmanifest
│   ├── sw.js             # service worker (offline shell, installable)
│   └── icons/            # PWA icons
├── docs/
│   ├── SPEC.md           # ← the shared contract (read first)
│   ├── ARCHITECTURE.md   # this file
│   ├── ORCHESTRATION.md  # how the agent team built it
│   └── DEPLOY.md         # turnkey deploy + local dev
├── requirements.txt
├── vercel.json
├── .env.example
└── README.md
```
