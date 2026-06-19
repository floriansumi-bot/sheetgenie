# SheetGenie — Deploy & Local Dev (turnkey)

Plain-language steps. You do not need to be a developer — follow in order.

---

## Part A — Get your API keys (one time, ~3 min)
1. **Gemini (free, required)** — go to **https://aistudio.google.com**, click
   **"Get API key"** (no credit card). Copy the key. This is the primary provider.
2. **Groq (free, recommended)** — go to **https://console.groq.com** → **API Keys** →
   create one (starts with `gsk_`, no credit card). This is the **free fallback**: when
   Gemini's daily free quota is used up, the app automatically uses Groq instead, so it
   never shows "AI temporarily unavailable". Two free providers = effectively always-on.
3. **Grok / xAI (optional, PAID)** — only if you want a third fallback. ⚠️ xAI has **no
   free tier** (returns `403` without credit). **Skip it** — Gemini + Groq are both free.
4. Keep your key(s) secret — you'll paste them into Vercel in Part C.

---

## Part B — Put the code on GitHub (one time, ~5 min)
From this folder (`sheet-genie`):
```bash
git init
git add .
git commit -m "SheetGenie: initial app"
```
Then create an empty repo on github.com (e.g. `sheetgenie`) and push:
```bash
git remote add origin https://github.com/<your-username>/sheetgenie.git
git branch -M main
git push -u origin main
```

---

## Part C — Deploy on Vercel (one time, ~5 min) → online 24/7
1. Go to **https://vercel.com**, sign in **with GitHub**.
2. **Add New → Project**, import the `sheetgenie` repo.
3. Framework Preset: **Other** (it's a static site + Python functions — no build needed).
4. Expand **Environment Variables** and add:
   - `GEMINI_API_KEY` = your free Gemini key from Part A  ← **required (free)**
   - `GROQ_API_KEY` = your `gsk_...` key  ← **recommended free fallback — keeps the app up**
   - `XAI_API_KEY` = your `xai-...` key  *(optional — paid; leave out)*
   - `PROVIDERS` = `gemini,groq`  *(optional; this is the default. Add `,grok` only if you funded xAI.)*
   - `WEB_SEARCH` = `on`  *(optional; live data via Gemini grounding — set `off` to disable)*
5. Click **Deploy**. After ~1 min you get a live URL like `https://sheetgenie.vercel.app`.
6. (Optional) **Settings → Domains** to attach a custom domain for your portfolio.

Every future `git push` to `main` auto-redeploys. That's it — it stays online.

### Install it as an app
Open the live URL on your phone/computer → browser menu → **Install / Add to Home Screen**.

---

## Part D — Run it locally first (optional, recommended)
Test before you deploy.

**Option 1 — Vercel CLI (closest to production, runs both functions + site):**
```bash
npm i -g vercel
vercel dev          # serves the site + /api/* on http://localhost:3000
```
Create a local `.env` (copy from `.env.example`) with your key first.

**Option 2 — verify the generator with no key (free, offline):**
The Excel generator needs no API key. See `docs/DEPLOY.md` tests in the repo or run:
```bash
python -m pip install -r requirements.txt
```
then exercise `api/generate.py` with a sample spec (the QA step does this automatically).

---

## Dependencies (already pinned)
`requirements.txt` is pinned to the exact versions verified during QA
(`google-genai==2.8.0`, `openai==2.42.0`, `openpyxl==3.1.5`), so Vercel installs the
same ones that were tested. No action needed unless you intentionally upgrade.

---

## Troubleshooting
| Symptom | Fix |
|---------|-----|
| "Server not configured" / 500 on Improve | `GEMINI_API_KEY` (and/or `XAI_API_KEY`) missing or wrong in Vercel env vars → re-add, redeploy |
| "AI is temporarily unavailable" | Gemini free-tier rate limit hit, or the key is invalid → wait a bit, or check the key / add an `XAI_API_KEY` fallback |
| Voice button does nothing on iPhone Chrome | iOS Chrome can't do speech (Apple limitation) — use Safari, or just type |
| Build fails on Vercel about Python deps | Ensure `requirements.txt` is at repo root; redeploy |
| Charts missing in the file | Check the spec's `valueColumns`/`categoriesColumn` indices are valid (see SPEC.md) |
