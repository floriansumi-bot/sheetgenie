# SheetGenie — Deploy & Local Dev (turnkey)

Plain-language steps. You do not need to be a developer — follow in order.

---

## Part A — Get your Anthropic API key (one time, ~3 min)
1. Go to **https://console.anthropic.com** and sign in (create an account if needed).
2. Add a payment method under **Billing**. With the default **Fable 5** model expect
   roughly a few cents up to ~40¢ per spreadsheet (it's the most capable model). To
   spend less, set `MODEL=claude-opus-4-8` or lower `EFFORT` in Part C. Set a low
   monthly spend limit for peace of mind. Your $5 is plenty to demo it.
3. Open **API Keys → Create Key**, name it `sheetgenie`, and **copy the key**
   (starts with `sk-ant-`). You will paste it into Vercel in Part C. Keep it secret.

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
   - `ANTHROPIC_API_KEY` = the `sk-ant-...` key from Part A
   - `MODEL` = `claude-opus-4-8`  *(recommended for a new key — Fable 5 access is
     often not enabled on a fresh account, and the app would waste a round-trip
     failing over to Opus on every request. Set `claude-fable-5` instead once your
     account has Fable; it then auto-falls back to Opus 4.8 → Sonnet 4.6 anyway.)*
   - `EFFORT` = `high`  *(optional; `low`/`medium`/`high`/`max` — lower = cheaper & faster)*
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
(`anthropic==0.109.2`, `openpyxl==3.1.5`), so Vercel installs the same ones that were
tested. No action needed unless you intentionally upgrade.

---

## Troubleshooting
| Symptom | Fix |
|---------|-----|
| "Server not configured" / 500 on Improve | `ANTHROPIC_API_KEY` missing or wrong in Vercel env vars → re-add, redeploy |
| Voice button does nothing on iPhone Chrome | iOS Chrome can't do speech (Apple limitation) — use Safari, or just type |
| Build fails on Vercel about Python deps | Ensure `requirements.txt` is at repo root; redeploy |
| Charts missing in the file | Check the spec's `valueColumns`/`categoriesColumn` indices are valid (see SPEC.md) |
