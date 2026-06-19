# SheetGenie — Raspberry Pi "last resort" worker

This is the **ultimate fallback**. When Gemini *and* Groq (and any other cloud
provider) are all busy, the website offers: *"our fast AI is busy — enter your
email and the backup server will build it and send it to you."* That backup server
is your Pi: it runs a local LLM with **no time limit**, builds the real `.xlsx`, and
emails it.

The cloud never waits for the Pi — it hands off the job in a split second (HTTP 202)
and the Pi does the slow part on its own. So Vercel's 60-second limit never applies.

```
website  →  /api/queue (Vercel, instant hand-off)  →  your Pi (tunnel)
                                                         → local LLM builds spec
                                                         → openpyxl renders .xlsx
                                                         → emails it to the user
```

You only need to set this up **once**. ~30 minutes. Everything below runs on the Pi
unless it says "on Vercel".

---

## 1. Install Ollama + a model (the local AI)
```bash
curl -fsSL https://ollama.com/install.sh | sh
ollama pull qwen2.5:3b        # ~2 GB; good quality-for-size on a Pi 5
```
Test it: `ollama run qwen2.5:3b "say hi"` then Ctrl-D. (Bigger = better but slower:
`qwen2.5:7b` works on the 16 GB Pi at ~2 tok/s if you don't mind the wait.)

## 2. Get the code + Python deps
```bash
git clone https://github.com/floriansumi-bot/sheetgenie.git
cd sheetgenie
python3 -m venv .venv && . .venv/bin/activate
pip install -r pi/requirements.txt
```

## 3. Set up email sending (Gmail app password)
The worker emails the finished file via SMTP. Easiest is your own Gmail:
1. Turn on **2-Step Verification** at https://myaccount.google.com/security
2. Create an **App password**: https://myaccount.google.com/apppasswords → "Mail" →
   copy the 16-character password.
3. You'll use your Gmail address as `SMTP_USER` and that app password as `SMTP_PASS`.

(Any SMTP works — set `SMTP_HOST`/`SMTP_PORT`/`SMTP_USER`/`SMTP_PASS` accordingly.)

## 4. Pick a strong shared secret
This stops strangers from using your Pi. Generate one and keep it:
```bash
python3 -c "import secrets; print(secrets.token_urlsafe(32))"
```

## 5. Run the worker (test it locally)
```bash
WORKER_SECRET="<the secret from step 4>" \
SMTP_USER="florian.sumi@gmail.com" \
SMTP_PASS="<your 16-char app password>" \
EMAIL_FROM="florian.sumi@gmail.com" \
OLLAMA_MODEL="qwen2.5:3b" \
python pi/worker.py
```
In another terminal: `curl http://localhost:8080/health` → `{"ok": true, ...}`.

## 6. Expose it to the internet with a Cloudflare Tunnel (free, no port-forwarding)
```bash
# install cloudflared (Pi / arm64):
curl -L https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-linux-arm64 -o cloudflared
sudo install cloudflared /usr/local/bin/cloudflared

# quick tunnel (gives a random https URL, good for testing):
cloudflared tunnel --url http://localhost:8080
```
It prints a URL like `https://something-random.trycloudflare.com`. That's your
**PI_WORKER_URL**. (For a permanent URL, set up a *named* tunnel on your own domain —
see Cloudflare's docs — otherwise the quick-tunnel URL changes each restart.)

## 7. Point the website at your Pi (on Vercel)
In **Vercel → sheetgenie → Settings → Environment Variables**, add:
- `PI_WORKER_URL` = your tunnel URL from step 6 (e.g. `https://xxxx.trycloudflare.com`)
- `PI_WORKER_SECRET` = the secret from step 4 (must match exactly)

Then redeploy (or just push any change). From now on, when all cloud AIs are busy,
the site shows the "email it to me" option and your Pi fulfils it.

## 8. Make it permanent (survive reboots) — systemd
Create `/etc/systemd/system/sheetgenie-worker.service` (edit paths/secrets):
```ini
[Unit]
Description=SheetGenie Pi worker
After=network-online.target

[Service]
WorkingDirectory=/home/pi/sheetgenie
ExecStart=/home/pi/sheetgenie/.venv/bin/python /home/pi/sheetgenie/pi/worker.py
Environment=WORKER_SECRET=REPLACE_ME
Environment=SMTP_USER=florian.sumi@gmail.com
Environment=SMTP_PASS=REPLACE_ME
Environment=EMAIL_FROM=florian.sumi@gmail.com
Environment=OLLAMA_MODEL=qwen2.5:3b
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
```
```bash
sudo systemctl daemon-reload
sudo systemctl enable --now sheetgenie-worker
```
Do the same for `cloudflared` (a named tunnel can run as a service too), so both
start on boot.

---

## Configuration reference (environment variables)
| Var | Default | Meaning |
|-----|---------|---------|
| `PORT` | `8080` | Port the worker listens on |
| `WORKER_SECRET` | *(empty)* | Shared secret; **set this** — must match Vercel's `PI_WORKER_SECRET` |
| `OLLAMA_URL` | `http://localhost:11434` | Ollama endpoint |
| `OLLAMA_MODEL` | `qwen2.5:3b` | Local model to use |
| `MAX_TOKENS` | `4096` | Max tokens the model may generate |
| `OLLAMA_TIMEOUT` | `900` | Seconds to allow for generation (Pi is slow) |
| `SMTP_HOST` / `SMTP_PORT` | `smtp.gmail.com` / `587` | Mail server (STARTTLS) |
| `SMTP_USER` / `SMTP_PASS` | *(empty)* | Mailbox + app password |
| `EMAIL_FROM` / `EMAIL_FROM_NAME` | `SMTP_USER` / `SheetGenie` | From address / display name |

## Notes
- **Quality:** a 3B local model is weaker than the cloud — fewer formulas, simpler
  data. That's fine: it only runs when everything else is down, and a decent sheet a
  few minutes later beats nothing.
- **Attachments** (photos/PDFs) aren't read on this path — the backup is text-only.
- **Security:** keep `WORKER_SECRET` private; the tunnel gives HTTPS; only
  `/generate-async` requires the secret (`/health` is open for tunnel checks).
- **Reliability:** the Pi is only as up as your home power + internet. With Gemini +
  Groq in front of it, it should rarely be needed.
