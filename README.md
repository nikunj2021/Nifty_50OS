# Nifty 500 Weekly/Monthly RSI Screener

Scans every symbol in `nifty500.csv`, computes Weekly RSI and Monthly RSI
(Wilder's method), and flags stocks where **Weekly RSI < 30** and
**Monthly RSI > 40**. Runs automatically every **Saturday 9:00 AM IST** via
GitHub Actions and publishes an Excel + HTML report to GitHub Pages, with
Telegram and webhook alerts.

## 1. Replace the CSV

Overwrite `nifty500.csv` with the full Nifty 500 list. Keep the header row
exactly as `Symbol`, one NSE ticker per row (no `.NS` suffix — the script
adds it), e.g.:

```
Symbol
RELIANCE
TCS
HDFCBANK
```

## 2. Create the repository and push this code

```bash
cd nifty500-rsi-screener
git init
git add .
git commit -m "Initial commit: Nifty 500 RSI screener"
gh repo create nifty500-rsi-screener --public --source=. --remote=origin --push
```

(No `gh` CLI? Create an empty repo at github.com/new, then
`git remote add origin <your-repo-url>` and `git push -u origin main`.)

## 3. Enable GitHub Pages

Repo → **Settings → Pages** → Source: `Deploy from a branch` → Branch:
`main`, folder `/docs` → Save. Your report will then be live at:

```
https://<your-username>.github.io/nifty500-rsi-screener/
```

## 4. Add secrets

Repo → **Settings → Secrets and variables → Actions**:

| Name | Required | Notes |
|---|---|---|
| `TELEGRAM_BOT_TOKEN` | optional | from @BotFather |
| `TELEGRAM_CHAT_ID` | optional | your chat/group id |
| `WEBHOOK_URL` | optional | any endpoint that accepts a JSON POST |

Also add a repo **variable** (Settings → Secrets and variables → Actions →
Variables tab) called `PAGES_BASE_URL` set to your Pages URL above, so the
HTML report's download link and Telegram message point to the live report.

## 5. Test it

Repo → **Actions** tab → "Nifty 500 RSI Screener" → **Run workflow** to
trigger it manually before waiting for Saturday.

## How it works

- `screener.py` pulls full daily history per symbol via `yfinance`,
  resamples to weekly (`W-FRI`) and monthly closes, and computes RSI(14)
  with Wilder's smoothing on each.
- The 52-week high and current price come from the same daily history
  (trailing 365 calendar days), so "% Down from 52W High" is
  `(current − 52W High) / 52W High × 100`.
- Output: `docs/nifty500_rsi_report.xlsx` and `docs/index.html`, committed
  back to the repo each run so GitHub Pages always serves the latest report.
- The GitHub Actions cron `30 3 * * 6` is 03:30 UTC Saturday = 09:00 IST.
