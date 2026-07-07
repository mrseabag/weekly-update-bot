# Weekly Update Bot — Setup Guide

## What this bot does
- Every **Monday 9am UTC** it messages all registered team members asking for their weekly update
- Collects: ✅ what they completed, 🔨 what they're working on, 📎 any file attachments
- Saves everything to `data/weekly_updates.xlsx`
- Every **Friday 5pm UTC** it emails the Excel file to the manager

---

## Step 1 — Install Python dependencies

```
cd C:\Users\seble\weekly-update-bot
pip install -r requirements.txt
```

---

## Step 2 — Configure email (so the manager gets the weekly report)

1. Log into the Gmail account you want the bot to send FROM
2. Go to: **Google Account → Security → 2-Step Verification → App passwords**
3. Create an App Password for "Mail"
4. Copy the 16-character password

Copy `.env.example` to `.env` and fill in the values there (never edit `bot.py` directly with real credentials):

```
MANAGER_EMAIL=boss@yourcompany.com     # where the weekly report goes
SENDER_EMAIL=yourbot@gmail.com          # the Gmail sending it
SENDER_PASSWORD=xxxx xxxx xxxx xxxx    # the App Password from above
```

---

## Step 3 — Get a new bot token (IMPORTANT — the old one is compromised)

Since the token was shared in a chat, regenerate it:

1. Open Telegram → search **@BotFather**
2. Send `/mybots` → select your bot → **API Token → Revoke current token**
3. Copy the new token into your `.env` file (or the host's environment variables — see Step 5):
   ```
   BOT_TOKEN=123456789:xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
   ```

---

## Step 4 — Run the bot locally (for testing)

```
cd C:\Users\seble\weekly-update-bot
python bot.py
```

Tell each team member to open Telegram, search for your bot by name, and send `/start`.

---

## Step 5 — Host it so it runs 24/7 (FREE options)

### Option A: PythonAnywhere (easiest — recommended)

1. Go to [pythonanywhere.com](https://www.pythonanywhere.com) and create a free account
2. Go to **Files** → upload `bot.py` and `requirements.txt`
3. Open a **Bash console** and run:
   ```bash
   pip install -r requirements.txt
   python bot.py &
   ```
4. To keep it running: go to **Tasks** → set up an "Always-on task"
   - Free accounts get 1 always-on task
   - Command: `python /home/yourusername/bot.py`

### Option B: Railway.app

1. Go to [railway.app](https://railway.app) — free tier available
2. Connect your GitHub repo (or upload files via their CLI)
3. Set environment variables in Railway dashboard:
   - `BOT_TOKEN`
   - `MANAGER_EMAIL`
   - `SENDER_EMAIL`
   - `SENDER_PASSWORD`
4. Deploy — Railway keeps it running automatically

### Option C: Run on your own PC (24/7 only if PC stays on)

Use Windows Task Scheduler to start the bot on login:
1. Open Task Scheduler → Create Basic Task
2. Trigger: At log on
3. Action: Start a program → `python` with argument `C:\Users\seble\weekly-update-bot\bot.py`

---

## Using the bot

| Command | Who | What it does |
|---------|-----|-------------|
| `/start` | Team members | Register with the bot |
| `/update` | Team members | Submit a weekly update now |
| `/cancel` | Team members | Cancel an in-progress update |
| `/status` | Anyone | See how many updates are saved |
| `/sendreport` | Admin | Email the Excel report right now |
| `/promptall` | Admin | Send update reminder to all users |
| `/help` | Anyone | Show command list |

---

## Excel report format

The file `data/weekly_updates.xlsx` has these columns:

| Week | Date Submitted | Name | Username | What did you complete? | What are you working on? | Files Submitted |
|------|---------------|------|----------|----------------------|--------------------------|-----------------|

A new row is added for every update submitted. The file is emailed every Friday.

---

## Adjusting the schedule

In `bot.py`, near the top:

```python
PROMPT_DAY  = 0     # 0=Monday, 4=Friday, 6=Sunday, etc.
PROMPT_HOUR = 9     # 9 AM UTC

REPORT_DAY  = 4     # Friday
REPORT_HOUR = 17    # 5 PM UTC
```

UTC+0 is used. If your team is in UTC+2, set PROMPT_HOUR=7 to deliver at 9am local.
