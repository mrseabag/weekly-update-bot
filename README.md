# Weekly Update Bot

A Telegram bot that automates weekly status reporting for a small team. It collects structured updates from each team member on a schedule and delivers a consolidated report by email, removing the need to manually chase people down and compile their responses by hand.

## Overview

Each Monday morning, the bot messages every registered team member with two prompts:

1. What did you complete this week
2. What are you currently working on

Team members can optionally attach supporting files (documents, screenshots, reports). Every response is logged automatically to a formatted Excel spreadsheet, and every Friday afternoon that spreadsheet is emailed to the manager — no manual copying, formatting, or follow-up required.

## Features

- Scheduled prompts and reports (Monday check-in, Friday summary email) — fully configurable
- Two-question structured format, with optional file attachments (documents, photos, video, audio)
- Automatic Excel logging with per-week tracking
- On-demand updates at any time via `/update`, not just the scheduled prompt
- Admin controls to trigger reports or reminders manually
- No third-party data sharing — all data stays within the bot's own spreadsheet and the manager's inbox

## Commands

| Command | Access | Description |
|---|---|---|
| `/start` | Everyone | Registers the user with the bot |
| `/update` | Everyone | Submits a weekly update on demand |
| `/status` | Everyone | Shows how many users and updates are on record |
| `/cancel` | Everyone | Cancels an update in progress |
| `/help` | Everyone | Lists available commands |
| `/sendreport` | Admin | Sends the Excel report by email immediately |
| `/promptall` | Admin | Sends the update reminder to all users immediately |

## Setup

Full installation, configuration, and hosting instructions are in [`SETUP.md`](./SETUP.md), covering:

- Installing dependencies
- Creating and configuring the Telegram bot token
- Setting up the Gmail App Password used to send reports
- Deployment options (this project currently runs on Railway via the included `Dockerfile`)

Configuration is handled entirely through environment variables (see `.env.example`) — no credentials are stored in source code.

## Data & Security

- Credentials (bot token, email account, app password) are read from environment variables at runtime and are never committed to source control
- Communication with Telegram uses their standard encrypted Bot API
- Reports are sent over an authenticated, encrypted SMTP connection (SSL)
- Collected data is limited to: name, Telegram username, submitted text responses, and any files explicitly attached by the user

## Tech Stack

- **Python** — [`python-telegram-bot`](https://github.com/python-telegram-bot/python-telegram-bot) for the Telegram integration and scheduling
- **openpyxl** for spreadsheet generation and formatting
- **smtplib** for report delivery
- **Docker** for deployment

## Project Structure

```
bot.py              Main application
requirements.txt     Python dependencies
Dockerfile           Container build definition
SETUP.md             Setup and hosting guide
.env.example         Environment variable template
data/                Runtime data (Excel report, registered users) — not tracked in git
```
