# AI Intern Review Bot

Automates the full intern review cycle for 100 interns using an Excel sheet, Claude, Gmail and Google Meet.

**Flow**
1. `interns.xlsx` holds intern details (ID, name, email, task) and the task list with reference solutions.
2. `kickoff`: Claude explains each task like a mentor on a call; the bot emails it with the review date, time, Google Meet link and meeting ID.
3. Interns submit work into `submissions/<ID>.py` (e.g. `I001.py`).
4. `review` (every 2 days): Claude compares each submission with the reference solution, writes human-like feedback, score and improvement areas, emails them, and books the next session (date, time, Meet link, meeting ID). Everything is written back to Excel.

## Setup
```bash
pip install -r requirements.txt
cp .env.example .env        # add ANTHROPIC_API_KEY, SMTP_USER, SMTP_PASS, DEMO_EMAIL
python make_sample_data.py  # creates interns.xlsx (100 sample interns) + 3 demo submissions
```
Edit `interns.xlsx` (columns A-D) with your real interns.

## Run
```bash
python intern_bot.py kickoff --limit 3        # dry-run: emails saved in ./outbox, nothing sent
python intern_bot.py review                   # dry-run review of submitted files
python intern_bot.py kickoff --live           # really send mail + create Meet events
python intern_bot.py review --live --only I001
```
Options: `--limit N`, `--only I001`, `--live`.

## Real Google Meet links
Create OAuth credentials (Desktop app) in Google Cloud Console, enable the Google Calendar API, save the file as `credentials.json`. The first `--live` run opens a browser to authorise. Without it, the bot uses clearly fake demo codes.

## Notes
- Meetings are grouped 10 interns per 30-minute slot starting 10:00 (change via `.env`).
- Without `ANTHROPIC_API_KEY` a crude offline fallback runs, so always set the key for real feedback.
- Set `DEMO_EMAIL` while testing so the 100 interns don't get mail.
