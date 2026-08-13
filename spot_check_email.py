"""
spot_check_email.py — Daily quality spot-check email.

Picks 10-15 profiles updated today at random and emails them to
support@aeroprofessional.com so the team can spot-check them in Tracker.

Run daily at 5pm via GitHub Actions (see aerotracker.yml).
"""
import json, os, random, datetime, smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

DAILY_LOG_FILE  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "daily_updates.json")
EMAIL_FROM      = os.environ.get("EMAIL_LOGIN_USER", "emily.walton@aeroprofessional.com")
EMAIL_PASSWORD  = os.environ.get("EMAIL_PASSWORD",  "PurpleAutumn96?")
SMTP_SERVER     = "smtp.office365.com"
SMTP_PORT       = 587
SEND_TO         = "support@aeroprofessional.com"
SAMPLE_SIZE     = 15

def send_spot_check():
    today = datetime.date.today().isoformat()

    try:
        with open(DAILY_LOG_FILE, "r") as f:
            log = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        print("No daily log found — nothing to spot-check today.")
        return

    today_updates = log.get(today, [])
    if not today_updates:
        print(f"No profiles updated today ({today}) — skipping spot-check email.")
        return

    sample = random.sample(today_updates, min(SAMPLE_SIZE, len(today_updates)))
    total  = len(today_updates)

    lines = [
        f"Daily Spot-Check — {today}",
        f"{'=' * 50}",
        f"Total profiles updated today: {total}",
        f"Showing {len(sample)} randomly selected for review.",
        f"Please check these in Tracker and reply if anything looks wrong.",
        "",
    ]

    for i, c in enumerate(sample, 1):
        lines += [
            f"[{i}] {c['name']}  (ID: {c['id']})",
            f"     Job title  : {c['job_title']}",
            f"     Employer   : {c['employer']}",
            f"     Work type  : {c['work_type']}",
            f"     Skills     : {', '.join(c['skills']) if c['skills'] else '(none)'}",
            "",
        ]

    body = "\n".join(lines)
    subject = f"Tracker Spot-Check — {len(sample)} profiles to review ({today})"

    # ── Method 1: Outlook COM (Windows) ──────────────────────────────────────
    try:
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To      = SEND_TO
        mail.Subject = subject
        mail.Body    = body
        mail.Send()
        print(f"  ✉  Spot-check email sent via Outlook ({len(sample)} profiles)")
        return
    except Exception:
        pass

    # ── Method 2: SMTP (GitHub Actions / Linux) ───────────────────────────────
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = EMAIL_FROM
        msg["To"]      = SEND_TO
        msg.attach(MIMEText(body, "plain"))
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.ehlo()
            server.starttls()
            server.login(EMAIL_FROM, EMAIL_PASSWORD)
            server.sendmail(EMAIL_FROM, SEND_TO, msg.as_string())
        print(f"  ✉  Spot-check email sent via SMTP ({len(sample)} profiles)")
    except Exception as e:
        print(f"  ⚠  Could not send spot-check email: {e}")

if __name__ == "__main__":
    send_spot_check()
