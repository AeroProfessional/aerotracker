"""
no_cv_report.py — Daily report of candidates with no CV uploaded.

Scheduled to run at 3pm every day. Sends an email to the Aero Professional
team listing candidates who registered but haven't uploaded a CV. The team
will manually email those candidates and delete their profiles if no CV
arrives within 3 days.

This script NEVER emails candidates directly.

Run:  py no_cv_report.py
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import json, datetime, smtplib, requests
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

from update_tracker import get_jwt, h, TRACKER_API, PENDING_CV_FILE

# ── Who receives the daily report ─────────────────────────────────────────────
REPORT_TO         = "support@aeroprofessional.com"
EMAIL_FROM        = "support@aeroprofessional.com"
EMAIL_PASSWORD    = os.environ.get("EMAIL_PASSWORD", "")
EMAIL_LOGIN       = "emily.walton@aeroprofessional.com"
SMTP_SERVER       = "smtp.office365.com"
SMTP_PORT         = 587
GMAIL_USER        = os.environ.get("GMAIL_USER", "")
GMAIL_APP_PASSWORD= os.environ.get("GMAIL_APP_PASSWORD", "")

# ── How many days before flagging as overdue ───────────────────────────────────
OVERDUE_DAYS = 3


def check_has_cv(jwt, resource_id):
    """Return True if the candidate now has at least one document in Tracker."""
    try:
        r = requests.get(
            f"{TRACKER_API}/api/v1/Resource/{resource_id}/Documents",
            headers=h(jwt), timeout=10
        )
        if r.status_code == 200:
            data = r.json()
            if isinstance(data, list):
                return len(data) > 0
            if isinstance(data, dict):
                return bool(data.get("items") or data.get("documents") or
                            data.get("value") or data.get("results"))
    except Exception:
        pass
    return False


def send_report_email(still_waiting, newly_resolved, overdue):
    """Send the daily no-CV summary to the team inbox."""
    today_str = datetime.date.today().strftime("%d %B %Y")
    subject = f"[AeroTracker] No-CV Report — {today_str}"

    lines = []
    lines.append(f"Daily No-CV Report  |  {today_str}")
    lines.append("=" * 60)

    if overdue:
        lines.append(f"\n⚠  OVERDUE ({OVERDUE_DAYS}+ days — action required):\n")
        for e in overdue:
            days = e["days"]
            lines.append(f"  • {e['name']}")
            lines.append(f"    Tracker ID : {e['resourceId']}")
            if e.get("email"):
                lines.append(f"    Email      : {e['email']}")
            lines.append(f"    Waiting    : {days} day{'s' if days != 1 else ''}")
            lines.append("")

    if still_waiting:
        lines.append(f"\n⏳  Awaiting CV ({len(still_waiting)} candidate{'s' if len(still_waiting) != 1 else ''}):\n")
        for e in still_waiting:
            days = e["days"]
            lines.append(f"  • {e['name']}")
            lines.append(f"    Tracker ID : {e['resourceId']}")
            if e.get("email"):
                lines.append(f"    Email      : {e['email']}")
            lines.append(f"    Waiting    : {days} day{'s' if days != 1 else ''}")
            lines.append("")

    lines.append("\n" + "=" * 60)
    lines.append("To action: manually email the candidates above to request their CV.")
    lines.append("If no CV received within 3 days, delete their Tracker profile manually.")
    lines.append("Do NOT reply to this email — it is sent automatically by AeroTracker.")

    body = "\n".join(lines)

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = f"AeroTracker <{GMAIL_USER}>" if GMAIL_USER else EMAIL_FROM
    msg["To"]      = REPORT_TO
    msg.attach(MIMEText(body, "plain"))

    # ── Method 1: Gmail SMTP (GitHub Actions — M365 SMTP is blocked) ─────────
    if GMAIL_USER and GMAIL_APP_PASSWORD:
        try:
            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.ehlo()
                server.starttls()
                server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
                server.sendmail(GMAIL_USER, REPORT_TO, msg.as_string())
            print(f"  ✓ Report emailed via Gmail to {REPORT_TO}")
            return True
        except Exception as e:
            print(f"  ⚠  Gmail send failed: {e}")

    # ── Method 2: Office 365 SMTP (local use) ────────────────────────────────
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.ehlo()
            server.starttls()
            server.login(EMAIL_LOGIN, EMAIL_PASSWORD)
            server.sendmail(EMAIL_FROM, REPORT_TO, msg.as_string())
        print(f"  ✓ Report emailed to {REPORT_TO}")
        return True
    except Exception as e:
        print(f"  ✗ Email failed: {e}")
        return False


def main():
    print("=" * 60)
    print("  NO-CV DAILY REPORT")
    print("=" * 60)

    # Load pending_cv.json
    try:
        with open(PENDING_CV_FILE, "r") as f:
            pending = json.load(f)
    except FileNotFoundError:
        print("  No pending_cv.json — nothing to report.")
        pending = []
    except json.JSONDecodeError:
        print("  ✗ Cannot read pending_cv.json.")
        return

    if not pending:
        print("  Watchlist is empty — nothing to report.")
        return

    jwt = get_jwt()
    today = datetime.date.today()

    still_pending   = []  # back into the JSON
    newly_resolved  = []  # just got their CV — removed from watchlist
    still_waiting   = []  # for email: waiting < OVERDUE_DAYS
    overdue         = []  # for email: waiting >= OVERDUE_DAYS

    for entry in pending:
        rid       = entry.get("resourceId")
        name      = entry.get("name", "Unknown")
        added_str = entry.get("addedDate", "")
        email     = entry.get("email", "")

        try:
            added_date = datetime.date.fromisoformat(added_str)
        except Exception:
            added_date = today
        days_waiting = (today - added_date).days

        print(f"  Checking {name} (ID {rid}, {days_waiting} day(s))...")

        if check_has_cv(jwt, rid):
            print(f"    → CV now uploaded — removing from watchlist")
            newly_resolved.append(name)
        else:
            still_pending.append(entry)
            row = {**entry, "days": days_waiting}
            if days_waiting >= OVERDUE_DAYS:
                overdue.append(row)
                print(f"    → OVERDUE ({days_waiting} days)")
            else:
                still_waiting.append(row)
                print(f"    → Still waiting")

    # Save updated watchlist
    with open(PENDING_CV_FILE, "w") as f:
        json.dump(still_pending, f, indent=2)

    print(f"\n  Still waiting : {len(still_waiting)}")
    print(f"  Overdue       : {len(overdue)}")
    print(f"  Resolved today: {len(newly_resolved)}")

    # Only email if there are candidates still without a CV
    if still_waiting or overdue:
        send_report_email(still_waiting, newly_resolved, overdue)
    else:
        print("  No candidates without CVs — no email sent.")

    print("=" * 60)


if __name__ == "__main__":
    main()
