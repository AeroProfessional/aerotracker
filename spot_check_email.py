"""
spot_check_email.py — Daily quality spot-check email.

Picks 10-15 profiles updated today at random and emails them to
support@aeroprofessional.com so the team can spot-check them in Tracker.

Run daily at 5pm via GitHub Actions (see aerotracker.yml).
"""
import json, os, random, datetime, smtplib, requests
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

DAILY_LOG_FILE    = os.path.join(os.path.dirname(os.path.abspath(__file__)), "daily_updates.json")
EMAIL_FROM        = os.environ.get("EMAIL_LOGIN_USER", "emily.walton@aeroprofessional.com")
EMAIL_PASSWORD    = os.environ.get("EMAIL_PASSWORD",   "PurpleAutumn96?")
GRAPH_TENANT_ID   = os.environ.get("GRAPH_TENANT_ID",    "")
GRAPH_CLIENT_ID   = os.environ.get("GRAPH_CLIENT_ID",    "")
GRAPH_CLIENT_SECRET = os.environ.get("GRAPH_CLIENT_SECRET", "")
SEND_FROM_MAILBOX = "support@aeroprofessional.com"  # Graph API sends as this mailbox
SMTP_SERVER       = "smtp.office365.com"
SMTP_PORT         = 587
SEND_TO           = "support@aeroprofessional.com"
SAMPLE_SIZE       = 15


def _get_graph_token():
    """Get a Graph API token using client credentials (app-only auth)."""
    import msal
    app = msal.ConfidentialClientApplication(
        GRAPH_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}",
        client_credential=GRAPH_CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if result and "access_token" in result:
        return result["access_token"]
    raise Exception(f"Graph auth failed: {result.get('error_description', result)}")

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

    # ── Method 1: Microsoft Graph API (GitHub Actions — SMTP basic auth disabled) ──
    if GRAPH_TENANT_ID and GRAPH_CLIENT_ID and GRAPH_CLIENT_SECRET:
        try:
            token  = _get_graph_token()
            hdrs   = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
            payload = {
                "message": {
                    "subject": subject,
                    "body": {"contentType": "Text", "content": body},
                    "toRecipients": [{"emailAddress": {"address": SEND_TO}}],
                },
                "saveToSentItems": "false",
            }
            r = requests.post(
                f"https://graph.microsoft.com/v1.0/users/{SEND_FROM_MAILBOX}/sendMail",
                json=payload, headers=hdrs, timeout=30,
            )
            if r.status_code in (200, 202):
                print(f"  ✉  Spot-check email sent via Microsoft Graph API ({len(sample)} profiles)")
                return
            print(f"  ⚠  Graph send failed ({r.status_code}): {r.text[:200]}")
        except Exception as e:
            print(f"  ⚠  Graph API send error: {e}")

    # ── Method 2: Outlook COM (Windows) ──────────────────────────────────────
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

    # ── Method 3: SMTP (basic auth — may be blocked by M365 tenant policy) ───
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
