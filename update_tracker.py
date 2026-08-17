"""
=============================================================
  Aero Professional — Tracker Profile Updater  v2
=============================================================
  Processes new candidate registrations from "NEW REGS TO
  ACTION" in Outlook, finds each candidate in Tracker RMS,
  reads their CV, extracts profile data via AI, and updates
  their record — one at a time, with your approval.

  HOW TO RUN:
    py update_tracker.py

  BEFORE FIRST RUN:
    1. Paste your Anthropic API key below (ANTHROPIC_API_KEY)
    2. Save the file
    3. Run it — it will walk you through the first candidate
=============================================================
"""

import re, json, time, io, os, datetime, smtplib, requests, sys, threading
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ── Self-heal: strip any null bytes introduced by editor bugs ──────────────────
def _self_heal_nulls():
    try:
        _p = os.path.abspath(__file__)
        with open(_p, "rb") as _f:
            _d = _f.read()
        if b"" in _d:
            with open(_p, "wb") as _f:
                _f.write(_d.replace(b"", b""))
    except Exception:
        pass
_self_heal_nulls()

# ── Tesseract OCR path (Windows) ───────────────────────────────────────────────
# Auto-detect tesseract binary so image CVs can be OCR'd
def _set_tesseract_path():
    try:
        import pytesseract as _pt
        # Try common Windows install locations
        _candidates = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
            os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\WindowsApps\tesseract.exe"),
            os.path.expandvars(r"%APPDATA%\Local\Programs\Tesseract-OCR\tesseract.exe"),
        ]
        for _c in _candidates:
            if os.path.isfile(_c):
                _pt.pytesseract.tesseract_cmd = _c
                return
        # Last resort: check PATH
        import shutil as _sh
        _which = _sh.which("tesseract")
        if _which:
            _pt.pytesseract.tesseract_cmd = _which
    except ImportError:
        pass
_set_tesseract_path()

# ── Thread safety for parallel processing ─────────────────────────────────────
_outlook_lock     = threading.Lock()   # serialise all Outlook COM calls
_processed_lock   = threading.Lock()   # serialise tracker_processed.json writes
_jwt_holder: list = []                 # [token] — shared mutable JWT for worker threads

# ── Configuration ──────────────────────────────────────────────────────────────

TRACKER_BEARER    = os.environ.get("TRACKER_BEARER",  "b28cae06af044958afb45fa8b1445fa7")
TRACKER_API       = "https://evoglapi.tracker-rms.com"

# ── Email settings ─────────────────────────────────────────────────────────────
EMAIL_FROM        = "support@aeroprofessional.com"
EMAIL_PASSWORD    = os.environ.get("EMAIL_PASSWORD",   "PurpleAutumn96?")
EMAIL_LOGIN_USER  = "emily.walton@aeroprofessional.com"  # Personal account used to access the shared support mailbox
SMTP_SERVER       = "smtp.office365.com"
SMTP_PORT         = 587
# ── Gmail relay (used on GitHub Actions where M365 SMTP is blocked) ───────────
GMAIL_USER        = os.environ.get("GMAIL_USER",         "")
GMAIL_APP_PASSWORD= os.environ.get("GMAIL_APP_PASSWORD", "")
PENDING_CV_FILE   = "pending_cv.json"   # tracks candidates awaiting CV
DAILY_LOG_FILE    = "daily_updates.json"  # tracks profiles updated each day for spot-check email

# ── Email source config ────────────────────────────────────────────────────────
# EMAIL_SOURCE: "subfolder" reads from NEW REGS TO ACTION subfolder.
# When that subfolder is deleted, change to "inbox" to read from the main inbox.
SUPPORT_MAILBOX   = "support@aeroprofessional.com"
EMAIL_SOURCE      = "subfolder"          # reads from REGS subfolder in support inbox
EMAIL_SUBFOLDER   = "REGS"
EMAIL_DONE_FOLDER     = "New regs/Tracker updates"  # subfolder to move processed emails into
EMAIL_NOT_FOUND_FOLDER = "New regs/Not Found"       # subfolder for candidates not found in Tracker

# ── Run mode ───────────────────────────────────────────────────────────────────
# TEST_MODE: process only this many candidates then stop (so you can check results).
# Set to 0 for fully autonomous — processes ALL candidates with no human input.
TEST_MODE         = int(os.environ.get("TEST_MODE", "0"))  # 0 = process all; set env var to limit

# ── Groq LLM — free cloud CV parser ───────────────────────────────────────────
# Free — no credit card needed. One-time setup:
#   1. Go to https://console.groq.com and sign up (free)
#   2. Create an API key and copy it
#   3. Add it as a GitHub Actions secret called GROQ_API_KEY
#      (Settings → Secrets and variables → Actions → New repository secret)
#   4. For local runs, set the environment variable GROQ_API_KEY
#      or paste the key directly below (between the quotes)
GROQ_API_KEY  = os.environ.get("GROQ_API_KEY", "")
GROQ_MODEL    = "llama-3.1-8b-instant"   # fast, accurate, high free-tier limits
_GROQ_WARNED  = False   # printed once if key is missing

# DRY_RUN: show proposed changes but do NOT write anything to Tracker.
# Set to False only once you're confident the parsing is correct.
DRY_RUN           = False  # ← change to False to enable actual updates

# SEND_CV_REQUESTS: when False, candidates with no CV are skipped silently
# (no email sent, no watchlist entry). Set to True to re-enable CV request emails.
SEND_CV_REQUESTS  = False

# PROTECTED_IDS: Tracker resource IDs that must NEVER be auto-updated by the main run.
# These profiles have been manually corrected via reprocess_batch.py.
# Even if tracker_processed.json is cleared, these will still be protected.
PROTECTED_IDS = {
    74070,   # Mikko Vainio        — Finland, B1, B2
    106282,  # kawinraj Yuganandam — Bahrain
    139906,  # TAIF HAMEED         — Iraq, Greece, EASA, First Officer
    139907,  # Ibrahim Zubair      — Iraq, Greece, EASA, First Officer, B737-NG
    106766,  # Josè Fiori          — correct from initial run
    139953,  # Dipanshu Khatri     — India, India, ICAO, Flight Instructor, FSTC
    89019,   # Mohammad Aghdaie    — wrong-match victim: First Officer, Iran Air, Iran, ICAO, A320, A330
    129932,  # Rebekah Gaston      — Cabin Crew, United Kingdom, First Class, Main Crew, VIP
}

# Collects names of candidates skipped because they have no CV — printed at end of run.
_NO_CV_NAMES: list = []

# ── Microsoft Graph API ────────────────────────────────────────────────────────
# Required to read emails from the shared mailbox (replaces broken Outlook COM).
# ONE-TIME SETUP (5 minutes):
#   1. Go to https://portal.azure.com → search "App registrations" → New registration
#   2. Name: "Tracker Email Reader" | Accounts: "Single tenant" | No redirect URI → Register
#   3. Copy "Application (client) ID"  → paste as GRAPH_CLIENT_ID below
#   4. Copy "Directory (tenant) ID"    → paste as GRAPH_TENANT_ID below
#   5. API permissions → Add → Microsoft Graph → Delegated → add:
#        Mail.Read   and   Mail.ReadWrite
#   6. (Optional) Grant admin consent — or Emily will be prompted on first sign-in
#   7. pip install msal
#
# After setup, the first run will show a code and URL — sign in once, token is saved.
GRAPH_TENANT_ID    = os.environ.get("GRAPH_TENANT_ID",    "")  # Directory (tenant) ID from Azure AD app
GRAPH_CLIENT_ID    = os.environ.get("GRAPH_CLIENT_ID",    "")  # Application (client) ID from Azure AD app
GRAPH_CLIENT_SECRET= os.environ.get("GRAPH_CLIENT_SECRET","")  # Client secret — enables app-only auth on GitHub Actions
GRAPH_TOKEN_FILE   = os.path.join(os.path.expanduser("~"), "tracker_graph_token.json")

# ── Tracker web session cookie (for CV downloads) ──────────────────────────────
# To refresh: in Chrome, open Tracker → F12 → Network → download any CV →
# click the request → Request Headers → copy the Cookie: value and paste below.
TRACKER_WEB_COOKIE = "_fw_crm_v=aea0e107-8e84-4bb4-8999-58e6ef59af63; TrackerRMSResumePreviewTab=#resumeTab_Activity; ASP.NET_SessionId=x1wvkvxsufpwmdsq5etzkwnw; TrackerRMSSummary=show; trackerlastlogonts=20260624141143; token=2Ur2I04BcRuv69BuAkAXniyaTa8Q3D8vZHaQmCNN8RyQTfmQ+JO+V3dXVfGfKjY03jYkFJolGANhQvU+FoRn+2cEBOxxpLal8YfOquVhc7VFvxGKCbZc9V4meMkPNq70iYc3iGjQZPEZmbJFRASgU1ly+1w9nscOjZEfPg4BlLdwNOb/USoBKHdEF1rdvuvieAVhVS4sKVo6sH6wWlY2WrGbnToAUfjCL1qEeQjoGBaqI634du7rfuaEq3ZKzLGnkaKslpLCRxaf7wyJYX1n6g==; first_session=%7B%22visits%22%3A10032%2C%22start%22%3A1768923804005%2C%22last_visit%22%3A1782308067545%2C%22url%22%3A%22https%3A%2F%2Fevouk.tracker-rms.com%2FActivity%3FView%3D00%22%2C%22path%22%3A%22%2FActivity%22%2C%22referrer%22%3A%22https%3A%2F%2Fmy.tracker-rms.com%2F%22%2C%22referrer_info%22%3A%7B%22host%22%3A%22my.tracker-rms.com%22%2C%22path%22%3A%22%2F%22%2C%22protocol%22%3A%22https%3A%22%2C%22port%22%3A80%2C%22search%22%3A%22%22%2C%22query%22%3A%7B%7D%7D%2C%22search%22%3A%7B%22engine%22%3Anull%2C%22query%22%3Anull%7D%2C%22prev_visit%22%3A1782308066991%2C%22time_since_last_visit%22%3A554%2C%22version%22%3A0.4%7D"

# ── Cookie age note ───────────────────────────────────────────────────────────
# TRACKER_WEB_COOKIE is now a last-resort fallback only. CVs are downloaded via
# the Tracker REST API (JWT auth, auto-refreshed every run) so the cookie is not
# normally needed. It is only used if the REST API endpoint fails for a candidate.
_COOKIE_SET_DATE = "20260624"  # date the current cookie was captured (YYYYMMDD)

# No Anthropic API key needed — CV parsing is rule-based and completely free

# ── Candidate list (names extracted from "NEW REGS TO ACTION" emails) ──────────
# Both email types are included:
#   "Candidate Registration through Website" → "[Name] has registered..."
#   "New Candidate Application - Jobs+"      → "[Name] has applied for Job XXXX"

CANDIDATES = [
    # ── Test profile ──
    "Emily Test",
    # ── Registrations ──
    "Guillaume ZOUARI",
    "Ali Alfarid",
    "Loren Mae Batobalonos",
    "Muhammad Uzair",
    "Neetin Vatsya",
    # ── Jobs+ applications ──
    "Taha Mahmood",
    "Nader Qulays",
    "ASAD UR RAHMAN",
    "SALMAN FARIS",
    "Abdul Rahuman Syed Ibrahim",
    # Add more here as needed
]

# ── Work type IDs ──────────────────────────────────────────────────────────────

WORK_TYPES = {
    "flight deck":  {"id": 472, "name": "Flight Deck"},
    "cabin crew":   {"id": 469, "name": "Cabin Crew"},
    "engineering":  {"id": 471, "name": "Engineering"},
    "head office":  {"id": 474, "name": "Head Office"},
    "management":   {"id": 470, "name": "Managment"},  # Tracker typo kept
    "managment":    {"id": 470, "name": "Managment"},  # alias for Tracker typo
    "operations":   {"id": 473, "name": "Operations"},
    "airport":      {"id": 468, "name": "Airport"},
    "executive":    {"id": 475, "name": "Executive"},
}

# Skill area IDs — used to fetch individual skill entries
SKILL_AREAS = {
    "nationality":        43,
    "aircraft":           40,
    "country":            39,
    "issuing_authority":  38,
    "position":           37,
    "tri_tre":            44,
    "flight_instructor":  45,   # skillId 1077 = "Flight instructor"
    "cabin_seniority":    48,
    "cabin_service":      47,
    "eng_licence":        46,
    "easa_atpl_a":        19,
    "easa_b1":            15,
    "easa_b2":            17,
    "easa_cpl_a":         21,
    "faa_atpl_a":          7,
    "faa_b1":              3,
    "faa_b2":              5,
    "faa_cpl_a":           9,
    "icao_atpl_a":        31,
    "icao_b1":            27,
    "icao_b2":            29,
    "icao_cpl_a":         33,
}

# ── Email: CV request ──────────────────────────────────────────────────────────

CV_REQUEST_SUBJECT = "Please upload your CV — Aero Professional"

CV_REQUEST_BODY = """\
Dear {name},

Thank you for registering with Aero Professional.

To complete your candidate profile we need a copy of your CV. Please reply to this email with your CV attached, or upload it directly to your profile.

If you have any questions, please do not hesitate to contact us.

Kind regards,
Emily
"""


def send_cv_request_email(to_address, candidate_name):
    """
    Send a CV request email via Outlook desktop app (no SMTP auth needed).
    Outlook must be installed and logged in.
    """
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = MailItem
        mail.To      = to_address
        mail.Subject = CV_REQUEST_SUBJECT
        mail.Body    = CV_REQUEST_BODY.format(name=candidate_name.split()[0])
        mail.Send()
        return True
    except Exception as e:
        print(f"  ✗ Outlook error: {e}")
        return False


def send_run_summary_email(done, skipped, already_done, error_summary):
    """
    Email the team a summary of the tracker update run.
    Uses Outlook COM on Windows; falls back to SMTP on Linux/GitHub Actions.
    """
    if error_summary:
        subject = f"⚠ Tracker Update — {len(error_summary)} candidate(s) need attention"
    else:
        subject = f"✓ Tracker Update complete — {done} processed"

    lines = [
        "Tracker update run complete.",
        "",
        f"  Processed successfully : {done}",
        f"  Skipped / not found    : {skipped}",
    ]
    if already_done:
        lines.append(f"  Already done (prev run): {already_done}")
    if error_summary:
        lines += ["", "The following candidates did NOT complete — please review:", ""]
        for cname, reason in error_summary:
            lines.append(f"  • {cname}  —  {reason}")
        lines += ["", "Full error details are in run_errors.txt in the script folder."]
    else:
        lines += ["", "No errors."]
    body = "\n".join(lines)

    # ── Method 1: Outlook COM (Windows only) ──────────────────────────────────
    try:
        try:
            import pythoncom as _pc_sum
            _pc_sum.CoInitialize()
        except Exception:
            pass
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To      = "support@aeroprofessional.com"
        mail.Subject = subject
        mail.Body    = body
        mail.Send()
        print("  ✉  Run summary emailed via Outlook")
        return
    except Exception:
        pass  # fall through to SMTP

    # ── Method 2: Gmail SMTP (works on GitHub Actions — M365 SMTP is blocked) ───
    if GMAIL_USER and GMAIL_APP_PASSWORD:
        try:
            msg = MIMEMultipart("alternative")
            msg["Subject"] = subject
            msg["From"]    = f"AeroTracker <{GMAIL_USER}>"
            msg["To"]      = "support@aeroprofessional.com"
            msg.attach(MIMEText(body, "plain"))
            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.ehlo()
                server.starttls()
                server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
                server.sendmail(GMAIL_USER, "support@aeroprofessional.com", msg.as_string())
            print("  ✉  Run summary emailed via Gmail")
            return
        except Exception as e:
            print(f"  ⚠  Gmail send failed: {e}")

    # ── Method 3: Microsoft Graph API (works on GitHub Actions when SMTP is blocked) ──
    if GRAPH_TENANT_ID and GRAPH_CLIENT_ID and GRAPH_CLIENT_SECRET:
        try:
            _token = get_graph_token()
            _hdrs  = {"Authorization": f"Bearer {_token}", "Content-Type": "application/json"}
            _payload = {
                "message": {
                    "subject": subject,
                    "body": {"contentType": "Text", "content": body},
                    "toRecipients": [{"emailAddress": {"address": "support@aeroprofessional.com"}}],
                },
                "saveToSentItems": "false",
            }
            _r = requests.post(
                f"https://graph.microsoft.com/v1.0/users/{SUPPORT_MAILBOX}/sendMail",
                json=_payload, headers=_hdrs, timeout=30,
            )
            if _r.status_code in (200, 202):
                print("  ✉  Run summary emailed via Microsoft Graph API")
                return
            print(f"  ⚠  Graph send failed ({_r.status_code}): {_r.text[:200]}")
        except Exception as _ge:
            print(f"  ⚠  Graph API send error: {_ge}")

    # ── Method 3: SMTP (basic auth — may be blocked by M365 tenant policy) ───
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = EMAIL_FROM
        msg["To"]      = "support@aeroprofessional.com"
        msg.attach(MIMEText(body, "plain"))
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.ehlo()
            server.starttls()
            server.login(EMAIL_LOGIN_USER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_FROM, "support@aeroprofessional.com", msg.as_string())
        print("  ✉  Run summary emailed to support@aeroprofessional.com")
    except Exception as e:
        print(f"  ⚠  Could not send summary email: {e}")


def save_pending_cv(resource_id, candidate_name, candidate_email):
    """Add candidate to the pending-CV watchlist (pending_cv.json)."""
    today = datetime.date.today().isoformat()
    try:
        with open(PENDING_CV_FILE, "r") as f:
            pending = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        pending = []

    # Avoid duplicates
    pending = [p for p in pending if p.get("resourceId") != resource_id]
    pending.append({
        "resourceId":    resource_id,
        "name":          candidate_name,
        "email":         candidate_email,
        "addedDate":     today,
    })

    with open(PENDING_CV_FILE, "w") as f:
        json.dump(pending, f, indent=2)


def log_daily_update(resource_id, candidate_name, job_title, employer, work_types, skills):
    """Append a successfully updated profile to today's daily log for the spot-check email."""
    today = datetime.date.today().isoformat()
    try:
        with open(DAILY_LOG_FILE, "r") as f:
            log = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        log = {}
    if today not in log:
        log[today] = []
    # Avoid duplicates within the same day
    log[today] = [e for e in log[today] if e.get("id") != resource_id]
    log[today].append({
        "id":        resource_id,
        "name":      candidate_name,
        "job_title": job_title,
        "employer":  employer,
        "work_type": ", ".join(w.get("name", "") for w in work_types),
        "skills":    [s.get("name", "") for s in skills],
    })
    # Keep only last 14 days to avoid file growing indefinitely
    cutoff = (datetime.date.today() - datetime.timedelta(days=14)).isoformat()
    log = {d: v for d, v in log.items() if d >= cutoff}
    with open(DAILY_LOG_FILE, "w") as f:
        json.dump(log, f, indent=2)


# ── Tracker auth ───────────────────────────────────────────────────────────────

def _send_smtp_alert(subject, body):
    """Send a plain-text alert email to support@ via SMTP. Used for bearer token expiry etc."""
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = EMAIL_FROM
        msg["To"]      = "support@aeroprofessional.com"
        msg.attach(MIMEText(body, "plain"))
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.ehlo()
            server.starttls()
            server.login(EMAIL_LOGIN_USER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_FROM, "support@aeroprofessional.com", msg.as_string())
        print(f"  ✉  Alert sent: {subject}")
    except Exception as _ae:
        print(f"  ⚠  Could not send alert email: {_ae}")


def get_jwt():
    try:
        r = requests.post(f"{TRACKER_API}/api/Auth/ExchangeToken",
                          json={"bearerToken": TRACKER_BEARER}, timeout=15)
    except Exception as e:
        raise Exception(f"Tracker API unreachable: {e}")
    if r.status_code == 401:
        _send_smtp_alert(
            "⚠ AeroTracker — Tracker API token has EXPIRED",
            "The Tracker RMS bearer token has expired and must be renewed.\n\n"
            "Steps:\n"
            "  1. Log into Tracker RMS\n"
            "  2. Go to Tools & Settings → API Access\n"
            "  3. Generate a new bearer token\n"
            "  4. Update the TRACKER_BEARER secret in GitHub → Settings → Secrets\n\n"
            "Until the token is renewed, no candidate profiles will be updated.\n\n"
            "This message was sent automatically by AeroTracker."
        )
        raise Exception("Tracker bearer token expired (HTTP 401) — renew it in Tracker → Tools & Settings")
    r.raise_for_status()
    return r.json()["token"]

def h(jwt):
    return {"Authorization": f"Bearer {jwt}", "Content-Type": "application/json"}

# ── Skills lookup ──────────────────────────────────────────────────────────────

def load_all_skills(jwt):
    """
    Fetch every skill from every area and return a name→{id,name} dict.
    Field names from Tracker API: skillId, skillName.
    """
    lookup = {}
    licence_country_lookup = {}  # area-39 only: name → {id, name}
    country_skills_set  = set()  # all country/nationality skill names (areas 39+43)
    nationality_ids     = set()  # skill IDs from area 43 (candidate's nationality)
    licence_country_ids = set()  # skill IDs from area 39 (licence-issuing country)
    seen_areas = set()
    for label, area_id in SKILL_AREAS.items():
        if area_id in seen_areas:
            continue
        seen_areas.add(area_id)
        try:
            # Fetch ALL skills in this area, handling pagination.
            # Strategy: try increasing page numbers; stop on empty page OR duplicate page.
            all_skills = []
            seen_ids = set()
            page = 1
            max_pages = 50  # safety cap
            while page <= max_pages:
                r = requests.get(
                    f"{TRACKER_API}/api/v1/MetaData/Skills/Areas/{area_id}/Skills",
                    params={"pageSize": 100, "pageNumber": page},
                    headers=h(jwt), timeout=15
                )
                if r.status_code != 200:
                    if label in ("nationality", "country"):
                        print(f"  ⚠  Skills area '{label}' (ID {area_id}) returned HTTP {r.status_code}: {r.text[:200]}")
                    break
                data = r.json()
                page_skills = data.get("skills", data) if isinstance(data, dict) else data
                if not isinstance(page_skills, list) or not page_skills:
                    break  # empty page — done
                # Duplicate-page detection: if every ID on this page was already seen, the
                # API is returning the same data (ignoring pageNumber) — stop immediately.
                page_ids = {s.get("skillId") or s.get("id") for s in page_skills}
                if page_ids and page_ids.issubset(seen_ids):
                    break  # duplicate page — API doesn't support pagination
                new_skills = [s for s in page_skills
                              if (s.get("skillId") or s.get("id")) not in seen_ids]
                all_skills.extend(new_skills)
                seen_ids.update(page_ids)
                if len(page_skills) < 100:
                    break  # last page (partial)
                page += 1

            count_before = len(lookup)
            for s in all_skills:
                sid  = s.get("skillId") or s.get("id")
                name = s.get("skillName") or s.get("name")
                if sid and name:
                    key = name.strip().lower()
                    if label == "country":
                        # Area-39 (licence country): store in dedicated lookup.
                        # Only add to main lookup if no area-43 entry already set
                        # so CV parsing always resolves countries to nationality IDs.
                        licence_country_lookup[key] = {"id": sid, "name": name.strip()}
                        if key not in lookup:
                            lookup[key] = {"id": sid, "name": name.strip()}
                    elif label == "position":
                        # Area-37 returns two IDs per position name (e.g. 182 and 1079 for
                        # "First Officer"). The higher IDs (1078=Captain, 1079=FO) are the
                        # correct quickSkills IDs — the lower ones (181, 182) are new-API-only
                        # and are silently dropped by Tracker's quickSkills PATCH.
                        existing = lookup.get(key)
                        if existing and existing["id"] > sid:
                            pass  # keep the already-stored higher ID
                        else:
                            lookup[key] = {"id": sid, "name": name.strip()}
                    else:
                        lookup[key] = {"id": sid, "name": name.strip()}
                    if label in ("nationality", "country"):
                        country_skills_set.add(key)
                    if label == "nationality":
                        nationality_ids.add(sid)
                    elif label == "country":
                        licence_country_ids.add(sid)
            added = len(lookup) - count_before
            if label in ("nationality", "country"):
                sample = list(lookup.keys())[-5:] if lookup else []
                print(f"  ℹ  Skills area '{label}' (ID {area_id}): {len(all_skills)} skills loaded ({added} new). Sample: {sample}")
                debug_keys = [k for k in lookup if any(x in k for x in ["iraq","india","united kingdom","ukraine","morocco"])]
                if debug_keys:
                    print(f"     Key nationalities found: {sorted(debug_keys)}")
        except Exception as e:
            print(f"  Warning: could not load skills area {area_id} ({label}): {e}")
    return lookup, country_skills_set, nationality_ids, licence_country_ids, licence_country_lookup

# Reverse of COUNTRY_ALIASES: country name → adjective form (e.g. "India" → "indian").
# Built once so resolve_skills can fall back to adjective form when Tracker stores
# nationalities as adjectives rather than country names.
def resolve_skills(names, lookup):
    """Map skill name strings → Tracker quickSkill objects {id, name}."""
    # Skills we never want to apply — Tracker has them but they're not useful
    SKILL_BLOCKLIST = {
        "british indian ocean territory",
    }
    # Licence types are never Tracker skills — the authority (ICAO/FAA/EASA) is used instead
    LICENCE_TYPE_BLOCKLIST = {
        "atpl", "atpl(a)", "atpl(h)", "cpl", "cpl(a)", "cpl(h)",
        "ppl", "ppl(a)", "ppl(h)", "ir", "me/ir", "se/ir",
        "faa atpl", "easa atpl", "icao atpl",
    }
    result = []
    seen = set()
    for raw in names:
        key = raw.strip().lower()
        if key in seen:
            continue
        seen.add(key)
        # Never add licence type strings — authority (ICAO/FAA/EASA) is used instead
        if key in LICENCE_TYPE_BLOCKLIST:
            continue
        # Prefer area-43 demonym over area-39 country name when both exist.
        # e.g. CV says "Finland" — prefer "Finnish" (area-43 nationality ID) over
        # "Finland" (area-39 FCL country ID), otherwise the engineering/management
        # rule strips it as a licence-country skill, leaving no nationality at all.
        rev_adj = _REVERSE_COUNTRY_ALIASES.get(key, "")
        if rev_adj and rev_adj in lookup:
            result.append(lookup[rev_adj])
            continue
        # Direct match (used for area-39 FCL countries like "Greece", "India" when
        # explicitly requested as fcl_country rather than extracted from CV text)
        if key in lookup:
            result.append(lookup[key])
            continue
        # Try country alias adjective → country name (e.g. "Filipino" → "Philippines")
        alias = COUNTRY_ALIASES.get(key, "")
        if alias and alias.lower() in lookup:
            result.append(lookup[alias.lower()])
            continue
        # Startswith match — catches "Saudi Arabia, Kingdom of" when searching "saudi arabia"
        sw_matches = [(k, v) for k, v in lookup.items()
                      if k not in SKILL_BLOCKLIST and (k.startswith(key) or key.startswith(k))
                      and len(k) >= 4 and len(key) >= 4]
        if sw_matches:
            sw_matches.sort(key=lambda x: len(x[0]))
            result.append(sw_matches[0][1])
            continue
        # Partial match — prefer shorter matches (closer to exact), min 4 chars
        matches = [(k, v) for k, v in lookup.items()
                   if len(k) >= 4 and len(key) >= 4 and (key == k or key in k or k in key)
                   and k not in SKILL_BLOCKLIST]
        if matches:
            # Sort by closeness: exact first, then shortest key
            matches.sort(key=lambda x: (x[0] != key, len(x[0])))
            result.append(matches[0][1])
        else:
            # Never send id=0 for a known country/nationality name — that creates a
            # free-text skill in the wrong Tracker area.  Log and skip instead.
            if key in _ALL_KNOWN_COUNTRY_NAMES:
                print(f"  ⚠  Country '{raw}' not found in Tracker skills lookup — skipped "
                      f"(check that Tracker skill areas 39/43 contain this country)")
            else:
                print(f"  ⚠  Skill not in lookup: '{raw}' — adding as free text")
                result.append({"id": 0, "name": raw.strip()})
    return result

# ── Candidate index (full scan) ────────────────────────────────────────────────

def build_candidate_index(jwt):
    """
    Page through all resources and build a name→resourceId lookup.
    Also collects quickSkills from existing records to enrich the skills lookup.
    Also collects candidates with incomplete profiles (no job title or no skills)
    so GitHub Actions can process them directly without reading emails.
    Returns: (name_index dict, extra_skills dict, incomplete_candidates list)
    """
    print("  Loading all candidates from Tracker (this takes ~60 seconds)...")
    name_index   = {}  # "firstname surname" → resourceId
    extra_skills = {}  # name.lower() → {id, name}
    incomplete_candidates = []  # candidates that appear to need profile completion

    _PLACEHOLDER_TITLES_QUICK = {
        "unknown", "n/a", "tbd", "to be determined", "candidate",
        "pilot candidate", "cabin crew candidate", "crew candidate",
        "flight deck candidate", "aviation candidate", "applicant",
        "seeking", "open to work",
    }

    import unicodedata
    def norm(s):
        """Strip accents so Ünver → unver, Çelik → celik, etc."""
        return unicodedata.normalize("NFD", s).encode("ascii", "ignore").decode("ascii").lower()

    PAGE_SIZE = 50   # larger pages = fewer round trips
    # No MAX_PAGES cap — scan until Tracker returns an empty page

    def scan_status(status_filter=None):
        label = status_filter or "all statuses"
        if status_filter:
            print(f"  Also scanning Pre-Registered candidates...")
        page = 1
        empty_streak = 0
        no_progress_streak = 0
        retries = 0
        prev_index_size = len(name_index)
        while True:
            try:
                body = {"pageSize": PAGE_SIZE, "pageNumber": page}
                if status_filter:
                    body["resourceStatus"] = status_filter
                r = requests.post(
                    f"{TRACKER_API}/api/v1/Resource/Search",
                    json=body, headers=h(jwt), timeout=20
                )
                if r.status_code != 200:
                    break
                try:
                    _raw = r.json()
                except Exception:
                    page += 1
                    continue
                # Handle both list response and dict response (items/value/results)
                if isinstance(_raw, list):
                    items = _raw
                elif isinstance(_raw, dict):
                    items = (_raw.get("items") or _raw.get("value") or
                             _raw.get("results") or _raw.get("data") or [])
                else:
                    items = []
                if not items:
                    empty_streak += 1
                    if empty_streak >= 3:
                        break
                    page += 1
                    continue

                empty_streak = 0
                retries = 0
                for rec in items:
                    rid   = rec.get("resourceId")
                    first = (rec.get("firstname") or rec.get("firstName") or "").strip()
                    sur   = (rec.get("surname") or rec.get("lastName") or "").strip()
                    # Some records have name in a combined field with empty first/surname
                    if rid and not first and not sur:
                        _fn_raw = (rec.get("name") or rec.get("resourceName") or
                                   rec.get("fullName") or "").strip()
                        if _fn_raw:
                            _fp = _fn_raw.split()
                            first = _fp[0] if _fp else ""
                            sur   = " ".join(_fp[1:]) if len(_fp) > 1 else ""
                    if rid and (first or sur):
                        full_key = f"{first} {sur}".strip().lower()
                        # Store lists for duplicate names (e.g. two people named "Ed Nawar")
                        _existing = name_index.get(full_key)
                        if _existing is None:
                            name_index[full_key] = rid
                        elif isinstance(_existing, int) and _existing != rid:
                            name_index[full_key] = [_existing, rid]
                        elif isinstance(_existing, list) and rid not in _existing:
                            name_index[full_key] = _existing + [rid]
                        # Also store accent-stripped version so "Ünver" matches "Unver"
                        norm_key = norm(f"{first} {sur}".strip())
                        if norm_key != full_key:
                            _existing_n = name_index.get(norm_key)
                            if _existing_n is None:
                                name_index[norm_key] = rid
                            elif isinstance(_existing_n, int) and _existing_n != rid:
                                name_index[norm_key] = [_existing_n, rid]
                            elif isinstance(_existing_n, list) and rid not in _existing_n:
                                name_index[norm_key] = _existing_n + [rid]
                        if sur:
                            name_index[sur.lower()] = rid
                            norm_sur = norm(sur)
                            if norm_sur != sur.lower():
                                name_index[norm_sur] = rid
                    for qs in rec.get("quickSkills", []):
                        sid   = qs.get("id")
                        sname = qs.get("name")
                        if sname and sid is not None:  # include id=0 (free text) skills
                            key = sname.strip().lower()
                            # Prefer a real ID over free-text (id=0)
                            if key not in extra_skills or extra_skills[key]["id"] == 0:
                                extra_skills[key] = {"id": sid, "name": sname.strip()}

                # Stop only when the API has stopped producing new candidates for an
                # extended run. Use a high threshold (500 pages = 25,000 records) so
                # Tracker's overlapping/non-monotonic pagination doesn't cause early exit.
                # If we've seen more pages with no progress than the entire estimated
                # database size, we've definitely wrapped around — stop then too.
                # ── Collect incomplete profiles during the same scan ───────────
                for _rec in items:
                    _rid = _rec.get("resourceId")
                    if not _rid:
                        continue
                    _first = (_rec.get("firstname") or _rec.get("firstName") or "").strip()
                    _sur   = (_rec.get("surname")   or _rec.get("lastName")   or "").strip()
                    _name  = f"{_first} {_sur}".strip()
                    if not _name:
                        continue
                    _job    = (_rec.get("jobTitle") or _rec.get("currentPosition") or "").strip()
                    _skills = _rec.get("quickSkills") or []
                    _needs  = (
                        not _job or
                        _job.lower() in _PLACEHOLDER_TITLES_QUICK or
                        not _skills
                    )
                    if _needs:
                        incomplete_candidates.append({
                            "name":        _name,
                            "item":        None,
                            "email_id":    f"tracker:{_rid}",
                            "graph_token": None,
                        })
                # ─────────────────────────────────────────────────────────────

                new_size = len(name_index)
                if new_size == prev_index_size:
                    no_progress_streak += 1
                    # Stop if: 500 consecutive all-duplicate pages, OR we've gone 3× the
                    # estimated total pages without a single new entry (wrap-around guard).
                    estimated_total = max(500, prev_index_size // 2)
                    if no_progress_streak >= 500 or (page > estimated_total * 3 and no_progress_streak > 50):
                        print(f"  Index scan complete — {no_progress_streak} consecutive duplicate pages at page {page}")
                        break
                else:
                    no_progress_streak = 0
                prev_index_size = new_size

                page += 1
                time.sleep(0.02)

                if page % 25 == 0:
                    print(f"  ...page {page}, {len(name_index)} entries so far...")

            except Exception as e:
                retries += 1
                if retries <= 3:
                    print(f"  Connection error on page {page} — retrying ({retries}/3)...")
                    time.sleep(2)
                else:
                    print(f"  Warning: giving up on page {page} ({label}): {e}")
                    break

    scan_status(None)
    scan_status("Pre-Registered")
    scan_status("Active")
    scan_status("Inactive")

    # Show highest resource ID indexed (helps verify new registrations are included)
    all_ids = [v for v in name_index.values() if isinstance(v, int)]
    if all_ids:
        print(f"  Loaded {len(name_index)} name entries total (highest ID: {max(all_ids)})")
    else:
        print(f"  Loaded {len(name_index)} name entries total")
    # Deduplicate incomplete list by resource ID (multiple status scans may add duplicates)
    _seen_inc = set()
    incomplete_candidates = [
        c for c in incomplete_candidates
        if c["email_id"] not in _seen_inc and not _seen_inc.add(c["email_id"])
    ]
    print(f"  Candidates with incomplete profiles: {len(incomplete_candidates)}")
    return name_index, extra_skills, incomplete_candidates


# Name abbreviation expansions — module-level so process_one can reference it too
FIRST_NAME_VARIANTS = {
    "mohd": ["mohammed", "muhammad", "mohamed", "mohammad"],
    "mohammed": ["mohd", "muhammad", "mohamed", "mohammad"],
    "muhammad": ["mohd", "mohammed", "mohamed", "mohammad"],
    "mohamed": ["mohd", "mohammed", "muhammad", "mohammad"],
    "abdulrahman": ["abd al-rahman", "abdulrahman", "abd rahman"],
    "abd": ["abdulrahman", "abdullah", "abdulaziz"],
}


def find_candidate(full_name, name_index, jwt=None):
    """Find a candidate's resourceId — checks index first, then live API search."""
    import unicodedata as _ucdn

    # ── Name pre-processing ────────────────────────────────────────────────────
    # Detect Arabic/non-Latin script — can't match these reliably
    _arabic_chars = sum(1 for c in full_name if '؀' <= c <= 'ۿ')
    if _arabic_chars > len(full_name) * 0.4:
        print(f"  ⚠ Name is in Arabic script — cannot match automatically: '{full_name}'")
        return None

    _RANK_WORDS = [
        "air chief marshal", "air vice-marshal", "air vice marshal", "air marshal",
        "air commodore", "wing commander", "squadron leader", "flight lieutenant",
        "group captain", "commodore", "rear admiral", "vice admiral", "admiral",
        "lieutenant general", "major general", "brigadier general", "brigadier",
        "lieutenant colonel", "colonel", "lieutenant commander", "commander",
        "major", "captain", "lieutenant", "sergeant", "corporal",
        "air commodore retd", "retd", "ret'd", "retired",
        "dr", "prof", "professor", "sir", "dame", "lord", "lady",
    ]

    _cleaned = full_name.strip()
    # Strip invisible Unicode characters (RTL marks, zero-width spaces etc.)
    _cleaned = "".join(c for c in _cleaned if not _ucdn.category(c).startswith('C'))
    # Strip stray special characters that shouldn't appear in names
    _cleaned = re.sub(r"[#@\[\]{}|\\`']", '', _cleaned).strip()
    # Normalize multiple spaces
    _cleaned = re.sub(r'\s+', ' ', _cleaned).strip()
    # Normalize " - " separators in double-barrelled names
    _cleaned = re.sub(r'\s+-\s+', '-', _cleaned)
    # Split camelCase words in names: "UrRehman" → "Ur Rahman", "AlGhandaf" → "Al Ghandaf"
    _cleaned = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', _cleaned)
    _cleaned = re.sub(r'\s+', ' ', _cleaned).strip()

    # Deduplicate: "Abraham Orozco Abraham Orozco" → "Abraham Orozco"
    _wds = _cleaned.split()
    if len(_wds) >= 4 and len(_wds) % 2 == 0:
        _h = len(_wds) // 2
        if [w.lower() for w in _wds[:_h]] == [w.lower() for w in _wds[_h:]]:
            _cleaned = " ".join(_wds[:_h])

    # Strip rank/title suffixes from end of name (case-insensitive)
    _cleaned_lc = _cleaned.lower()
    for _rank in sorted(_RANK_WORDS, key=len, reverse=True):
        if _cleaned_lc.endswith(" " + _rank):
            _before = _cleaned[:-(len(_rank) + 1)].strip()
            if len(_before.split()) >= 2:
                _cleaned = _before
                _cleaned_lc = _cleaned.lower()
                break

    if _cleaned.strip() != full_name.strip():
        print(f"  ~ Name pre-processed: '{full_name.strip()}' → '{_cleaned}'")
    full_name = _cleaned.strip()

    # Build "bin/binti" alternative: "A B bin C" → try searching as "A B C" (drop bin)
    _BIN_RE  = re.compile(r'\s+\b(bin|binti|bte|bt|ibn)\b\s+', re.IGNORECASE)
    _bin_alt = None
    if _BIN_RE.search(full_name):
        _bin_alt = _BIN_RE.sub(' ', full_name).strip()
        _bin_alt = re.sub(r'\s+', ' ', _bin_alt).strip()
    # ──────────────────────────────────────────────────────────────────────────

    parts   = full_name.strip().split()
    first   = parts[0].lower() if parts else ""
    sur     = " ".join(parts[1:]).lower() if len(parts) > 1 else ""
    full_lc = full_name.strip().lower()

    # (FIRST_NAME_VARIANTS defined at module level above)

    # 1. Exact full name in index
    if full_lc in name_index:
        _hit = name_index[full_lc]
        if isinstance(_hit, list):
            # Multiple people share this name — return as list of (key, rid) tuples
            return [(full_lc, rid) for rid in _hit]
        return _hit
    # 2. Reverse order (Tracker stores some as SURNAME FIRSTNAME)
    rev = f"{sur} {first}".strip()
    if rev in name_index:
        _hit = name_index[rev]
        if isinstance(_hit, list):
            return [(rev, rid) for rid in _hit]
        return _hit
    # 2b. First name variants (Mohd → Mohammed etc.)
    for alt_first in FIRST_NAME_VARIANTS.get(first, []):
        alt_key = f"{alt_first} {sur}".strip()
        if alt_key in name_index:
            print(f"  ~ Name variant match: '{full_name}' → '{alt_key}'")
            return name_index[alt_key]
        alt_rev = f"{sur} {alt_first}".strip()
        if alt_rev in name_index:
            print(f"  ~ Name variant match: '{full_name}' → '{alt_rev}'")
            return name_index[alt_rev]
    # 2c. Hyphen normalization — Tracker may store "Mihai-George" as "Mihai George"
    # or "Al-Dougish" as "Al Dougish" and vice versa.  Try both directions.
    if "-" in full_lc:
        hyph_lc = full_lc.replace("-", " ")
        if hyph_lc in name_index:
            return name_index[hyph_lc]
        # also try accent-stripped hyphen-normalised version
        import unicodedata as _uh
        def _nh(s):
            return re.sub(r"[^a-z0-9 ]", "",
                          _uh.normalize("NFD", s).encode("ascii","ignore").decode("ascii").lower()).strip()
        hyph_norm = _nh(hyph_lc)
        for k, v in name_index.items():
            if _nh(k) == hyph_norm:
                return v
    # Reverse: index key may have hyphen but search name has space
    #   e.g. search "mihai george anghelina" vs index "mihai-george anghelina"
    if " " in first:
        hyph_first = first.replace(" ", "-")
        alt_lc = f"{hyph_first} {sur}".strip()
        if alt_lc in name_index:
            return name_index[alt_lc]

    # 3. Surname-only index lookup — only if surname is unique AND first initial agrees.
    # This prevents "Syed Tanvir" matching "M.Amjad Tanvir" (same surname, different people).
    if sur and sur in name_index and "." not in first:
        sur_count = sum(1 for k in name_index if k.endswith(f" {sur}") or k == sur)
        if sur_count <= 2:
            # Also require that the found name's first initial matches the search first initial
            matched_key = next((k for k in name_index if k == sur or k.endswith(f" {sur}")), None)
            matched_first_initial = ""
            if matched_key:
                parts2 = matched_key.strip().split()
                if len(parts2) > 1:
                    matched_first_initial = parts2[0][0].lower()
            if not matched_first_initial or not first or first[0] == matched_first_initial:
                return name_index[sur]
    # 4. Partial match — both first and surname appear somewhere in any key.
    # For short surnames (≤2 chars) use word-boundary matching to avoid substring false-positives
    # (e.g. "Ho" appears inside "Thomas", "Johnson" etc.)
    def _word_in(word, text):
        if len(word) <= 2:
            return bool(re.search(r'(?<![a-z])' + re.escape(word) + r'(?![a-z])', text))
        return word in text
    candidates = []
    for key, rid in name_index.items():
        if first and sur and _word_in(first, key) and _word_in(sur, key):
            candidates.append((key, rid))
    if len(candidates) == 1:
        return candidates[0][1]
    if len(candidates) > 1:
        return candidates
    # 4b. For double-barrelled surnames, try each significant part.
    # Skip common name particles (de/del/van/etc.) — they appear in too many names
    # and cause false positives when matched as substrings.
    _NAME_PARTICLES = {
        "de", "del", "da", "di", "van", "von", "el", "al", "bin", "bint",
        "le", "la", "du", "des", "den", "der", "los", "las", "af", "av",
        "op", "ter", "ten", "ap",
    }
    sur_parts = sur.split()
    if len(sur_parts) > 1:
        for sp in sur_parts:
            if len(sp) < 3 or sp in _NAME_PARTICLES:
                continue
            # Use word-boundary matching so "andrés" doesn't match "guillermo andrés del río"
            # via substring — require the surname part to appear as a standalone word.
            hits = [
                (k, v) for k, v in name_index.items()
                if re.search(r'(?<![a-z])' + re.escape(first) + r'(?![a-z])', k)
                and re.search(r'(?<![a-z])' + re.escape(sp)   + r'(?![a-z])', k)
            ]
            if len(hits) == 1:
                return hits[0][1]

    # 4b2. Arabic "Al/El/Bin" prefix — try splitting/joining prefix forms
    # e.g. "Alsuhaymi" → "Al Suhaymi", "Al-Suhaymi"; "Al Suhaymi" → "Alsuhaymi"
    _AL_PREFIXES = ("al", "el", "bin", "bint", "abu", "abd")
    for pfx in _AL_PREFIXES:
        if sur.startswith(pfx) and len(sur) > len(pfx) + 1 and not sur[len(pfx)].isspace():
            # Try split: "alsuhaymi" → "al suhaymi" (with space) or "al-suhaymi" (with hyphen)
            alt_sur_split = pfx + " " + sur[len(pfx):]
            alt_sur_hyph  = pfx + "-" + sur[len(pfx):]
            for _alt_sur in (alt_sur_split, alt_sur_hyph):
                hits = [(k, v) for k, v in name_index.items()
                        if _word_in(first, k) and _alt_sur in k]
                if len(hits) == 1:
                    print(f"  ~ Al-prefix split match: '{full_name}' → '{hits[0][0]}'")
                    return hits[0][1]
        elif " " in sur:
            # Try joining: "al suhaymi" → "alsuhaymi"
            alt_sur_joined = sur.replace(" ", "")
            hits = [(k, v) for k, v in name_index.items()
                    if _word_in(first, k) and alt_sur_joined in k]
            if len(hits) == 1:
                print(f"  ~ Prefix-join match: '{full_name}' → '{hits[0][0]}'")
                return hits[0][1]
            break  # only try once per surname

    # 4c-pre. For 3+ word names, try truncated prefixes — handles the case where the
    # email name has an extra surname component not stored in Tracker
    # e.g. "Abdena Misgana MORKA" → try "Abdena Misgana" against the index.
    if len(parts) >= 3:
        for _trunc_len in range(len(parts) - 1, 1, -1):
            _trunc_key = " ".join(p.lower() for p in parts[:_trunc_len])
            if _trunc_key in name_index:
                _hit = name_index[_trunc_key]
                if isinstance(_hit, list):
                    print(f"  ~ Truncated-name match: '{full_name}' → '{_trunc_key}' ({len(_hit)} profiles)")
                    return [(_trunc_key, rid) for rid in _hit]
                print(f"  ~ Truncated-name match: '{full_name}' → '{_trunc_key}' (ID {_hit})")
                return _hit

    # 4c. For 3+ word names, try every possible first/surname split.
    # e.g. "Jeliana Pardo Jiménez" → try ("Jeliana","Pardo Jiménez") and ("Jeliana Pardo","Jiménez")
    import unicodedata as _ud2
    def _sa2(s):
        return _ud2.normalize("NFD", s).encode("ascii", "ignore").decode("ascii").lower()
    if len(parts) >= 3:
        norm_idx = {_sa2(k): v for k, v in name_index.items()}
        for split_at in range(1, len(parts)):
            alt_first = _sa2(" ".join(parts[:split_at]))
            alt_sur   = _sa2(" ".join(parts[split_at:]))
            for a, b in [(alt_first, alt_sur), (alt_sur, alt_first)]:
                alt_key = f"{a} {b}".strip()
                if alt_key in norm_idx:
                    print(f"  ~ Split-name match: '{full_name}' → '{alt_key}'")
                    return norm_idx[alt_key]
                # partial: both parts appear somewhere in the same key
                cands2 = [(k, v) for k, v in name_index.items()
                          if a in _sa2(k) and b in _sa2(k)]
                if len(cands2) == 1:
                    print(f"  ~ Split-name partial: '{full_name}' → '{cands2[0][0]}'")
                    return cands2[0][1]

    # 4. Live API scan — the /Resource/Search endpoint does not reliably filter by
    # firstname/surname. The only thing that always works is the unfiltered paginated
    # scan (same approach used to build the index). Scan the most recent 40 pages
    # (2000 records) so we catch candidates added since the cache was built, plus
    # try every possible first/surname split so multi-word names like "Shafeeq Ur
    # Rahman" are found regardless of how Tracker split the name on registration.
    if jwt:
        try:
            import unicodedata, difflib as _dl

            def _norm(s):
                s = unicodedata.normalize("NFD", s).encode("ascii","ignore").decode("ascii")
                return re.sub(r"[^a-z0-9 ]", "", s.lower()).strip()

            norm_target = _norm(full_lc)
            name_parts_lc = [_norm(p) for p in parts]  # accent-normalised

            def _matches(rf, rs):
                """True if the Tracker record name is a plausible match for full_name."""
                combined = f"{rf} {rs}".strip()
                combined_rev = f"{rs} {rf}".strip()
                norm_combined = _norm(combined)
                norm_rev      = _norm(combined_rev)
                # All words from the target name must appear in the record name
                if all(w in norm_combined or w in norm_rev for w in name_parts_lc if len(w) > 1):
                    return True
                # For 3+ word names: accept if all-but-one trailing words match
                # (e.g. "Abdena Misgana MORKA" → "Abdena Misgana" — MORKA is extra)
                if len(name_parts_lc) >= 3:
                    if all(w in norm_combined or w in norm_rev
                           for w in name_parts_lc[:-1] if len(w) > 1):
                        return True
                # Fuzzy similarity — but require at least 1 word from target to appear
                # in the record name (prevents completely different names matching on short strings)
                _has_word_overlap = any(
                    w in norm_combined or w in norm_rev
                    for w in name_parts_lc if len(w) > 2
                )
                if _has_word_overlap:
                    # First name must be similar to prevent e.g. Suraj/Saroj false positives
                    _target_first = name_parts_lc[0] if name_parts_lc else ""
                    _record_first = _norm(rf).split()[0] if rf else ""
                    if _target_first and _record_first:
                        _first_sim = _dl.SequenceMatcher(None, _target_first, _record_first).ratio()
                        if _first_sim < 0.75:
                            return False
                    if _dl.SequenceMatcher(None, norm_target, norm_combined).ratio() >= 0.82:
                        return True
                    if _dl.SequenceMatcher(None, norm_target, norm_rev).ratio() >= 0.82:
                        return True
                return False

            # Scan a wide range of pages to catch candidates missed by the index.
            # Cover pages 1-10 (oldest records) and a large window around the
            # estimated end of the index (most recent records).
            unique_ids = len(set(v for v in name_index.values() if isinstance(v, int)))
            estimated_pages = max(10, unique_ids // 50)
            pages_to_scan = list(range(1, 11))  # first 10 pages (oldest records)
            # Also probe pages at 25%, 50%, 75% of the database
            # (candidates registered months ago sit in the middle)
            for pct in [0.25, 0.5, 0.75]:
                mid = max(1, int(estimated_pages * pct))
                for p in range(mid - 2, mid + 3):
                    if p not in pages_to_scan:
                        pages_to_scan.append(p)
            for p in range(max(1, estimated_pages - 5), estimated_pages + 15):
                if p not in pages_to_scan:
                    pages_to_scan.append(p)

            found_id = None
            for page in pages_to_scan:
                try:
                    r = requests.post(
                        f"{TRACKER_API}/api/v1/Resource/Search",
                        json={"pageSize": 50, "pageNumber": page},
                        headers=h(jwt), timeout=15
                    )
                    if r.status_code != 200:
                        continue
                    try:
                        _rj = r.json()
                    except Exception:
                        continue
                    if isinstance(_rj, list):
                        items = _rj
                    elif isinstance(_rj, dict):
                        items = (_rj.get("items") or _rj.get("value") or
                                 _rj.get("results") or _rj.get("data") or [])
                    else:
                        items = []
                    if not items:
                        continue
                    for rec in items:
                        rf  = (rec.get("firstname") or rec.get("firstName") or "").strip().lower()
                        rs  = (rec.get("surname") or rec.get("lastName") or "").strip().lower()
                        if not rf and not rs:
                            _fn_s4 = (rec.get("name") or rec.get("resourceName") or
                                      rec.get("fullName") or "").strip()
                            if _fn_s4:
                                _fp_s4 = _fn_s4.split()
                                rf = _fp_s4[0].lower() if _fp_s4 else ""
                                rs = " ".join(_fp_s4[1:]).lower() if len(_fp_s4) > 1 else ""
                        rid = rec.get("resourceId")
                        if rid and _matches(rf, rs):
                            # Add to index so future runs are instant
                            key = f"{rf} {rs}".strip()
                            name_index[key] = rid
                            print(f"  ~ Found via live scan (page {page}): '{full_name}' → '{rf} {rs}' (ID {rid})")
                            found_id = rid
                            break
                    if found_id:
                        break
                except Exception:
                    continue

            if found_id:
                return found_id

        except Exception as e:
            pass  # fall through to fuzzy match

    # 5. Fuzzy match against full name index (handles apostrophes, hyphens, etc.)
    if sur and len(sur) <= 3:
        first_hits = [(k, v) for k, v in name_index.items()
                      if k.startswith(first + " ") and " " in k]
        if len(first_hits) == 1:
            print(f"  ~ First-name match (initials surname): '{full_name}' → '{first_hits[0][0]}'")
            return first_hits[0][1]

    try:
        import difflib
        def normalise(s):
            import unicodedata
            s = unicodedata.normalize("NFD", s).encode("ascii", "ignore").decode("ascii")
            return re.sub(r"[^a-z0-9 ]", "", s.lower()).strip()
        norm_target = normalise(full_lc)
        norm_index  = {normalise(k): (k, v) for k, v in name_index.items()}
        close = difflib.get_close_matches(norm_target, norm_index.keys(), n=3, cutoff=0.82)

        def _fuzzy_ok(search_norm, match_norm):
            """Guard against false positives: require first-name similarity AND last-word similarity."""
            s_words = search_norm.split()
            m_words = match_norm.split()
            if not s_words or not m_words:
                return True
            # First initial must match
            if s_words[0] and m_words[0] and s_words[0][0] != m_words[0][0]:
                return False
            # First name must have similarity ≥ 0.75 (prevents Suraj/Saroj false positives)
            if difflib.SequenceMatcher(None, s_words[0], m_words[0]).ratio() < 0.75:
                return False
            # Last word (surname) must have similarity ≥ 0.5
            s_last = s_words[-1]
            m_last = m_words[-1]
            if difflib.SequenceMatcher(None, s_last, m_last).ratio() < 0.5:
                return False
            return True

        valid = [c for c in close if _fuzzy_ok(norm_target, c)]
        if len(valid) == 1:
            orig_key, rid = norm_index[valid[0]]
            print(f"  ~ Fuzzy match: '{full_name}' → '{orig_key}' (ID {rid})")
            return rid
        if len(valid) > 1:
            return [(norm_index[c][0], norm_index[c][1]) for c in valid]
    except Exception:
        pass

    # 5b. Direct per-name API search — catches candidates the index missed due to
    # Tracker returning duplicate pages during the initial build. Try searching by
    # first name only, then surname only, then full name as a keyword.
    if jwt:
        try:
            import unicodedata as _ud4, difflib as _dl4
            def _nd4(s):
                return re.sub(r"[^a-z0-9 ]", "",
                              _ud4.normalize("NFD", s).encode("ascii","ignore").decode("ascii").lower()).strip()
            norm_target4 = _nd4(full_lc)
            name_parts4  = [_nd4(p) for p in parts]

            def _matches4(rf, rs):
                combined = f"{_nd4(rf)} {_nd4(rs)}".strip()
                combined_rev = f"{_nd4(rs)} {_nd4(rf)}".strip()
                if all(w in combined or w in combined_rev for w in name_parts4 if len(w) > 1):
                    return True
                # Accept initial-surname: Tracker has "Firstname Initial" for "Firstname Fullsurname"
                # e.g. index has "Gokulan S", search is "Gokulan Sivakumar"
                _rf_n = _nd4(rf); _rs_n = _nd4(rs)
                if (name_parts4 and len(name_parts4) >= 2 and len(_rs_n) == 1
                        and _rf_n.startswith(name_parts4[0])
                        and _rs_n == name_parts4[-1][0]):
                    return True
                return False

            # Build search payloads — most specific first to avoid false matches
            _direct_payloads = []
            # 1. Combined firstName + surname (most targeted — cuts through large result sets)
            if first and sur:
                # Use first word of first name and last word of surname for the combined search
                _first_word = parts[0].title()
                _last_word  = parts[-1].title()
                _direct_payloads.append({"pageSize": 50, "pageNumber": 1,
                                         "firstName": _first_word, "surname": _last_word})
                # Also try with full multi-word surname if different from last word
                if " " in sur:
                    _direct_payloads.append({"pageSize": 50, "pageNumber": 1,
                                             "firstName": _first_word, "surname": sur.title()})
            # 2. Surname only — usually more unique than first name
            if sur and len(sur) > 2:
                for _pg in [1, 2]:  # try 2 pages for surname (catches offset results)
                    _direct_payloads.append({"pageSize": 50, "pageNumber": _pg,
                                             "surname": sur.title()})
                    # Also try last word only if surname is multi-word
                    if " " in sur:
                        _direct_payloads.append({"pageSize": 50, "pageNumber": _pg,
                                                 "surname": parts[-1].title()})
            # 3. First name only
            if first:
                _direct_payloads.append({"pageSize": 50, "pageNumber": 1,
                                         "firstName": first.title()})
            # 4. Keyword / searchTerm (full name string) — try both field names
            _direct_payloads.append({"pageSize": 50, "pageNumber": 1, "keyword": full_name})
            _direct_payloads.append({"pageSize": 50, "pageNumber": 1, "searchTerm": full_name})
            # 5. Accent-stripped keyword
            _kw_stripped = re.sub(r"[^a-zA-Z0-9 ]", "",
                                   _ud4.normalize("NFD", full_name).encode("ascii","ignore").decode("ascii")).strip()
            if _kw_stripped and _kw_stripped.lower() != full_name.lower():
                _direct_payloads.append({"pageSize": 50, "pageNumber": 1, "keyword": _kw_stripped})
            # 6. Typo cleanup — strip trailing non-letter from last word (e.g. "Oliveiraq" → "Oliveira")
            _last_clean = re.sub(r'[^a-zA-Z]$', '', parts[-1])
            if _last_clean != parts[-1] and len(_last_clean) > 2:
                _direct_payloads.append({"pageSize": 50, "pageNumber": 1,
                                         "surname": _last_clean.title()})

            for _payload in _direct_payloads:
                try:
                    _r = requests.post(f"{TRACKER_API}/api/v1/Resource/Search",
                                       json=_payload, headers=h(jwt), timeout=15)
                    if _r.status_code != 200:
                        continue
                    try:
                        _rjj = _r.json()
                    except Exception:
                        continue
                    if isinstance(_rjj, list):
                        _items = _rjj
                    elif isinstance(_rjj, dict):
                        _items = (_rjj.get("items") or _rjj.get("value") or
                                  _rjj.get("results") or _rjj.get("data") or [])
                    else:
                        _items = []
                    for _rec in _items:
                        _rf  = ((_rec.get("firstname") or _rec.get("firstName") or "")).strip().lower()
                        _rs  = (_rec.get("surname") or _rec.get("lastName") or "").strip().lower()
                        if not _rf and not _rs:
                            _fn_5b = (_rec.get("name") or _rec.get("resourceName") or
                                      _rec.get("fullName") or "").strip()
                            if _fn_5b:
                                _fp_5b = _fn_5b.split()
                                _rf = _fp_5b[0].lower() if _fp_5b else ""
                                _rs = " ".join(_fp_5b[1:]).lower() if len(_fp_5b) > 1 else ""
                        _rid = _rec.get("resourceId")
                        if _rid and _matches4(_rf, _rs):
                            _key = f"{_rf} {_rs}".strip()
                            name_index[_key] = _rid
                            print(f"  ~ Found via direct search: '{full_name}' → '{_rf} {_rs}' (ID {_rid})")
                            return _rid
                except Exception:
                    continue
        except Exception:
            pass

    # 5c. Bin-name variant — try without "bin/binti" particle
    # e.g. "Mohammad Hazreen Izat bin Hashim" → search "Mohammad Hazreen Izat Hashim"
    if _bin_alt and _bin_alt.lower() != full_name.lower():
        print(f"  ~ Trying bin-name variant: '{_bin_alt}'")
        _bin_result = find_candidate(_bin_alt, name_index, jwt=jwt)
        if _bin_result:
            return _bin_result

    # 5d. Accent-stripped search — try the fully normalised (no accents) form
    import unicodedata as _ucdn2
    _accent_stripped = re.sub(r"[^a-z0-9 \-]", "",
                              _ucdn2.normalize("NFD", full_lc).encode("ascii", "ignore").decode("ascii")).strip()
    if _accent_stripped and _accent_stripped != full_lc:
        # Check index directly with accent-stripped form
        if _accent_stripped in name_index:
            print(f"  ~ Accent-stripped match: '{full_name}'")
            return name_index[_accent_stripped]
        # Also try reversed
        _as_parts = _accent_stripped.split()
        if len(_as_parts) >= 2:
            _as_rev = " ".join(_as_parts[1:]) + " " + _as_parts[0]
            if _as_rev in name_index:
                print(f"  ~ Accent-stripped reverse match: '{full_name}'")
                return name_index[_as_rev]

    # 5e-alt. Extended endpoint search — try PagedSearch, searchTerm field, and
    # Contact/Search (Jobs+ candidates may be stored as Contacts rather than Resources).
    if jwt:
        try:
            import unicodedata as _udja, difflib as _dlja
            def _nja(s):
                return re.sub(r"[^a-z0-9 ]", "",
                              _udja.normalize("NFD", s).encode("ascii","ignore").decode("ascii").lower()).strip()
            _tja = _nja(full_lc)
            _pja = [p for p in [_nja(x) for x in parts] if len(p) > 1]

            def _matches_ja(rf, rs, rname=""):
                if not rf.strip() and not rs.strip() and rname.strip():
                    _rp = rname.split()
                    rf = _rp[0].lower() if _rp else ""
                    rs = " ".join(_rp[1:]).lower() if len(_rp) > 1 else ""
                combined = _nja(f"{rf} {rs}")
                combined_r = _nja(f"{rs} {rf}")
                if _pja and all(w in combined or w in combined_r for w in _pja):
                    return True
                return max(
                    _dlja.SequenceMatcher(None, _tja, combined).ratio(),
                    _dlja.SequenceMatcher(None, _tja, combined_r).ratio()
                ) >= 0.80

            _first_w = parts[0].title() if parts else ""
            _last_w  = parts[-1].title() if parts else ""

            _ja_endpoints = [
                # Resource/PagedSearch — may honour filters unlike Resource/Search
                (f"{TRACKER_API}/api/v1/Resource/PagedSearch",
                 {"pageSize": 50, "pageNumber": 1, "searchTerm": full_name}),
                (f"{TRACKER_API}/api/v1/Resource/PagedSearch",
                 {"pageSize": 50, "pageNumber": 1, "keyword": full_name}),
                # searchTerm field (used by Client/Lead/Invoice search — may work on Resource too)
                (f"{TRACKER_API}/api/v1/Resource/Search",
                 {"pageSize": 50, "pageNumber": 1, "searchTerm": full_name}),
                # Contact/Search — Jobs+ applicants may be stored as Contact records
                (f"{TRACKER_API}/api/v1/Contact/Search",
                 {"pageSize": 50, "pageNumber": 1,
                  "firstName": _first_w, "lastName": _last_w}),
                (f"{TRACKER_API}/api/v1/Contact/Search",
                 {"pageSize": 50, "pageNumber": 1, "searchTerm": full_name}),
                (f"{TRACKER_API}/api/v1/Contact/PagedSearch",
                 {"pageSize": 50, "pageNumber": 1, "searchTerm": full_name}),
            ]

            for _url_ja, _body_ja in _ja_endpoints:
                try:
                    _rja = requests.post(_url_ja, json=_body_ja, headers=h(jwt), timeout=(5, 10))
                    if _rja.status_code not in (200, 201):
                        continue
                    _raw_ja = _rja.json()
                    # PagedSearch returns {items: [...]} rather than a plain list
                    _items_ja = (_raw_ja if isinstance(_raw_ja, list)
                                 else _raw_ja.get("items") or _raw_ja.get("results") or [])
                    for _rec_ja in _items_ja:
                        _rf_ja = (_rec_ja.get("firstname") or _rec_ja.get("firstName") or "").strip()
                        _rs_ja = (_rec_ja.get("surname") or _rec_ja.get("lastName") or "").strip()
                        _rn_ja = (_rec_ja.get("name") or _rec_ja.get("resourceName") or
                                  _rec_ja.get("fullName") or "").strip()
                        _rid_ja = (_rec_ja.get("resourceId") or _rec_ja.get("contactId") or
                                   _rec_ja.get("id"))
                        if _rid_ja and _matches_ja(_rf_ja, _rs_ja, _rn_ja):
                            _key_ja = f"{_rf_ja.lower()} {_rs_ja.lower()}".strip()
                            if _key_ja:
                                name_index[_key_ja] = _rid_ja
                            print(f"  ~ Found via extended search ({_url_ja.split('/')[-1]}): "
                                  f"'{full_name}' → '{_rf_ja} {_rs_ja}' (ID {_rid_ja})")
                            return _rid_ja
                except Exception:
                    continue
        except Exception:
            pass

    # 5e. EXHAUSTIVE FALLBACK — scan every Tracker page until the candidate is found.
    # Triggered only when all prior methods fail. Since every Aero Professional
    # registration email guarantees the candidate exists in Tracker under exactly that
    # name, this scan MUST find them. As a side-effect, all records encountered are
    # added to name_index and the cache is refreshed on disk so subsequent not-found
    # candidates hit the enriched index and skip this step entirely.
    if jwt:
        try:
            import unicodedata as _udex, difflib as _dlex, json as _jex, time as _tex
            def _nex(s):
                return re.sub(r"[^a-z0-9 ]", "",
                              _udex.normalize("NFD", s).encode("ascii","ignore").decode("ascii").lower()).strip()
            _target_ex  = _nex(full_lc)
            _parts_ex   = [p for p in [_nex(x) for x in parts] if len(p) > 1]

            def _matches_ex(rf, rs, rname=""):
                """Match a Tracker API record against the target name."""
                # Fall back to combined name field if firstName/surname both empty
                if not rf.strip() and not rs.strip() and rname.strip():
                    _rp = rname.split()
                    rf = _rp[0].lower() if _rp else ""
                    rs = " ".join(_rp[1:]).lower() if len(_rp) > 1 else ""
                combined   = _nex(f"{rf} {rs}")
                combined_r = _nex(f"{rs} {rf}")
                if _parts_ex and all(w in combined or w in combined_r for w in _parts_ex):
                    return True
                # Fuzzy fallback (≥0.85 similarity)
                return max(
                    _dlex.SequenceMatcher(None, _target_ex, combined).ratio(),
                    _dlex.SequenceMatcher(None, _target_ex, combined_r).ratio()
                ) >= 0.80

            _unique_ex  = len(set(v for v in name_index.values() if isinstance(v, int)))
            _max_ex     = max(200, (_unique_ex // 50) + 200)  # generous ceiling
            print(f"  ⚑ Exhaustive Tracker scan for '{full_name}' — up to {_max_ex} pages …")

            _found_ex      = None
            _new_ex        = 0
            _empty_ex      = 0
            _dup_streak_ex = 0
            _cpath_ex = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                     "tracker_cache.json")

            def _save_index_ex():
                """Save enriched index to disk mid-scan so progress survives force-quit."""
                try:
                    with open(_cpath_ex, "r", encoding="utf-8") as _cf:
                        _c = _jex.load(_cf)
                    _c["names"] = name_index
                    _c["ts"] = __import__("time").time()
                    with open(_cpath_ex, "w", encoding="utf-8") as _cf:
                        _jex.dump(_c, _cf, ensure_ascii=False)
                except Exception:
                    pass

            for _pg_ex in range(1, _max_ex + 1):
                try:
                    _rex = requests.post(
                        f"{TRACKER_API}/api/v1/Resource/Search",
                        json={"pageSize": 50, "pageNumber": _pg_ex},
                        headers=h(jwt), timeout=15
                    )
                    if _rex.status_code != 200:
                        continue
                    try:
                        _raw_ex = _rex.json()
                    except Exception:
                        continue
                    # Handle both list and dict (items/value) responses
                    if isinstance(_raw_ex, list):
                        _its_ex = _raw_ex
                    elif isinstance(_raw_ex, dict):
                        _its_ex = (_raw_ex.get("items") or _raw_ex.get("value") or
                                   _raw_ex.get("results") or _raw_ex.get("data") or [])
                    else:
                        _its_ex = []
                    if not _its_ex:
                        _empty_ex += 1
                        if _empty_ex >= 3:
                            print(f"  ⚑ Exhaustive scan: no more pages after {_pg_ex}.")
                            break
                        continue
                    _empty_ex  = 0
                    _page_new  = 0
                    for _re_ex in _its_ex:
                        _rf_ex  = (_re_ex.get("firstname") or _re_ex.get("firstName") or "").strip()
                        _rs_ex  = (_re_ex.get("surname") or _re_ex.get("lastName") or "").strip()
                        _rn_ex  = (_re_ex.get("name") or _re_ex.get("resourceName") or
                                   _re_ex.get("fullName") or "").strip()
                        _rid_ex = _re_ex.get("resourceId")
                        if _rid_ex:
                            _key_ex = f"{_rf_ex.lower()} {_rs_ex.lower()}".strip()
                            if _key_ex and _key_ex not in name_index:
                                name_index[_key_ex] = _rid_ex
                                _nex_key = _nex(f"{_rf_ex} {_rs_ex}")
                                if _nex_key and _nex_key not in name_index:
                                    name_index[_nex_key] = _rid_ex
                                _new_ex   += 1
                                _page_new += 1
                            # Check if this is the candidate we're looking for
                            if _matches_ex(_rf_ex, _rs_ex, _rn_ex):
                                print(f"  ⚑ FOUND (page {_pg_ex}): '{full_name}' → "
                                      f"'{_rf_ex} {_rs_ex}' (ID {_rid_ex})")
                                _found_ex = _rid_ex
                                break
                    if _found_ex:
                        break
                    # Track pages that add nothing new (API wrap-around detection)
                    if _page_new == 0:
                        _dup_streak_ex += 1
                    else:
                        _dup_streak_ex = 0
                    # Stop early if index is already large — API is returning duplicates.
                    # 15 consecutive empty pages = we've covered everything.
                    _ex_stop = 15 if len(name_index) > 5000 else 100
                    if _dup_streak_ex >= _ex_stop:
                        print(f"  ⚑ Exhaustive scan: {_dup_streak_ex} all-duplicate pages "
                              f"— stopping (index has {len(name_index)} entries).")
                        break
                    if _pg_ex % 25 == 0:
                        print(f"  ⚑ … page {_pg_ex}/{_max_ex}, "
                              f"+{_new_ex} new records indexed …")
                        # Save progress every 25 pages so force-quit doesn't lose work
                        if _new_ex > 0:
                            _save_index_ex()
                    _tex.sleep(0.01)
                except Exception:
                    continue

            # Persist the enriched index so subsequent candidates skip this scan
            if _new_ex > 0:
                print(f"  ⚑ Exhaustive scan complete — {_new_ex} new records added to index.")
                try:
                    _cpath = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                          "tracker_cache.json")
                    with open(_cpath, "r", encoding="utf-8") as _cf:
                        _cache_ex = _jex.load(_cf)
                    _cache_ex["names"] = name_index
                    _cache_ex["ts"] = __import__("time").time()
                    with open(_cpath, "w", encoding="utf-8") as _cf:
                        _jex.dump(_cache_ex, _cf, ensure_ascii=False)
                    print(f"  ⚑ Enriched index saved ({len(name_index)} total entries).")
                except Exception as _e_save:
                    print(f"  ⚑ Warning: could not save enriched index: {_e_save}")
            elif not _found_ex:
                print(f"  ⚑ Exhaustive scan: all pages already indexed — "
                      f"no new records found for '{full_name}'.")

            if _found_ex:
                return _found_ex

        except Exception as _e_ex:
            print(f"  ⚑ Exhaustive scan error: {_e_ex}")

    # 6. Debug fallback — log closest index entries to file so we can diagnose misses
    try:
        import difflib as _df2, unicodedata as _ud3
        def _nd(s):
            return re.sub(r"[^a-z0-9 ]", "",
                          _ud3.normalize("NFD", s).encode("ascii","ignore").decode("ascii").lower()).strip()
        _nt2 = _nd(full_lc)
        _ni2 = {_nd(k): k for k in name_index if isinstance(name_index[k], int)}
        top3 = _df2.get_close_matches(_nt2, _ni2.keys(), n=5, cutoff=0.4)
        _closest = [_ni2[c] for c in top3] if top3 else []
        if _closest:
            print(f"  ⚠ NOT FOUND. Closest in index: {_closest}")
        else:
            print(f"  ⚠ NOT FOUND. No close matches in index.")

        # Subset acceptance: if all words of the shorter name appear in the longer,
        # and the first names are prefix-compatible, treat as a match.
        # Handles: "Ramon Aymami" → "Ramon Aymami Perez", "Gokulan Sivakumar" → "Gokulan S",
        #          "Lillian Pokua Quansah" → "Lilian Quansah", "Christiaan Du Plessis" → "Chris Du Plessis"
        for _cm_k in top3[:3]:
            _orig_cm  = _ni2[_cm_k]
            _cm_id    = name_index.get(_orig_cm)
            if not isinstance(_cm_id, int):
                continue
            _ratio_cm = _df2.SequenceMatcher(None, _nt2, _cm_k).ratio()
            if _ratio_cm < 0.60:
                continue
            # Determine which is shorter (search name vs index key, both accent-stripped)
            if len(_nt2) <= len(_cm_k):
                _short_w, _long_s = _nt2.split(), _cm_k
            else:
                _short_w, _long_s = _cm_k.split(), _nt2
            # Filter to words longer than 1 char (ignore particles/initials in short name)
            _sig_cm = [w for w in _short_w if len(w) > 1]
            if not _sig_cm:
                continue
            # First-name must be prefix-compatible (one starts with the other)
            _fn_short = _sig_cm[0]
            _fn_long  = _long_s.split()[0] if _long_s.split() else ""
            _fn_ok    = (_fn_short.startswith(_fn_long) or _fn_long.startswith(_fn_short))
            if not _fn_ok:
                continue
            # All significant short words must appear as substrings in the long name
            if all(w in _long_s for w in _sig_cm):
                print(f"  ~ Subset name match: \'{full_name}\' → \'{_orig_cm}\'")
                return _cm_id

        # Write to not_found_debug.txt so we can see what Tracker has for each person
        _debug_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "not_found_debug.txt")
        with open(_debug_path, "a", encoding="utf-8") as _dbf:
            _dbf.write(f"SEARCH: {full_name}\n")
            if _closest:
                _dbf.write(f"  Closest in index: {_closest}\n")
            else:
                _dbf.write(f"  No close matches in index\n")
            # Also record what the direct API keyword search returned
            if jwt:
                try:
                    _kr = requests.post(f"{TRACKER_API}/api/v1/Resource/Search",
                                        json={"pageSize": 5, "pageNumber": 1, "keyword": full_name},
                                        headers=h(jwt), timeout=10)
                    _kitems = _kr.json() if _kr.status_code == 200 and isinstance(_kr.json(), list) else []
                    if _kitems:
                        _dbf.write(f"  Keyword API results:\n")
                        for _ki in _kitems[:5]:
                            _kfn = _ki.get("firstname") or _ki.get("firstName") or ""
                            _ksn = _ki.get("surname") or _ki.get("lastName") or ""
                            _knn = _ki.get("name") or _ki.get("resourceName") or ""
                            _kid = _ki.get("resourceId") or ""
                            _dbf.write(f"    ID {_kid} | first='{_kfn}' sur='{_ksn}' name='{_knn}'\n")
                    else:
                        _dbf.write(f"  Keyword API: 0 results\n")
                except Exception:
                    pass
            _dbf.write("\n")
    except Exception:
        pass
    return None

# ── Employer validation (module-level, used to validate stored Tracker values) ─

_SECTION_HEADER_SIMPLE_RE = re.compile(
    r"^(aviation career|work experience|employment history|career history|"
    r"career summary|professional experience|professional background|"
    r"professional history|work history|job history|career profile|"
    r"career overview|aviation experience|flying experience|"
    r"experience|employment|career)\s*$",
    re.IGNORECASE
)
_EMPLOYER_ACTION_SIMPLE_RE = re.compile(
    r"\b(management|leadership|training|development|operations|strategy|"
    r"procurement|stakeholder|implementation|responsible|coordination|"
    r"improvement|sop|oversight|delivery|initiative|planning|maintenance)\b",
    re.IGNORECASE
)

def _clean_employer_simple(raw):
    """Quick validation of a stored employer string.  Returns '' if it looks like
    CV noise (section headers, date-only strings, job-description fragments)."""
    if not raw or not raw.strip():
        return ""
    raw = raw.strip()
    # Strip European date format and everything after
    raw = re.sub(r"\s*:?\s*\d{2}[./]\d{2}[./]\d{4}.*$", "", raw).strip()
    # Strip trailing "| 2020 – Present", "YYYY–present" etc.
    raw = re.sub(r"[\|–\-]\s*\d{4}.*$", "", raw).strip()
    raw = re.sub(r"\s*[\|–\-]\s*(present|current|now).*$", "", raw, flags=re.IGNORECASE).strip()
    raw = raw.rstrip(".,;|– ")
    if not raw or len(raw) < 3:
        return ""
    if _SECTION_HEADER_SIMPLE_RE.match(raw):
        return ""
    return raw


# Regex for recognised professional job title keywords
_JOB_TITLE_KEYWORDS_RE = re.compile(
    r"\b(captain|officer|pilot|engineer|maintenance|technician|mechanic|"
    r"attendant|crew|purser|manager|director|supervisor|coordinator|"
    r"controller|dispatcher|instructor|examiner|trainer|analyst|planner|"
    r"specialist|consultant|operator|inspector|administrator|assistant|"
    r"lead|head|chief|senior|junior|first|second|cabin|flight|ground|"
    r"ramp|cargo|safety|quality|compliance|technical|avionics|commercial|"
    r"certifying|base maintenance|line maintenance|overhaul|lame|ame\b)\b",
    re.IGNORECASE
)


def _is_plausible_job_title(title):
    """Returns True if 'title' looks like a genuine professional job title.
    Used to protect existing Tracker titles from being overwritten with garbage.
    """
    if not title:
        return False
    t = title.strip()
    words = t.split()
    if not words:
        return False
    # Too long = bullet point / description, not a title
    if len(words) > 8:
        return False
    # Starts with a digit or non-letter
    if not words[0][0].isalpha():
        return False
    # Starts with a preposition/article → sentence fragment, not a title
    _BAD_TITLE_STARTS = {"of", "in", "the", "a", "an", "for", "to", "at",
                         "with", "and", "or", "from", "by", "on", "as", "via"}
    if words[0].lower() in _BAD_TITLE_STARTS:
        return False
    # Contains a recognised job title keyword
    if _JOB_TITLE_KEYWORDS_RE.search(t):
        return True
    # Proper-noun pattern: at least half the words are capitalised
    cap_count = sum(1 for w in words if w and w[0].isupper())
    if len(words) >= 2 and cap_count / len(words) >= 0.5:
        return True
    # Single capitalised word of reasonable length
    if len(words) == 1 and t[0].isupper() and len(t) >= 5:
        return True
    return False


# Regex for strong company-name signals
_EMPLOYER_PLAUSIBILITY_RE = re.compile(
    r"\b(airline|airways|aviation|airport|cargo|express|group|"
    r"ltd\.?|llc|inc\.?|plc|corp\.?|company|co\b|"
    r"technic|tech\b|services?|solutions?|"
    r"international|global|consulting|enterprises?|authority|council|"
    r"logistics|transport|trading|industries|institute|"
    r"engineering|maintenance|overhaul|mro|bureau|agency|"
    r"university|college|school|academy|hospital|bank|"
    r"limited|incorporated|association|foundation)\b",
    re.IGNORECASE
)


_EMPLOYER_SINGULAR_ROLE_END_RE = re.compile(
    r"\b(engineer|captain|officer|pilot|technician|mechanic|specialist|"
    r"analyst|consultant|inspector|controller|dispatcher|coordinator|trainer|instructor|examiner|"
    r"manager|director|supervisor|administrator|operator|planner|scheduler|auditor)\s*$",
    re.IGNORECASE
)
_EMPLOYER_COMPANY_FORMATION_RE = re.compile(
    r"\b(ltd|llc|inc|plc|corp\b|company|co\b|group\b|authority|council|"
    r"services?|solutions?|industries|enterprises?|holdings|associates?|"
    r"partners?|foundation|limited|incorporated|association|institute)\b",
    re.IGNORECASE
)

def _is_plausible_employer(name):
    """Returns True if 'name' looks like a genuine employer/company name.
    Used as a second-line defence — if the regex extraction returns something
    implausible, we keep the existing Tracker value instead.
    """
    if not name:
        return False
    n = name.strip()
    words = n.split()
    if not words:
        return False
    # Sentence-fragment signal: first word starts lowercase → almost certainly not a company
    if words[0][0].islower():
        return False
    # Training/course content signal: starts with a word that indicates a module/course
    _BAD_STARTS = {"module", "course", "training", "certification", "level",
                   "phase", "unit", "chapter", "section", "part", "exam",
                   "theory", "practical", "lesson", "lecture"}
    if words[0].lower() in _BAD_STARTS:
        return False
    # Too long to be a company name (>7 words = sentence)
    if len(words) > 7:
        return False
    # Too many characters — real company names don't exceed ~50 characters
    if len(n) > 50:
        return False
    # Generic single-word aviation terms are not company names
    _GENERIC_AVIATION_TERMS = {"airlines", "airline", "airways", "aviation", "airport", "aeroline"}
    if len(words) == 1 and n.strip().lower() in _GENERIC_AVIATION_TERMS:
        return False
    # Job description signal: ends with a singular role noun and has no company formation word.
    # e.g. "Aviation & Aerospace Engineer", "Airline Captain" — role descriptions, not companies.
    if _EMPLOYER_SINGULAR_ROLE_END_RE.search(n) and not _EMPLOYER_COMPANY_FORMATION_RE.search(n):
        return False
    # Contains a strong company-type keyword → high confidence
    if _EMPLOYER_PLAUSIBILITY_RE.search(n):
        return True
    # Proper-noun test: most words start uppercase (handles "GMR AeroTechnic", "South African Airways")
    # Also handles ALL-CAPS company names like "SITSCO WLL", "AVIATOR SOLUTIONS LIMITED"
    cap_count = sum(1 for w in words if w and (w[0].isupper() or w.isupper()))
    if len(words) >= 2 and cap_count / len(words) >= 0.6:
        return True
    # Single capitalised word of reasonable length (e.g. "Emirates", "Ryanair", "Iberia")
    if len(words) == 1 and n[0].isupper() and len(n) >= 5:
        return True
    return False


# ── CV retrieval ───────────────────────────────────────────────────────────────

def get_cv_text(jwt, resource_id):
    """Try to download and extract text from the candidate's CV."""
    # List documents — handle both list and dict (items/documents/value) responses
    r = requests.get(f"{TRACKER_API}/api/v1/Resource/{resource_id}/Documents",
                     headers=h(jwt), timeout=(5, 10))
    if r.status_code != 200:
        return None
    try:
        _raw_docs = r.json()
    except Exception:
        return None

    if isinstance(_raw_docs, list):
        docs = _raw_docs
    elif isinstance(_raw_docs, dict):
        docs = (_raw_docs.get("items") or _raw_docs.get("documents") or
                _raw_docs.get("value") or _raw_docs.get("$values") or
                _raw_docs.get("results") or _raw_docs.get("data") or [])
    else:
        docs = []

    if not docs:
        return None

    # Documents exist — if all downloads fail we still must NOT email the candidate
    _DOCS_EXIST_SENTINEL = "CV_EXISTS_NO_TEXT"

    # Prefer primary / resume type; fall back to first document
    cv_doc = next((d for d in docs if d.get("isPrimary")), None)
    if not cv_doc:
        cv_doc = next((d for d in docs
                       if "resume" in (d.get("documentType") or "").lower()
                       or "cv" in (d.get("documentType") or "").lower()
                       or "cv" in (d.get("filename") or d.get("name") or "").lower()), None)
    if not cv_doc:
        cv_doc = docs[0]

    # Try every possible field name for the document ID
    doc_id = (cv_doc.get("documentId") or cv_doc.get("id") or
              cv_doc.get("docId") or cv_doc.get("resourceDocumentId") or
              cv_doc.get("DocumentId") or cv_doc.get("document_id") or
              cv_doc.get("documentID"))
    filename = (cv_doc.get("filename") or cv_doc.get("name") or
                cv_doc.get("fileName") or cv_doc.get("originalFilename") or "cv")
    ext = "." + filename.rsplit(".", 1)[-1].lower() if "." in filename else ".pdf"

    if not doc_id:
        # Log what fields the document object actually has so we can diagnose
        print(f"  ⚠  Document found but ID field missing. Keys: {list(cv_doc.keys())}")
        # Try direct URL from the document object itself
        direct_url = (cv_doc.get("url") or cv_doc.get("downloadUrl") or
                      cv_doc.get("fileUrl") or cv_doc.get("downloadLink"))
        if not direct_url:
            return _DOCS_EXIST_SENTINEL

    # Try download URL patterns.
    # Primary: REST API endpoint confirmed via swagger — JWT auth only, never expires.
    # Fallback: legacy web endpoint — requires TRACKER_WEB_COOKIE which may expire.
    WEB_BASE = "https://evouk.tracker-rms.com"

    _download_urls = []
    if doc_id:
        _download_urls = [
            # Swagger-confirmed REST API — returns JSON with base64 filecontent field
            f"{TRACKER_API}/api/v1/Resource/{resource_id}/Document/{doc_id}",
            # Legacy web fallback — only used if API endpoint fails
            f"{WEB_BASE}/Document/Download/{doc_id}",
        ]
    # Also try any direct URL embedded in the document object
    _obj_url = (cv_doc.get("url") or cv_doc.get("downloadUrl") or
                cv_doc.get("fileUrl") or cv_doc.get("downloadLink"))
    if _obj_url:
        _download_urls.insert(0, _obj_url)

    # Accepted file magic bytes: PDF, DOCX/ZIP, old .doc (OLE2), RTF,
    # JPEG (\xff\xd8\xff), PNG (\x89PNG), GIF, BMP, TIFF (LE/BE), WEBP
    _ACCEPTED_MAGIC = (b"%PDF", b"PK\x03\x04", b"\xd0\xcf\x11\xe0", b"{\x5crt")
    _IMAGE_MAGIC    = (b"\xff\xd8\xff",  # JPEG (any variant)
                       b"\x89PNG",       # PNG
                       b"GIF8",          # GIF
                       b"BM",            # BMP
                       b"RIFF",          # WEBP (RIFF....WEBP)
                       b"II*",       # TIFF little-endian
                       b"MM*")       # TIFF big-endian

    def _is_image(raw):
        """Return True if raw bytes look like a raster image."""
        magic = raw[:4]
        return (magic[:3] in (m[:3] for m in _IMAGE_MAGIC) or
                magic in _IMAGE_MAGIC or
                magic[:2] == b"BM")

    def _ocr_image_bytes(raw):
        """OCR a raw image (JPEG/PNG etc.) and return extracted text, or None."""
        try:
            from PIL import Image as _PILImage
            import pytesseract as _tess
            import io as _io
            import concurrent.futures as _cf
            _img = _PILImage.open(_io.BytesIO(raw))
            # Resize very large images to speed up OCR (cap at 2000px wide)
            if _img.width > 2000:
                _ratio = 2000 / _img.width
                _img = _img.resize((2000, int(_img.height * _ratio)))
            # Run OCR with a 20-second timeout to prevent hanging
            with _cf.ThreadPoolExecutor(max_workers=1) as _ex:
                _fut = _ex.submit(_tess.image_to_string, _img, lang="eng")
                try:
                    _text = _fut.result(timeout=10)
                except _cf.TimeoutError:
                    print("  ⚠  Image OCR timed out — skipping CV text extraction.")
                    return None
            if _text.strip():
                print("  ℹ  Extracted text via image OCR (pytesseract)")
                return _text
        except ImportError:
            pass
        except Exception:
            pass
        # Fallback: try fitz (pymupdf) — it can OCR some image formats too
        try:
            import fitz as _fz
            _doc = _fz.open(stream=raw, filetype="jpeg")
            _parts = [p.get_text() for p in _doc]
            _text = "\n".join(_parts)
            if _text.strip():
                return _text
        except Exception:
            pass
        return None

    def _try_extract(raw, file_ext):
        """Download bytes → text. Returns text (possibly empty), or None on hard failure."""
        text = extract_text(raw, file_ext)
        if text is None:
            return None
        if text.strip():
            return text
        return ""

    for url in _download_urls:
        try:
            headers = h(jwt)
            # Only send web cookie for web-based URLs — REST API uses JWT only
            if "evouk.tracker-rms.com" in url:
                headers["Cookie"] = TRACKER_WEB_COOKIE
            resp = requests.get(url, headers=headers, timeout=(5, 8), allow_redirects=True)
            ct = resp.headers.get("Content-Type", "")
            if "html" in ct.lower():
                continue

            # ── REST API returns JSON with base64-encoded filecontent ──────────
            # Decode it to raw bytes so the rest of the pipeline handles it normally.
            if resp.status_code == 200 and "json" in ct.lower():
                try:
                    import base64 as _b64mod
                    _jdata = resp.json()
                    _b64 = (_jdata.get("filecontent") or _jdata.get("fileContent")
                            or _jdata.get("content") or _jdata.get("fileData"))
                    if not _b64:
                        continue
                    _raw = _b64mod.b64decode(_b64)
                    # Use filename from API response for accurate extension
                    _api_fname = _jdata.get("filename") or _jdata.get("fileName")
                    if _api_fname and "." in _api_fname:
                        ext = "." + _api_fname.rsplit(".", 1)[-1].lower()
                    # Inject decoded bytes back as resp.content for unified processing
                    resp._content = _raw
                    ct = ""  # treat as binary from here
                except Exception:
                    continue

            if resp.status_code == 200 and "json" not in ct and len(resp.content) > 200:
                magic4 = resp.content[:4]

                # ── Image file (JPEG/PNG/etc.) — run image OCR ─────────────
                if _is_image(resp.content):
                    _img_text = _ocr_image_bytes(resp.content)
                    if _img_text and _img_text.strip():
                        return _img_text
                    # Image exists but OCR got nothing (blank page, logo, etc.)
                    return _DOCS_EXIST_SENTINEL

                # Accept known binary formats; also accept if mostly printable (plain text)
                _printable = sum(1 for b in resp.content[:200]
                                 if 0x20 <= b <= 0x7e or b in (0x09, 0x0a, 0x0d))
                if magic4 not in _ACCEPTED_MAGIC and _printable < 150:
                    # Unknown binary — document exists but unreadable format
                    print(f"  ℹ  Document downloaded ({len(resp.content)} bytes) "
                          f"but format unrecognised (magic: {resp.content[:4]}) — "
                          f"marking as exists.")
                    return _DOCS_EXIST_SENTINEL
                text = _try_extract(resp.content, ext)
                if text is None:
                    # Try OCR for scanned PDFs before giving up
                    if ext == ".pdf":
                        _ocr = _ocr_pdf(resp.content)
                        if _ocr and _ocr.strip():
                            return _ocr
                    continue
                if text.strip():
                    return text
                # Empty text — try OCR before returning sentinel
                if ext == ".pdf":
                    _ocr = _ocr_pdf(resp.content)
                    if _ocr and _ocr.strip():
                        return _ocr
                return _DOCS_EXIST_SENTINEL
            if resp.status_code in (301, 302, 307, 308):
                redirect_url = resp.headers.get("Location")
                if redirect_url:
                    resp2 = requests.get(redirect_url, timeout=(5, 8))
                    if resp2.status_code == 200 and len(resp2.content) > 200:
                        text = _try_extract(resp2.content, ext)
                        if text is not None:
                            return text if text.strip() else _DOCS_EXIST_SENTINEL
        except Exception:
            pass

    # All download attempts failed but documents exist — return sentinel so caller knows
    print(f"  ⚠  {len(docs)} document(s) found in Tracker but none could be downloaded "
          f"(doc_id={doc_id}, tried {len(_download_urls)} URLs). "
          f"Doc keys: {list(cv_doc.keys())}")
    return _DOCS_EXIST_SENTINEL


def _ocr_pdf(content):
    """OCR a scanned/image-only PDF. Returns text or None."""
    # Method 0: pymupdf (fitz) — often extracts from PDFs that pdfplumber/pdfminer miss
    # Install: pip install pymupdf --break-system-packages
    try:
        import fitz  # pymupdf
        doc = fitz.open(stream=content, filetype="pdf")
        parts = []
        for page in doc:
            parts.append(page.get_text())
        text = "\n".join(parts)
        if text.strip():
            print("  ℹ  Extracted text via pymupdf")
            return text
    except ImportError:
        pass
    except Exception:
        pass

    # Method 1: pymupdf render to image + pytesseract (no poppler needed)
    try:
        import fitz as _fitz_ocr, pytesseract as _tess_ocr
        from PIL import Image as _PILImg_ocr
        import io as _io_ocr
        import concurrent.futures as _cf_ocr
        _doc_ocr = _fitz_ocr.open(stream=content, filetype="pdf")
        _parts_ocr = []
        for _pn in range(min(len(_doc_ocr), 5)):  # max 5 pages
            _pix = _doc_ocr.load_page(_pn).get_pixmap(dpi=200)
            _img = _PILImg_ocr.open(_io_ocr.BytesIO(_pix.tobytes("png")))
            # Cap image width to 1800px
            if _img.width > 1800:
                _r = 1800 / _img.width
                _img = _img.resize((1800, int(_img.height * _r)))
            with _cf_ocr.ThreadPoolExecutor(max_workers=1) as _ex_ocr:
                _fut = _ex_ocr.submit(_tess_ocr.image_to_string, _img, lang="eng")
                try:
                    _parts_ocr.append(_fut.result(timeout=10))
                except _cf_ocr.TimeoutError:
                    print("  ⚠  PDF page OCR timed out — skipping page.")
                    break
        _text = "\n".join(_parts_ocr)
        if _text.strip():
            print("  ℹ  Extracted text via pymupdf+pytesseract OCR")
            return _text
    except (ImportError, Exception):
        pass

    return None


def extract_text(content, ext):
    try:
        if ext == ".pdf":
            import pdfplumber
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                text = "\n".join(p.extract_text() or "" for p in pdf.pages)
            if text.strip():
                return text

            # pdfplumber got nothing — try pdfminer (handles some cases pdfplumber misses)
            try:
                from pdfminer.high_level import extract_text as pdfminer_extract
                text = pdfminer_extract(io.BytesIO(content))
                if text and text.strip():
                    return text
            except Exception:
                pass

            # Still empty — scanned/image PDF, try OCR
            return _ocr_pdf(content) or ""
        elif ext == ".docx":
            from docx import Document
            doc = Document(io.BytesIO(content))
            return "\n".join(p.text for p in doc.paragraphs)
        elif ext == ".doc":
            # Old binary Word format — try win32com (Word must be installed)
            try:
                import win32com.client, tempfile as _tmp, os as _os
                _tf = _tmp.NamedTemporaryFile(suffix=".doc", delete=False)
                _tf.write(content); _tf.close()
                try:
                    _word = win32com.client.Dispatch("Word.Application")
                    _word.Visible = False
                    _d = _word.Documents.Open(_tf.name)
                    _text = _d.Content.Text
                    _d.Close(False)
                    _word.Quit()
                    return _text
                finally:
                    _os.unlink(_tf.name)
            except Exception:
                pass
            # Fallback: pull readable ASCII strings from the binary
            import re as _re
            chunks = _re.findall(rb"[\x20-\x7e]{5,}", content)
            return " ".join(c.decode("ascii", errors="ignore") for c in chunks)
        elif ext == ".rtf":
            text = content.decode("utf-8", errors="ignore")
            import re as _re
            text = _re.sub(r"\{[^{}]{0,60}\}", "", text)   # short control groups
            text = _re.sub(r"\\[a-z]+\-?\d*\s?", " ", text)  # control words
            text = _re.sub(r"[{}\\]", " ", text)
            return " ".join(text.split())
        elif ext in (".txt", ".text"):
            return content.decode("utf-8", errors="ignore")
        elif ext in (".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif", ".webp"):
            # Image file — run OCR directly
            try:
                from PIL import Image as _PILImg
                import pytesseract as _tess_et
                import io as _io_et, concurrent.futures as _cf_et
                _img = _PILImg.open(_io_et.BytesIO(content))
                if _img.width > 2000:
                    _r = 2000 / _img.width
                    _img = _img.resize((2000, int(_img.height * _r)))
                with _cf_et.ThreadPoolExecutor(max_workers=1) as _ex:
                    _fut = _ex.submit(_tess_et.image_to_string, _img, lang="eng")
                    try:
                        _t = _fut.result(timeout=10)
                        if _t and _t.strip():
                            return _t
                    except _cf_et.TimeoutError:
                        pass
            except Exception:
                pass
            return None
        else:
            # Unknown format — try UTF-8 first, then raw ASCII extraction
            decoded = content.decode("utf-8", errors="ignore").strip()
            if len(decoded) > 100:
                return decoded
            import re as _re
            chunks = _re.findall(rb"[\x20-\x7e]{5,}", content)
            return " ".join(c.decode("ascii", errors="ignore") for c in chunks)
    except Exception as e:
        print(f"  ⚠  CV text extraction failed ({ext}): {e}")
        return None

# ── AI CV parsing ──────────────────────────────────────────────────────────────

# ── Known aircraft types (from Tracker skills area 40) ────────────────────────
# Sorted longest-first so "B737-800" matches before "B737"
AIRCRAFT_TYPES = sorted([
    "A220","A300-600","A300-B4","A310","A318","A319","A320","A321","A330","A340",
    "A340-500","A340-600","A350","A380","ATR42","ATR72","B717","B727","B737",
    "B737-300","B737-400","B737-500","B737-700","B737-800","B737-900","B737NG",
    "B737MAX","B747","B747-400","B757","B767","B777","B777-200","B777-300",
    "B777X","B787","B787-8","B787-9","CRJ200","CRJ700","CRJ900","CRJ1000",
    "Dash8","DHC-8","E170","E175","E190","E195","ERJ135","ERJ145","ERJ170",
    "ERJ190","F100","MD11","MD80","MD90","Q400","Saab340","Saab2000",
    "C130","C17","C295","PC12","King Air","B200","B350","Caravan","C208",
    "EC135","EC145","EC155","EC175","EC225","AW139","AW169","AW189",
    "S76","S92","Bell206","Bell412","Bell429","R44","Robinson",
    # Embraer variants as they appear in CVs (mapped to Tracker names via AIRCRAFT_ALIASES below)
    "EMB170","EMB175","EMB190","EMB195",
    "ERJ-170","ERJ-175","ERJ-190","ERJ-195",
], key=len, reverse=True)

# Maps aircraft names as written in CVs → exact Tracker skill names
# Add entries here whenever you discover a mismatch between CV text and Tracker
def _norm_ac_text(t):
    """Normalise full aircraft names so B/A-short codes match: Boeing 737 → B737, Airbus A320 → A320."""
    import re as _re
    t = _re.sub(r"\bBoeing\s+(7\d\d)(?:\s*[-]?\s*(\d{3}|ER|LR|NGX?|MAX\d?|F|C|BBJ))?\b",
                lambda m: "B" + m.group(1) + ("-" + m.group(2).replace(" ","") if m.group(2) else ""),
                t, flags=_re.IGNORECASE)
    t = _re.sub(r"\bAirbus\s+A?(\d{3})(?:\s*[-]?\s*(\d{3}|neo|ceo|XWB|F|P|T))?\b",
                lambda m: "A" + m.group(1) + ("-" + m.group(2).replace(" ","") if m.group(2) else ""),
                t, flags=_re.IGNORECASE)
    # Also handle "737 NG", "737-NG" etc. when Boeing already stripped
    t = _re.sub(r"\b(B7\d\d)\s+NG\b", r"\1-NG", t, flags=_re.IGNORECASE)
    return t

AIRCRAFT_ALIASES = {
    "e175":    "EMB175",
    "e170":    "EMB170",
    "e190":    "EMB190",
    "e195":    "EMB195",
    "erj175":  "EMB175",
    "erj170":  "EMB170",
    "erj190":  "EMB190",
    "erj195":  "EMB195",
    "erj-175": "EMB175",
    "erj-170": "EMB170",
    "erj-190": "EMB190",
    "erj-195": "EMB195",
    "embraer 175": "EMB175",
    "embraer 170": "EMB170",
    "embraer 190": "EMB190",
    "embraer 195": "EMB195",
}

NATIONALITY_KEYWORDS = [
    "nationality","citizenship","passport","citizen of","national of",
]

LICENCE_AUTH_TO_COUNTRY = {
    "FAA":   "United States", "ANAC":  "Brazil",       "SACAA": "South Africa",
    "JCAB":  "Japan",         "CASA":  "Australia",    "TCCA":  "Canada",
    "CAAC":  "China",         "CAAM":  "Malaysia",     "CAAS":  "Singapore",
    "CAAV":  "Vietnam",       "GCAA":  "United Arab Emirates", "NCAA":  "Nigeria",
    "KCAA":  "Kenya",         "HCAA":  "Greece",       "BCAA":  "Bangladesh",
    "CAA":   "United Kingdom","DGCA":  "India",
}

LICENCE_AUTHORITIES = ["EASA","FAA","ICAO","GCAA","CAA","DGCA","CAAC","CASA",
                       "ANAC","DGAC","SACAA","KCAA","NCAA","HCAA","TCCA",
                       "JCAB","CAAS","CAAM","CAAV","AAI","BCAA"]

POSITIONS = ["Captain","Senior First Officer","First Officer","Second Officer",
             "Co-Pilot","Copilot","Relief Captain","Training Captain",
             "Check Captain","Line Captain","Junior First Officer",
             "Chief Flight Instructor","Senior Flight Instructor","Flight Instructor"]

CABIN_SENIORITY_MAP = {
    # Valid cabin service levels: Senior / Cabin Manager, VIP, Business, Main Crew
    # Only one per profile — most recent wins.
    # Senior / Cabin Manager
    "senior / cabin manager": "Senior / Cabin Manager",
    "cabin manager":          "Senior / Cabin Manager",
    "inflight manager":       "Senior / Cabin Manager",
    "in-flight manager":      "Senior / Cabin Manager",
    "chief cabin":            "Senior / Cabin Manager",
    "lead cabin":             "Senior / Cabin Manager",
    "senior cabin":           "Senior / Cabin Manager",
    "senior crew":            "Senior / Cabin Manager",
    "senior flight attendant":"Senior / Cabin Manager",
    "purser":                 "Senior / Cabin Manager",
    "chief purser":           "Senior / Cabin Manager",
    # VIP
    "vip":               "VIP",
    "vvip":              "VIP",
    "private jet":       "VIP",
    "private aviation":  "VIP",
    "charter":           "VIP",
    "head of state":     "VIP",
    # Business
    "business class":    "Business",
    "business cabin":    "Business",
    "j class":           "Business",
    "c class":           "Business",
    # Main Crew is the default — no explicit keyword check needed
}

# Country name → Tracker nationality skill name (common mismatches)
COUNTRY_ALIASES = {
    # United Kingdom variants
    "uk": "United Kingdom", "britain": "United Kingdom", "british": "United Kingdom",
    "great britain": "United Kingdom", "england": "United Kingdom", "english": "United Kingdom",
    "scotland": "United Kingdom", "scottish": "United Kingdom",
    "wales": "United Kingdom", "welsh": "United Kingdom", "irish": "Ireland",
    # United States
    "usa": "United States", "us": "United States", "american": "United States",
    "u.s.a": "United States", "u.s": "United States",
    # Middle East
    "uae": "United Arab Emirates", "emirati": "United Arab Emirates",
    "ksa": "Saudi Arabia", "saudi": "Saudi Arabia", "saudi arabian": "Saudi Arabia",
    "palestinian": "Palestine, State of", "palestine": "Palestine, State of",
    "jordanian": "Jordan", "lebanese": "Lebanon", "kuwaiti": "Kuwait",
    "bahraini": "Bahrain", "qatari": "Qatar", "omani": "Oman",
    "yemeni": "Yemen", "iraqi": "Iraq", "syrian": "Syria",
    "israeli": "Israel", "libyan": "Libya",
    "iranian": "Iran, Islamic Republic of", "iran": "Iran, Islamic Republic of",
    # East Asia
    "south korea": "Korea, Republic of", "korean": "Korea, Republic of",
    "north korea": "Korea, Democratic People's Republic of",
    "taiwan": "Taiwan, Province of China", "taiwanese": "Taiwan, Province of China",
    "chinese": "China", "japanese": "Japan", "hong konger": "Hong Kong",
    "hong kong": "Hong Kong",
    # Southeast Asia
    "filipino": "Philippines", "philippine": "Philippines",
    "indonesian": "Indonesia", "malaysian": "Malaysia",
    "thai": "Thailand", "singaporean": "Singapore",
    "vietnamese": "Viet Nam", "vietnam": "Viet Nam",
    "burmese": "Myanmar", "myanmar": "Myanmar",
    "cambodian": "Cambodia", "laotian": "Lao People's Democratic Republic",
    "laos": "Lao People's Democratic Republic",
    # South Asia
    "indian": "India", "pakistani": "Pakistan",
    "bangladeshi": "Bangladesh", "sri lankan": "Sri Lanka",
    "nepali": "Nepal", "nepalese": "Nepal",
    # Central Asia & Caucasus
    "kazakh": "Kazakhstan", "uzbek": "Uzbekistan",
    "georgian": "Georgia", "armenian": "Armenia", "azerbaijani": "Azerbaijan",
    # Russia & Eastern Europe
    "russia": "Russian Federation", "russian": "Russian Federation",
    "ukrainian": "Ukraine", "belarusian": "Belarus",
    "moldova": "Moldova, Republic of", "moldovan": "Moldova, Republic of",
    "macedonian": "North Macedonia", "macedonia": "North Macedonia",
    # Western/Central Europe
    "turkish": "Turkey", "greek": "Greece",
    "italian": "Italy", "french": "France", "german": "Germany",
    "spanish": "Spain", "portuguese": "Portugal", "dutch": "Netherlands",
    "belgian": "Belgium", "swiss": "Switzerland", "austrian": "Austria",
    "swedish": "Sweden", "norwegian": "Norway", "danish": "Denmark",
    "finnish": "Finland", "icelandic": "Iceland",
    # Eastern Europe
    "polish": "Poland", "romanian": "Romania", "bulgarian": "Bulgaria",
    "hungarian": "Hungary", "slovak": "Slovakia", "czech": "Czech Republic",
    "czechia": "Czech Republic", "croatian": "Croatia", "serbian": "Serbia",
    "slovenian": "Slovenia", "bosnian": "Bosnia and Herzegovina",
    "albanian": "Albania", "montenegrin": "Montenegro",
    # Africa
    "moroccan": "Morocco", "algerian": "Algeria", "tunisian": "Tunisia",
    "egyptian": "Egypt", "nigerian": "Nigeria", "kenyan": "Kenya",
    "ghanaian": "Ghana", "south african": "South Africa", "southafrican": "South Africa",
    "zimbabwean": "Zimbabwe", "ugandan": "Uganda", "rwandan": "Rwanda",
    "cameroonian": "Cameroon", "senegalese": "Senegal",
    "ivorian": "Côte d'Ivoire", "ethiopian": "Ethiopia",
    "tanzanian": "Tanzania, United Republic of",
    "tanzania": "Tanzania, United Republic of",
    "zambian": "Zambia", "malawian": "Malawi", "mozambican": "Mozambique",
    "angolan": "Angola", "congolese": "Congo, Democratic Republic of the",
    "sudanese": "Sudan", "somalian": "Somalia", "somali": "Somalia",
    "eritrean": "Eritrea", "djiboutian": "Djibouti",
    # Oceania
    "australian": "Australia", "new zealander": "New Zealand",
    "new zealand": "New Zealand",
    # Americas
    "canadian": "Canada", "mexican": "Mexico", "brazilian": "Brazil",
    "argentinian": "Argentina", "argentinean": "Argentina",
    "colombian": "Colombia", "peruvian": "Peru", "chilean": "Chile",
    "venezuelan": "Venezuela, Bolivarian Republic of",
    "venezuela": "Venezuela, Bolivarian Republic of",
    "ecuadorian": "Ecuador", "bolivian": "Bolivia, Plurinational State of",
    "bolivia": "Bolivia, Plurinational State of",
    "uruguayan": "Uruguay", "paraguayan": "Paraguay",
    "cuban": "Cuba", "jamaican": "Jamaica",
    "trinidad": "Trinidad and Tobago", "trinidadian": "Trinidad and Tobago",
    "syrian arab": "Syria",
}

# Reverse of COUNTRY_ALIASES: country name → adjective form (e.g. "India" → "indian").
# Built once so resolve_skills can fall back to adjective form when Tracker stores
# nationalities as adjectives rather than country names.
_REVERSE_COUNTRY_ALIASES = {}
for _adj, _cn in COUNTRY_ALIASES.items():
    _cn_lower = _cn.strip().lower()
    if _cn_lower not in _REVERSE_COUNTRY_ALIASES:
        _REVERSE_COUNTRY_ALIASES[_cn_lower] = _adj

# All valid country/nationality names (used to avoid sending id=0 for countries
# that genuinely don't exist in this Tracker instance's skill areas).
_ALL_KNOWN_COUNTRY_NAMES = {v.strip().lower() for v in COUNTRY_ALIASES.values()}


def _groq_parse_cv(cv_text, candidate_name):
    """
    Use Groq's free cloud LLM API to extract employer, job title and work type from a CV.
    Returns dict or None if the API key is missing / call fails (falls back to regex).
    Works anywhere — local machine, GitHub Actions, no machine required.
    """
    global _GROQ_WARNED
    import json as _json

    if not GROQ_API_KEY:
        if not _GROQ_WARNED:
            print("  ℹ  GROQ_API_KEY not set — using regex CV parser.")
            print("     Get a free key at https://console.groq.com and add it as GROQ_API_KEY.")
            _GROQ_WARNED = True
        return None

    # First 2500 chars cover employer/title in virtually every CV
    _text = (cv_text or "").strip()[:2500]

    _system = (
        "You are a precise data extractor for aviation recruitment. "
        "You extract exactly three fields from CV text and return valid JSON only. "
        "Never include explanations, markdown, or text outside the JSON object."
    )

    _user = f"""Extract from this aviation CV:

CV TEXT:
{_text}

Return this JSON with exactly these three keys:

{{
  "employer": "<current or most recent COMPANY NAME only. Must be just the organisation name — no job titles, no dates, no cities, no descriptions. E.g. 'British Airways' not 'May 2023 to Date British Airways'. E.g. 'Memphis Airlines' not 'Aeronautical Engineer — Memphis Airlines'. E.g. 'Air Arabia' not 'Air Arabia Dammam, Saudi Arabia'. null if genuinely not stated.>",
  "job_title": "<current or most recent JOB TITLE only. Never a company name, CV document header (like Resume or Curriculum Vitae), or description. null if not stated>",
  "work_type": "<exactly one of: Flight Deck | Engineering | Cabin Crew | Management | Operations>"
}}

RULES:
- employer must be ONLY the company name. Strip any date prefix, location suffix, or job title prefix.
- If the CV shows 'Aviation Manager, Glencore Coal' the employer is 'Glencore Coal'.
- If the CV shows 'May 2023 to Date British Airways' the employer is 'British Airways'.
- If the CV shows 'Aeronautical Engineer — Memphis Airlines' the employer is 'Memphis Airlines'.
- job_title must NEVER be 'Resume', 'CV', 'Curriculum Vitae' or any document header word.

work_type guide:
- Flight Deck: pilots, captains, first officers, flight instructors
- Engineering: aircraft maintenance engineers, AMEs, avionics, B1/B2 licence holders
- Cabin Crew: cabin crew, flight attendants, pursers
- Management: managers, directors, executives, CAMO, safety managers, HR, finance
- Operations: dispatchers, ground ops, logistics, load control, non-aviation roles"""

    try:
        _r = requests.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers={"Authorization": f"Bearer {GROQ_API_KEY}",
                     "Content-Type": "application/json"},
            json={
                "model":           GROQ_MODEL,
                "messages":        [{"role": "system", "content": _system},
                                    {"role": "user",   "content": _user}],
                "temperature":     0,
                "max_tokens":      200,
                "response_format": {"type": "json_object"},
            },
            timeout=30,
        )
        if _r.status_code != 200:
            return None

        _raw  = _r.json()["choices"][0]["message"]["content"].strip()
        _data = _json.loads(_raw)

        _emp = (_data.get("employer") or "").strip()
        _jt  = (_data.get("job_title") or "").strip()
        _wt  = (_data.get("work_type") or "").strip()

        # Treat placeholder strings as empty
        _NULL_VALS = {"null", "none", "n/a", "na", "unknown", "not stated",
                      "not mentioned", "not specified", ""}
        if _emp.lower() in _NULL_VALS:
            _emp = ""
        if _jt.lower() in _NULL_VALS:
            _jt = ""

        # Sanity: reject if employer ends in a role word (LLM confused job title with employer)
        if _emp and re.search(
                r"\b(manager|director|engineer|officer|pilot|captain|technician|"
                r"lead|head|chief|specialist|consultant|analyst|coordinator|"
                r"supervisor|instructor|trainer)\s*$", _emp, re.IGNORECASE):
            _emp = ""

        # Reject implausibly long values
        if len(_emp) > 80:
            _emp = ""
        if len(_jt) > 100:
            _jt = ""

        # Normalise work_type to exact strings process_one expects
        _WT_MAP = {
            "flight deck": "Flight Deck", "pilot": "Flight Deck",
            "engineering": "Engineering", "maintenance": "Engineering",
            "cabin crew":  "Cabin Crew",  "cabin": "Cabin Crew",
            "management":  "Management",
            "operations":  "Operations",  "ops": "Operations",
        }
        _wt_norm = _WT_MAP.get(_wt.lower(), _wt)
        if _wt_norm not in {"Flight Deck", "Engineering", "Cabin Crew", "Management", "Operations"}:
            _wt_norm = ""

        return {"employer": _emp, "job_title": _jt, "work_type": _wt_norm}

    except Exception:
        return None


def parse_cv(cv_text, candidate_name):
    """
    Aviation CV parser — uses Groq LLM for employer/job title/work type,
    regex for skills, nationality and licence extraction.
    Returns dict: {job_title, current_employer, work_type, skills}
    """
    text  = cv_text or ""
    lower = text.lower()

    # ── 1. Classify work type ──────────────────────────────────────────────────
    flight_deck_score = sum([
        lower.count("captain"),
        lower.count("first officer"),
        lower.count("co-pilot") + lower.count("copilot"),
        lower.count("second officer"),
        lower.count("atpl"),
        lower.count("cpl(a)") + lower.count("cpl (a)"),
        lower.count("type rating"),
        lower.count("flight deck"),
        lower.count("pf ") + lower.count("pm "),
        lower.count("command"),
        lower.count("flight instructor") * 3,  # strong signal — always Flight Deck
        lower.count("flying instructor") * 3,
    ])
    cabin_score = sum([
        lower.count("cabin crew"),
        lower.count("flight attendant"),
        lower.count("purser"),
        lower.count("inflight") + lower.count("in-flight"),
        lower.count("cabin manager"),
        lower.count("senior cabin"),
        lower.count("cabin service"),
    ])
    eng_score = sum([
        lower.count("aircraft maintenance"),
        lower.count(" ame ") + lower.count(" ame\n"),
        lower.count("licensed engineer"),
        lower.count("part-66") + lower.count("part 66"),
        lower.count("b1 ") + lower.count("b2 "),
        lower.count("mro"),
        lower.count("avionics"),
        lower.count("airframe"),
        lower.count("powerplant"),
    ])
    mgmt_score = sum([
        lower.count(" director"),
        lower.count("head of"),
        lower.count(" vp ") + lower.count("vice president"),
        lower.count("chief "),
        lower.count("ceo") + lower.count("coo") + lower.count("cfo") + lower.count("csco") + lower.count("cso"),
        lower.count("supply chain"),
        lower.count("logistics manager") + lower.count("logistics director"),
        lower.count("manager") * 1,
    ])
    ops_score = sum([
        lower.count("operations controller"),
        lower.count("ops controller"),
        lower.count("dispatcher"),
        lower.count("load controller"),
        lower.count("crew scheduling"),
        lower.count("flight operations"),
        lower.count("warehouse") + lower.count("procurement"),
    ])
    airport_score = sum([
        lower.count("ground staff"),
        lower.count("ramp agent"),
        lower.count("check-in agent"),
        lower.count("baggage handler"),
        lower.count("ground handling"),
    ])

    scores = {
        "flight deck": flight_deck_score,
        "cabin crew":  cabin_score,
        "engineering": eng_score,
        "management":  mgmt_score,
        "operations":  ops_score,
        "airport":     airport_score,
    }
    top_type = max(scores, key=scores.get)
    if scores[top_type] == 0:
        work_type = "cabin crew"  # default for unclassified registrations — almost always CC applicants
    else:
        # Include secondary work type when both management and operations score > 0
        # (e.g. Operations Manager → both "operations" and "management")
        work_types_out = [top_type]
        if top_type == "management" and scores["operations"] > 0:
            work_types_out = ["operations", "management"]
        elif top_type == "operations" and scores["management"] > 0:
            work_types_out = ["operations", "management"]
        # Dispatchers: Operations primary + Flight Deck secondary (they work with aircraft)
        elif top_type == "operations" and any(x in lower for x in ["dispatcher", "flight dispatch", "aircraft dispatch"]):
            work_types_out = ["operations", "flight deck"]
        work_type = ",".join(work_types_out)

    # ── 2. Extract current employer ────────────────────────────────────────────
    employer = ""

    # Words that make an employer string invalid (skills lists, bullet summaries, etc.)
    EMPLOYER_ACTION_WORDS = re.compile(
        r"\b(management|leadership|training|development|operations|strategy|"
        r"procurement|stakeholder|implementation|responsible|coordination|"
        r"improvement|sop|oversight|delivery|initiative|planning|maintenance|"
        r"outstation|line maintenance|assessment|damage|evaluation|inspection|"
        r"principles|procedures|standards|regulations|compliance|"
        r"shows?|display|exhibition|expo|simulator|simulation|"
        r"sector|supply|leasing|distribution|optimis[ae]tion|optimiz[ae]tion|"
        r"flight\s+time|total\s+hours|flying\s+hours|knowledge\s+of|excellent|"
        r"completion\s+date|qualification|certificate)\b",
        re.IGNORECASE
    )
    # Keywords that confirm something IS a real company name
    COMPANY_KEYWORDS = re.compile(
        r"\b(airline|airways|aviation|air |airport|cargo|express|group|ltd|llc|"
        r"inc\b|plc|corp|company|co\b|technic|tech|services|service|solutions|"
        r"international|global|consulting|enterprises|authority|council|"
        r"pharma|catering|logistics|transport|trading|industries)\b",
        re.IGNORECASE
    )
    # Name parts of the candidate (so we never use their name as employer)
    cand_name_parts = {p.lower() for p in candidate_name.split() if len(p) > 2}

    # Helper: strip trailing date/pipe noise from employer lines
    def clean_employer(raw):
        # Strip leading non-letter characters (e.g. "❖ Air Cairo" → "Air Cairo")
        raw = re.sub(r"^[^A-Za-z0-9]+", "", raw).strip()
        # Strip leading numbered list prefix: "1. Company" → "Company", "2) Company" → "Company"
        raw = re.sub(r"^\d+[.\)]\s*", "", raw).strip()
        # Strip "Personal information" CV section header prefix
        # e.g. "Personal information Aviation Expérience" → "Aviation Expérience"
        raw = re.sub(r"^personal\s+information\s*", "", raw, flags=re.IGNORECASE).strip()
        # Strip "EMPLOYER." / "EMPLOYER: " label prefix (OCR artefact from structured CVs)
        raw = re.sub(r"^EMPLOYER[.:\s]+", "", raw, flags=re.IGNORECASE).strip()
        # Strip trailing incomplete bracket / pipe characters left by parser
        raw = raw.rstrip(" []()|")
        # Strip unclosed parenthetical at end: "Airbus Helicopters (Helibras" → "Airbus Helicopters"
        # Also strips fully closed aircraft-type suffix: "Cargojet Airways ( B767/757 )" → "Cargojet Airways"
        raw = re.sub(r"\s*\([^)]*$", "", raw).strip()  # unclosed "("
        raw = re.sub(r"\s*\(\s*[^)]{1,30}\s*\)\s*$",
                     lambda m: "" if re.search(
                         r"\b[AB]\d{3}|CRJ|ERJ|ATR|E\d{3}|MD\d|DHC|Dash|Saab", m.group(), re.IGNORECASE
                     ) else m.group(),
                     raw).strip()  # closed parenthetical with aircraft type only
        # Strip leading year (e.g. "2019   Lufthansa Group" → "Lufthansa Group")
        raw = re.sub(r"^\d{4}\s+", "", raw).strip()
        # Strip leading date prefix: "May'23 to Date British Airways", "Jan 2019 – Till Date MASCO"
        raw = re.sub(
            r"^(?:from\s+)?(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|"
            r"jul(?:y)?|aug(?:ust)?|sep(?:tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
            r"['\s]\d{2,4}\s+(?:to|–|-|till|until)\s+(?:date|present|now|current)\s+",
            "", raw, flags=re.IGNORECASE).strip()
        # Strip "on [aircraft/equipment]" suffix: "Jordan Aviation Airline on B737-" → "Jordan Aviation Airline"
        raw = re.sub(r"\s+on\s+[A-Z]\d[\w\-]*.*$", "", raw).strip()
        # Strip trailing work-mode suffixes: ", Remote", ", On-site", ", Hybrid", etc.
        raw = re.sub(r",?\s+\b(remote|on-?site|hybrid|home.?based|virtual|contract|freelance)\s*$",
                     "", raw, flags=re.IGNORECASE).strip()
        # Strip trailing European-format dates and everything after them
        # e.g. "Work experience: 25.03.2024 – present" / "HELINT AVIATION LIMITED 6/10/2025 to"
        # Allow single-digit day/month (6/10/2025)
        raw = re.sub(r"\s*:?\s*\d{1,2}[./]\d{1,2}[./]\d{4}.*$", "", raw).strip()
        # Also strip trailing standalone "to", "from", "until", "since" left after date removal
        raw = re.sub(r"\s+\b(to|from|until|since)\s*$", "", raw, flags=re.IGNORECASE).strip()
        # Reject CV section headers — these are never company names
        # e.g. "AVIATION CAREER", "WORK EXPERIENCE", "EMPLOYMENT HISTORY"
        _SECTION_HEADER_RE = re.compile(
            r"^(aviation career|work experience|employment history|career history|"
            r"career summary|professional experience|professional background|"
            r"professional history|work history|job history|career profile|"
            r"career overview|aviation experience|flying experience|"
            r"experience|employment|career)\s*$",
            re.IGNORECASE
        )
        if _SECTION_HEADER_RE.match(raw.strip()):
            return ""
        # Strip "Job Title — Company" when the prefix is a role word (not itself a company)
        # e.g. "Aeronautical Engineer — Memphis Airlines" → "Memphis Airlines"
        # But preserve "Air France – KLM" (no role word before the dash)
        _em_pre_m = re.search(r"^(.+?)\s+[–\-]\s+(.+)$", raw)
        if _em_pre_m:
            _pre  = _em_pre_m.group(1).strip()
            _post = _em_pre_m.group(2).strip()
            _ROLE_PRE_RE = re.compile(
                r"\b(engineer|officer|manager|director|analyst|consultant|specialist|"
                r"supervisor|coordinator|instructor|technician|pilot|captain|crew|"
                r"executive|head|lead|chief|founder|owner|trainee|intern|assistant)\b",
                re.IGNORECASE
            )
            _CO_PRE_RE = re.compile(
                r"\b(airline|airways|aviation|air\b|airport|cargo|group|ltd|llc|inc\b|"
                r"plc|corp\b|company|co\b|services|authority|industries|solutions|"
                r"international|global|holdings|technologies|systems)\b",
                re.IGNORECASE
            )
            if (_ROLE_PRE_RE.search(_pre) and not _CO_PRE_RE.search(_pre) and _post):
                raw = _post

        # Handle "Job Title, Company Name" — "Aviation Manager, Glencore Coal" → "Glencore Coal"
        # Only trigger when before-comma has a role word and no company keyword.
        _TITLE_WORD_QUICK_COMMA = re.compile(
            r"\b(manager|director|officer|coordinator|instructor|executive|analyst|"
            r"specialist|consultant|supervisor|administrator|advisor|technician|"
            r"operator|planner|engineer|pilot|captain|crew|attendant|head|lead)\b",
            re.IGNORECASE
        )
        _comma_idx = raw.find(", ")
        if _comma_idx > 0:
            _before_c = raw[:_comma_idx]
            _after_c  = raw[_comma_idx + 2:].strip()
            _CO_QUICK_RE = re.compile(
                r"\b(airline|airways|aviation|air\b|airport|cargo|group|ltd|llc|inc\b|"
                r"plc|corp\b|company|co\b|services|authority|industries|solutions|"
                r"international|global|holdings|technologies|systems|express)\b",
                re.IGNORECASE
            )
            if (_TITLE_WORD_QUICK_COMMA.search(_before_c)
                    and not _CO_QUICK_RE.search(_before_c)
                    and _after_c
                    and len(_after_c) >= 4
                    and "," not in _after_c):
                raw = _after_c

        # Handle "Job Title at Company Name" — extract just the company part.
        # e.g. "B737 First Officer at Iraqi Airways" → "Iraqi Airways"
        # Only trigger when the part BEFORE "at" contains a role/title word.
        _TITLE_WORD_QUICK = re.compile(
            r"\b(officer|captain|pilot|crew|engineer|manager|director|instructor|"
            r"purser|attendant|controller|dispatcher|coordinator|supervisor|"
            r"executive|analyst|consultant|specialist|technician|operator)\b",
            re.IGNORECASE
        )
        # Case-sensitive search — avoids treating "AT" in "AT PHONE co." as the preposition "at"
        _at_idx = raw.find(" at ")
        if _at_idx > 0:
            _before_at = raw[:_at_idx]
            _after_at  = raw[_at_idx + 4:].strip()
            if _TITLE_WORD_QUICK.search(_before_at) and _after_at and len(_after_at) > 3:
                raw = _after_at
        # Handle "Job Title of Company" — e.g. "Marketing Manager of AT PHONE co., ltd."
        # Only extract when title word precedes "of" AND the after-part has a company keyword.
        _of_idx = raw.lower().find(" of ")
        if _of_idx > 0 and raw.find(" at ") < 0:   # don't double-split
            _before_of = raw[:_of_idx]
            _after_of  = raw[_of_idx + 4:].strip()
            _COMPANY_ENTITY_QUICK = re.compile(
                r"\b(airline|airways|aviation|air\b|airport|cargo|group|ltd|llc|inc\b|plc|corp\b|"
                r"company|co\b|services?|authority|council|logistics|transport|trading|"
                r"industries|enterprises?|solutions?|international|global|holdings|"
                r"technologies|systems|partners|associates|foundation|bank|express)\b",
                re.IGNORECASE
            )
            if (_TITLE_WORD_QUICK.search(_before_of)
                    and _after_of
                    and len(_after_of) > 3
                    and _COMPANY_ENTITY_QUICK.search(_after_of)):
                raw = _after_of
        # Remove trailing "| 2020 – Present", "2019-current", etc.
        raw = re.sub(r"[\|–\-]\s*\d{4}.*$", "", raw).strip()
        raw = re.sub(r"\s*[\|–\-]\s*(present|current|now).*$", "", raw, flags=re.IGNORECASE).strip()
        # Remove trailing "Jul 2018", "September 2019" etc. (month + year, no separator needed)
        raw = re.sub(r"\s+(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
                     r"Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
                     r"\s+\d{4}.*$", "", raw, flags=re.IGNORECASE).strip()
        # Remove trailing lone year
        raw = re.sub(r"\s+\d{4}$", "", raw).strip()
        raw = raw.rstrip(".,;|– ")
        # Strip trailing MM/YYYY or MM/YYYY – MM/YYYY date ranges
        # e.g. "Avion Express 08/2017 – 04/2019" → "Avion Express"
        raw = re.sub(r"\s+\d{1,2}/\d{4}.*$", "", raw).strip()
        raw = raw.rstrip(".,;|– ")
        # Strip trailing " – Location" suffix (city/country after em-dash, no company signal)
        # e.g. "Kenya Airways – Nairobi" → "Kenya Airways"
        # Preserved: "Air France – KLM" (KLM is uppercase airline-code-like token)
        _em_m = re.search(r"\s+[–\-]\s+([A-Za-z][a-zA-Z ]{2,30})$", raw)
        if _em_m:
            _em_suf = _em_m.group(1).strip()
            if not re.search(r"\b(air|ways|lines|group|ltd|llc|inc|plc|corp|aviation)\b",
                             _em_suf, re.IGNORECASE):
                raw = raw[:_em_m.start()].strip()
        # Strip trailing " – Job Title [Aircraft]" suffix: "ETIHAD AIRWAYS – Captain A320" → "ETIHAD AIRWAYS"
        # This handles cases not caught above (suffix contains numbers e.g. aircraft type)
        _em_role_m = re.search(r"\s+[–\-]\s+(.+)$", raw)
        if _em_role_m:
            _em_role_suf = _em_role_m.group(1).strip()
            _ROLE_SUFFIX_RE = re.compile(
                r"\b(captain|pilot|officer|engineer|manager|director|analyst|consultant|"
                r"specialist|supervisor|coordinator|instructor|technician|crew|attendant)\b",
                re.IGNORECASE
            )
            if _ROLE_SUFFIX_RE.search(_em_role_suf):
                raw = raw[:_em_role_m.start()].strip()
        # Strip "–Description" with no leading space (em-dash immediately before preposition/keyword)
        # e.g. "Shree Airlines –From Technical Services In-charge" → "Shree Airlines"
        raw = re.sub(r"[–\-](From|Until|Since|As\s|For\s|With\s|Technical|Handling|Based|Located)\b.*$",
                     "", raw).strip()
        raw = raw.rstrip(".,;|– ")
        # Reject if it looks like a job title — "Airline Pilot", "Captain", "First Officer" etc.
        # These get mistaken for employers when CVs contain "airline" in the title line.
        JOB_TITLE_PATTERN = re.compile(
            r"^(airline\s+(?:pilot|captain|officer)|commercial pilot|"
            r"senior first officer|first officer|"
            r"second officer|junior first officer|relief captain|training captain|"
            r"check captain|line captain|co-pilot|copilot|captain|pilot|"
            r"cabin crew|senior cabin crew|flight attendant|purser|chief purser|"
            r"aircraft engineer|maintenance engineer|avionics engineer|"
            r"flight engineer|ground engineer|certifying engineer)\b",
            re.IGNORECASE
        )
        if JOB_TITLE_PATTERN.match(raw):
            return ""
        # Reject any string that looks like a job title — covers ALL role descriptor words.
        # Logic: if the string contains a role word AND no company entity word → it's a job title.
        # e.g. "Partnerships Manager" → rejected; "Kish Air" → kept; "Airport Authority" → kept.
        _COMPANY_ENTITY_RE = re.compile(
            r"\b(airline|airways|aviation|air\b|airport|cargo|express|group|ltd|llc|"
            r"inc\b|plc|corp\b|company|co\b|services|authority|council|industries|"
            r"enterprises|solutions|international|global|logistics|transport|trading|"
            r"systems|technologies|partners|associates|holdings|ventures|foundation|"
            r"forces|corps\b|ministry|department|bureau|agency|board|bank|hospital|"
            r"hospital|clinic|hotel|resort|media|publishing|insurance|investments|"
            r"properties|construction|manufacturing|engineering\b)\b",
            re.IGNORECASE
        )
        _ALL_TITLE_WORDS_RE = re.compile(
            r"\b(manager|director|officer|coordinator|instructor|executive|analyst|"
            r"specialist|consultant|supervisor|administrator|associate|advisor|"
            r"representative|technician|operator|planner|scheduler|auditor|"
            r"accountant|lawyer|attorney|doctor|physician|nurse|teacher|trainer|"
            r"developer|designer|architect|researcher|scientist|professor|lecturer|"
            r"secretary|receptionist|clerk|assistant|intern|apprentice|trainee|"
            r"agent|handler|warehouse|ramp|ground|loader|dispatcher|controller|"
            r"owner|founder|partner|president|vice president|"
            r"ceo\b|coo\b|cfo\b|cto\b|cmo\b|cso\b|cio\b|chro\b|vp\b|"
            r"head of|chief of|director of|manager of|lead\b|principal\b|"
            r"captain|pilot|crew|attendant|purser|dispatcher|controller)\b",
            re.IGNORECASE
        )
        if _ALL_TITLE_WORDS_RE.search(raw) and not _COMPANY_ENTITY_RE.search(raw):
            return ""
        # Reject single generic-entity words that are never a real company name on their own.
        # E.g. "Logistics", "Services", "Solutions" — these pass _COMPANY_ENTITY_RE but
        # cannot identify an actual employer.
        _STANDALONE_GENERIC_WORDS = {
            "logistics", "services", "solutions", "transport", "transportation",
            "trading", "technologies", "systems", "industries", "enterprises",
            "management", "consulting", "consultancy", "advisory",
            "international", "partners", "associates", "holdings", "ventures",
            "properties", "construction", "manufacturing",
        }
        if raw.strip().lower() in _STANDALONE_GENERIC_WORDS:
            return ""
        # Extra check: strings ending in "Lead", "Head", "Chief" are job titles even when they
        # also contain "engineering" etc. (which would otherwise save them via _COMPANY_ENTITY_RE).
        # e.g. "Engineering Team Lead" → job title, not a company.
        _ENDS_WITH_LEAD_RE = re.compile(
            r"\b(team\s+lead|team\s+leader|lead\b|head\b|chief\b)\s*$", re.IGNORECASE
        )
        if _ENDS_WITH_LEAD_RE.search(raw):
            return ""
        # Strip trailing suffix after comma:
        #   - single-word location: "SITSCO WLL, Bahrain" → "SITSCO WLL"
        #   - title word fragment: "XYZ Ltd., Assistant Engineer" → "XYZ Ltd."
        _ci = raw.find(",")
        if _ci > 0:
            _after_comma = raw[_ci + 1:].strip()
            if (len(_after_comma.split()) == 1
                    and len(_after_comma) >= 5
                    and not _COMPANY_ENTITY_RE.search(_after_comma)):
                # Single-word location/country suffix → strip
                raw = raw[:_ci].strip()
            elif _ALL_TITLE_WORDS_RE.search(_after_comma):
                # Job title fragment after comma → strip
                raw = raw[:_ci].strip()
            elif len(_after_comma) <= 3 or re.match(r'^[\d\s]+$', _after_comma):
                # Short or purely numeric suffix (e.g. ", 9", ", 14") → strip
                raw = raw[:_ci].strip()
        # Reject if the string starts with an entity suffix (no company name before it)
        # e.g. "Ltd. Baramati (M.H)" — "Ltd." at position 0 means we have no real name
        if re.match(r"^(ltd\.?|llc|inc\.?|corp\.?|plc)\b", raw, re.IGNORECASE):
            return ""
        # Reject strings containing parenthetical role abbreviations like "(LAME)", "(AME)", "(retd)"
        # These are personal credentials or qualifiers embedded in CV text, not company names.
        # e.g. "AIRLINES (LAME) BOEING 787-", "Squadron Leader (retd) Indian Air Force"
        if re.search(r"\((?:LAME|AME|CAME|MRO|CAMO|TRE|TRI|SFI|retd|Retd|ret'd)\)", raw):
            # Try to extract a cleaner company name after the parenthetical
            _after_paren = re.sub(r".*\([^)]+\)\s*", "", raw).strip()
            if _after_paren and len(_after_paren.split()) <= 5 and _after_paren[0].isupper():
                raw = _after_paren
            else:
                return ""
        # Reject education / qualification lines
        EDUCATION_WORDS = re.compile(
            r"\b(bachelor|master|degree|diploma|bsc|b\.sc|msc|m\.sc|mba|phd|ph\.d|"
            r"b\.eng|m\.eng|licen[cs]es?|certif|qualification|graduate|undergraduate|"
            r"postgraduate|university|college|institute|academy|school|faculty|"
            r"education|courses?|programme|program\b|training programme|safety course|"
            r"atpl|cpl|ppl|ir\b|type rating|transport pilot|airline transport)\b",
            re.IGNORECASE
        )
        if EDUCATION_WORDS.search(raw):
            return ""
        # Reject if too long (>6 words = concatenated bullet points), or contains action words
        if len(raw.split()) > 6 or EMPLOYER_ACTION_WORDS.search(raw):
            return ""
        # Reject strings that start with prepositions/articles — these are sentence fragments,
        # not company names (e.g. "within the aviation industry", "in the UAE", "at a major airline")
        if re.match(r"^(within|in the|at the|at a|for the|for a|from the|from a|from\b|"
                    r"across|through|providing|supporting|covering|handling|managing|working|including|have\b|module\b|"
                    r"current\b|experience\b|employer\b|"
                    r"proficient\b|skilled\b|experienced\b|fluent\b|knowledge\b|expertise\b|"
                    r"responsible\b|accountable\b|specialised?\b)\b",
                    raw, re.IGNORECASE):
            return ""
        # Reject strings that contain CV section header phrases anywhere in them
        # (e.g. "AVIATION EMPLOYMENT HISTORY:", "PROFESSIONAL AVIATION EXPERIENCE")
        if re.search(r"\b(employment\s+history|work\s+history|career\s+history|"
                     r"professional\s+(?:aviation|experience|background)|"
                     r"aviation\s+employment|airline\s+captain|"
                     r"aviation\s+exp[eé]rience|aviation\s+experience)\b", raw, re.IGNORECASE):
            return ""
        # Reject strings that are just a date (e.g. "Oct 2021", "January 2020")
        if re.match(r"^(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
                    r"Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
                    r"\s+\d{4}\s*$", raw, re.IGNORECASE):
            return ""
        if re.match(r"^\d{4}\s*$", raw):
            return ""
        # Reject bare comma between letters — OCR garbage (e.g. "fefueling,air shows")
        if re.search(r"[A-Za-z],[A-Za-z]", raw):
            return ""
        # Reject standalone city names — common in aviation CVs as location references
        _CITY_NAMES = {
            "doha", "dubai", "abu dhabi", "riyadh", "jeddah", "mecca", "medina",
            "cairo", "amman", "beirut", "kuwait", "muscat", "manama", "sanaa",
            "karachi", "lahore", "islamabad", "mumbai", "delhi", "bangalore",
            "nairobi", "addis ababa", "lagos", "accra", "casablanca",
            "bangkok", "singapore", "manila", "jakarta", "kuala lumpur",
        }
        if raw.strip().lower() in _CITY_NAMES:
            return ""
        # Reject aircraft registration / MSN lines — these are aircraft identifiers, not companies
        # e.g. "G450, MSN 4088, Reg.: B"
        if re.search(r"\bMSN\s+\d+\b|\bReg\.\s*:", raw, re.IGNORECASE):
            return ""
        # Reject "N years of [airline/aviation] experience" — experience descriptions, not companies
        # e.g. "26 years of airline experience"
        if re.match(r"^\d+\s+years?\s+of\b", raw, re.IGNORECASE):
            return ""
        # Reject "Job Title (Company Name)" when prefix is a role word — extract just company
        # e.g. "Product Category Manager (KIKA Group" → "KIKA Group"
        _paren_m = re.match(r"^(.+?)\s*\((.+?)\)?$", raw)
        if _paren_m:
            _paren_pre  = _paren_m.group(1).strip()
            _paren_post = _paren_m.group(2).strip()
            _ROLE_PAREN_RE = re.compile(
                r"\b(manager|director|officer|coordinator|instructor|executive|analyst|"
                r"specialist|consultant|supervisor|administrator|advisor|technician|"
                r"engineer|pilot|captain|crew|head|lead|chief)\b",
                re.IGNORECASE
            )
            if _ROLE_PAREN_RE.search(_paren_pre) and _paren_post and len(_paren_post) >= 3:
                raw = _paren_post
        # Reject "Label: Value" patterns — CV form fields, not company names
        # e.g. "Aircraft: C68A", "Base: Dubai", "Aircraft Type: B737"
        # Also rejects description labels: "Total Aviation Work Experience: 13 years"
        if ": " in raw:
            _before_colon = raw.split(": ")[0].strip()
            _DESC_LABEL_RE = re.compile(
                r"\b(experience|history|background|summary|overview|total|years?|date|completion)\b",
                re.IGNORECASE
            )
            if len(_before_colon.split()) <= 2 or _DESC_LABEL_RE.search(_before_colon):
                return ""
        # Strip trailing aircraft-type suffix (e.g. "South African Express Airways - CRJ200")
        raw = re.sub(
            r"\s*[-–/]\s*(?:[AB]\d{3}\w{0,5}|CRJ\d*|ATR\d*|ERJ[-\d]*|EMB\d{3}|E\d{3}|MD\d{2})\s*$",
            "", raw
        ).strip()
        # Reject if it matches the candidate's own name (>50% word overlap)
        raw_words = {w.lower() for w in raw.split() if len(w) > 2}
        if raw_words and cand_name_parts and len(raw_words & cand_name_parts) / len(raw_words) >= 0.5:
            return ""
        return raw

    # 1. Explicit label — use [ \t]*:[ \t]* to avoid spanning newlines
    for pattern in [
        r"(?:current employer|employer|company|organisation|organization)[ \t]*:[ \t]*([A-Za-z][^\n]{2,60})",
        r"(?:working (?:at|for|with)|employed (?:at|by))[ \t]+([A-Za-z][^\n]{2,60})",
    ]:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            val = clean_employer(m.group(1))
            # Skip if what we captured looks like a job title label
            if not re.match(r"designation|position|role|title", val, re.IGNORECASE):
                employer = val
                break

    MONTHS = {"january","february","march","april","may","june","july",
              "august","september","october","november","december",
              "jan","feb","mar","apr","jun","jul","aug","sep","oct","nov","dec"}

    # 2. Look for a line that has "Present" as the end date — that's the current role
    if not employer:
        for line in text.split("\n"):
            line = line.strip()
            # Strip bullet characters before checking
            line_clean = re.sub(r"^[\-–•▪◦\*]+\s*", "", line).strip()
            if re.search(r"\bpresent\b", line_clean, re.IGNORECASE) and re.search(r"\d{4}", line_clean):
                # This line is like "Job Title | Company Name | 2023 – Present"
                # Scan ALL pipe-separated segments — pick the first that looks like a company.
                # This avoids the common error of taking the job title (first segment) as the employer.
                # e.g. "A320 Captain | Lufthansa | 2019-Present" → skips "A320 Captain", picks "Lufthansa"
                segments = re.split(r"\s*\|\s*", line_clean)
                for seg in segments:
                    seg = seg.strip()
                    # Skip date-only segments (e.g. "2019–2022", "2019")
                    if re.match(r"^[\d\s\-–/]+$", seg):
                        continue
                    # Skip segments that are purely a date range ending in present/current
                    if re.search(r"\d{4}\s*[\-–]\s*(present|current|now)\b", seg, re.IGNORECASE):
                        continue
                    # Strip leading year before passing to clean_employer
                    seg_clean = re.sub(r"^\d{4}\s+", "", seg).strip()
                    val = clean_employer(seg_clean)
                    # Skip if it's just a month name, very short, or contains action words
                    if (val and val.lower() not in MONTHS and len(val) > 4
                            and not re.search(r"\b(leadership|cross-functional|responsible|"
                                              r"managed|delivered|achieved|implemented)\b",
                                              val, re.IGNORECASE)):
                        employer = val
                        break
                if employer:
                    break

    # 3. Fallback: first line that looks like a company name
    # Strip leading bullet chars before checking so bulleted job-entry lines are not skipped.
    if not employer:
        for line in text.split("\n"):
            line = line.strip()
            line = re.sub(r"^[\-–•▪◦\*]+\s*", "", line).strip()  # strip bullets first
            if (len(line) > 5 and len(line) < 120
                    and not any(c in line for c in ["@", "http", "+"])
                    and re.search(r"(airline|airways|aviation|\bair\b|airport|cargo|express|group|ltd|llc|inc\b|plc|motors|logistics|supply)", line, re.IGNORECASE)):
                employer = clean_employer(line)
                if employer:
                    break

    # ── 3. Extract position & aircraft type (Flight Deck) ─────────────────────
    position = ""
    aircraft = []
    licence_auth = []
    licence_type = ""
    cabin_level = ""
    eng_licences = []

    if "flight deck" in work_type:
        # Position
        for pos in POSITIONS:
            if re.search(r"\b" + re.escape(pos) + r"\b", text, re.IGNORECASE):
                position = pos
                break
        # Aircraft types — normalise "Boeing 737 NG" → "B737-NG" before matching
        _ac_text = _norm_ac_text(text)
        for ac in AIRCRAFT_TYPES:
            if re.search(r"\b" + re.escape(ac) + r"\b", _ac_text, re.IGNORECASE):
                tracker_name = AIRCRAFT_ALIASES.get(ac.lower(), ac)
                if tracker_name not in aircraft:
                    aircraft.append(tracker_name)
                if len(aircraft) >= 5:
                    break
        # Licence authority
        for auth in LICENCE_AUTHORITIES:
            if re.search(r"\b" + re.escape(auth) + r"\b", text, re.IGNORECASE):
                licence_auth.append(auth)
        # Licence type
        for ltype in ["ATPL(A)","ATPL(H)","CPL(A)","CPL(H)","ATPL","CPL","MPL"]:
            if re.search(r"\b" + re.escape(ltype) + r"\b", text, re.IGNORECASE):
                licence_type = ltype
                break
        # TRI/TRE
        tri_tre = []
        if re.search(r"\bTRI\b", text):
            tri_tre.append("TRI")
        if re.search(r"\bTRE\b", text):
            tri_tre.append("TRE")
        if re.search(r"\bSFI\b", text):
            tri_tre.append("SFI")

    elif work_type == "cabin crew":
        for kw, level in CABIN_SENIORITY_MAP.items():
            if kw in lower:
                cabin_level = level
                break
        if not cabin_level:
            cabin_level = "Main Crew"

    elif work_type == "engineering":
        if re.search(r"\bB1\b", text):
            eng_licences.append("B1")
        if re.search(r"\bB2\b", text):
            eng_licences.append("B2")
        if not eng_licences:
            if re.search(r"\bB1\.1\b|\bB1\.3\b", text):
                eng_licences.append("B1")
            # (B2 already checked above; no need to repeat)
        for ac in AIRCRAFT_TYPES:
            if re.search(r"\b" + re.escape(ac) + r"\b", text, re.IGNORECASE):
                tracker_name = AIRCRAFT_ALIASES.get(ac.lower(), ac)
                if tracker_name not in aircraft:
                    aircraft.append(tracker_name)
                if len(aircraft) >= 3:
                    break

    # ── 4. Extract nationality ─────────────────────────────────────────────────
    nationality = ""
    nat_explicit = False   # True = found via explicit keyword (trustworthy)
    for kw in NATIONALITY_KEYWORDS:
        # Stop at newline, comma, digit or punctuation — prevents capturing the next line
        pattern = kw + r"[ \t]*:?[ \t]*([A-Za-z][A-Za-z ]{2,28}?)(?=\s*[\r\n,\.\d]|$)"
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            raw_nat = m.group(1).strip().rstrip(".,; \t").title()
            # Reject if it looks like a sentence or grabbed extra words (>3 words)
            if not raw_nat or len(raw_nat.split()) > 3:
                continue
            # Try full string then progressively shorter substrings.
            # Handles "South African None", "Details", and other section-header captures.
            _nat_words = raw_nat.split()
            _resolved_nat = None
            for _n in range(len(_nat_words), 0, -1):
                _substr = " ".join(_nat_words[:_n]).lower()
                if _substr in COUNTRY_ALIASES:
                    _resolved_nat = COUNTRY_ALIASES[_substr]
                    break
                if _substr in KNOWN_COUNTRY_NAMES_EXTRA:
                    _resolved_nat = COUNTRY_ALIASES.get(_substr, _substr.title())
                    break
            if not _resolved_nat:
                continue  # captured text isn't a recognisable country — skip
            nationality = _resolved_nat
            nat_explicit = True   # found via explicit keyword
            break


    # Fallback 1: scan education section for country (Emily rule: use school country)
    if not nationality:
        _edu_kws = ["university","college","school","institute","academy",
                    "bachelor","master","mba","b.sc","m.sc","degree",
                    "qualification","education","studied","graduated","diploma"]
        _edu_lines = []
        _in_edu = False
        for _ln in text.split("\n"):
            _ln_lc = _ln.strip().lower()
            if any(_kw in _ln_lc for _kw in _edu_kws):
                _in_edu = True
            if _in_edu:
                _edu_lines.append(_ln.strip())
                if len(_edu_lines) >= 25:
                    break
        _edu_text = " ".join(_edu_lines)
        for _cnt in sorted(KNOWN_COUNTRY_NAMES_EXTRA, key=len, reverse=True):
            if re.search(r"\b" + re.escape(_cnt) + r"\b", _edu_text, re.IGNORECASE):
                nationality = COUNTRY_ALIASES.get(_cnt, _cnt.title())
                break

    # Fallback 2: oldest job location — bottom half of CV (career beginnings = most likely home country)
    # Emily rule: "if no cv stated then it should revert to the location of their first job or school"
    if not nationality:
        _lines_cv = text.split("\n")
        _career_start = "\n".join(_lines_cv[len(_lines_cv) // 2:])
        # Skip countries that appear only in route/destination/transport context
        _ROUTE_CTX_RE = re.compile(
            r"\b(?:to|from|via|fly(?:ing)?|routes?|destination|based\s+in|"
            r"operated?(?:\s+to)?|flights?\s+to|serv(?:es?|ing)|travel(?:l?ing)?)\s*$",
            re.IGNORECASE
        )
        for _cnt in sorted(KNOWN_COUNTRY_NAMES_EXTRA, key=len, reverse=True):
            _pat = r"\b" + re.escape(_cnt) + r"\b"
            _found_valid = False
            for _m in re.finditer(_pat, _career_start, re.IGNORECASE):
                _pre = _career_start[max(0, _m.start() - 60):_m.start()]
                if _ROUTE_CTX_RE.search(_pre.strip()):
                    continue  # route/destination mention — skip
                _found_valid = True
                break
            if _found_valid:
                nationality = COUNTRY_ALIASES.get(_cnt, _cnt.title())
                break

    # Fallback 3: scan first 30 lines for address/header country
    # (least reliable — picks up current work location, not home country)
    if not nationality:
        _hdr = "\n".join(text.split("\n")[:30])
        for _cnt in sorted(KNOWN_COUNTRY_NAMES_EXTRA, key=len, reverse=True):
            if re.search(r"\b" + re.escape(_cnt) + r"\b", _hdr, re.IGNORECASE):
                nationality = COUNTRY_ALIASES.get(_cnt, _cnt.title())
                break

    # ── 5. Extract job title ───────────────────────────────────────────────────
    if "flight deck" in work_type:
        if position and aircraft:
            job_title = f"{aircraft[0]} {position}"
        elif position:
            job_title = position
        elif aircraft:
            job_title = f"Pilot - {aircraft[0]}"
        else:
            job_title = "Pilot"
    else:
        # For non-flight deck, try to extract from CV
        job_title = ""
        # Words that indicate a sentence fragment, not a job title
        TITLE_SKIP = re.compile(
            r"^(where|when|what|which|who|how|can|could|would|should|may|might|"
            r"working|employed|based|responsible|reporting|seeking|looking|"
            r"experience|graduate|i am|my name|dear|please|thank|attached|"
            r"career|objective|summary|profile|education|skill|certif|"
            r"language|reference|contact|address|phone|email|date of|born|"
            r"notice period|notice|availability|joining|willing to relocate|"
            r"professional summary|professional profile|work experience|"
            r"career history|career summary|employment history|key skills|"
            r"core competencies|personal statement|personal profile|personal info|"
            r"personal detail|personal|about me|curriculum vitae|"
            r"curriculum|professional|airline operations|ground operations|"
            r"flight operations|airport operations|cargo operations|"
            # Job portal / application system noise
            r"desde |via |applied via|applied through|submitted via|"
            r"infojobs|linkedin|indeed|glassdoor|reed|totaljobs|jobsite|"
            r"monster|caterer|cv-library|jobserve|aviationjobsearch|"
            r"click here|download|upload|attach|view|update|"
            # Single words that are CV date/section labels, not job titles
            r"present$|current$|ongoing$|till date$|to date$|"
            r"licenses?$|certifications?$|achievements?$|references?$|"
            r"licenses? &|certifications? &|licenses? and)",
            re.IGNORECASE
        )
        # Name parts to exclude (candidate's own name should never be the job title)
        name_parts = {p.lower() for p in candidate_name.split() if len(p) > 2}

        # Words that indicate a sentence/achievement, not a job title
        SENTENCE_WORDS = re.compile(
            r"\b(supplier|preferred|negotiated|implemented|delivered|responsible|"
            r"leadership|cross-functional|stakeholder|procurement|partnership|"
            r"as a |in order to|globally|across|revenue|million|billion|achieved|"
            r"managed a team|led a team|oversaw|spearheaded|streamlined|drove|"
            r"transformation leader|transformation|"
            r"professional with|years of experience|proven track|"
            r"supply chain operations|dynamic|versatile|seasoned|accomplished)\b",
            re.IGNORECASE
        )

        FOREIGN_HEADER_RE = re.compile(
            r"^(informations?\s+personnelles?|donn[eé]es?\s+personnelles?|"
            r"datos?\s+personales?|informaci[oó]n\s+personal|"
            r"persönliche\s+(?:daten|angaben)|angaben\s+zur\s+person|"
            r"informazioni\s+personali|dati\s+personali|"
            r"dane\s+osobowe|gegevens|perso[oó]nli[cj]ke\s+gegevens|"
            r"informações\s+pessoais|dados\s+pessoais|"
            r"(?:personal\s+)?particulars|personal\s+particulars|"
            r"quien\s+soy|experiencia\s+profesional|datos\s+de\s+contacto|"
            r"perfil\s+profesional|mastered\s+languages?|langues?\s+ma[iî]tris[eé]es?)\b",
            re.IGNORECASE
        )
        SCHOOL_IN_TITLE_RE = re.compile(
            r"(patts|college\s+of\s+aeronaut|university\s+of\s+aero|college\s+of\s+aviation|"
            r"aviation\s+(?:colleg|academ|universit|instit)|aeronautical\s+(?:universit|colleg|academ)|"
            r"aviation\s+training\s+cent|flight\s+(?:colleg|academ|school\b)|"
            r"polytechnic|cadet\s+school)",
            re.IGNORECASE
        )
        # CV section headers frequently mistaken for job titles
        SECTION_HDR_RE = re.compile(
            r"^(training\s+courses?|core\s+skills?|key\s+skills?|technical\s+skills?|"
            r"speciali[sz]ed?\s+skills?|specialist\s+skills?|"
            r"further\s+education|additional\s+skills?|professional\s+development|"
            r"diploma\s+in|bachelor|qualification|tax\s+return|income\s+tax|"
            r"in\s+which|areas?\s+of\s+(?:strength|expertise|interest)|"
            r"accomplishment|honour|honor|licensing)\b",
            re.IGNORECASE
        )

        def is_valid_title(t):
            t_lower = t.lower()
            if TITLE_SKIP.match(t):
                return False
            if any(c in t for c in ["@", "http", "+"]):
                return False
            if len(t) <= 3 or len(t) >= 70:
                return False
            # Reject titles that are too many words (objective statements, bullet text)
            if len(t.split()) > 6:
                return False
            # Reject foreign-language CV section headers
            if FOREIGN_HEADER_RE.match(t):
                return False
            # Reject CV section headers mistaken for job titles
            if SECTION_HDR_RE.match(t):
                return False
            # Reject school/institution names
            if SCHOOL_IN_TITLE_RE.search(t):
                return False
            # Reject sentences masquerading as titles
            if SENTENCE_WORDS.search(t):
                return False
            # Reject if it starts with a verb (achievement bullet)
            if re.match(r"^(Led|Managed|Developed|Built|Created|Drove|Designed|"
                        r"Delivered|Implemented|Negotiated|Achieved|Oversaw|"
                        r"Established|Launched|Transformed|Improved|Increased|"
                        r"Reduced|Supported|Coordinated|Executed|Mastered|Spoken|"
                        r"Contributed|Contributing|Contribute|Seeking|Seek|"
                        r"Providing|Provide|Leveraging|Leverage|Utilizing|Utilize|"
                        r"Pursuing|Pursue|Aspiring|Aspire|Obtaining|Obtain|"
                        r"Demonstrat|Maximiz|Ensuring|Ensure|Applying|Apply)\b", t, re.IGNORECASE):
                return False
            # Reject if it looks like the candidate's name (majority of words overlap).
            # Also check surname specifically — Tracker may store "kawinraj" but CV has "Kawin Raj"
            t_words = [w.lower() for w in t.split() if w.lower() not in {"mr","ms","dr","the"}]
            if name_parts and t_words:
                overlap = sum(1 for w in t_words if w in name_parts)
                if overlap / len(t_words) >= 0.5:
                    return False
                # Extra check: candidate's last name alone is strong signal it's their name
                _surname = candidate_name.split()[-1].lower() if candidate_name else ""
                if _surname and len(_surname) > 3 and any(
                    _surname == w or w.startswith(_surname) or _surname.startswith(w)
                    for w in t_words if len(w) > 3
                ):
                    return False
            # Reject single-word titles that are obviously generic/activity words
            if len(t.split()) == 1 and t.lower() in {
                "training","operations","management","support","consulting",
                "sales","marketing","finance","logistics","procurement",
                "planning","administration","admin","supervision","projects",
                "coordination","development","engineering","maintenance",
                "security","quality","compliance","services","recruitment",
            }:
                return False
            return True

        for pattern in [
            # Explicit label: "Job Title:", "Position:", "Role:" etc.
            r"(?:position[:\s]+|role[:\s]+|job title[:\s]+|designation[:\s]+)([A-Za-z][^\n]{3,80})",
            # "Currently/Presently: [title]" — only when directly followed by colon
            r"(?:current(?:ly)?|presently)\s*:\s*([A-Za-z][^\n]{3,80})",
            # "Currently/Presently employed/working as [title]"
            r"(?:current(?:ly)?|presently)\s+(?:employed|working)\s+as\s+([A-Za-z][^\n]{3,80})",
            # Standalone ALL-CAPS or Title-Case line that looks like a job title
            r"^([A-Z][a-zA-Z &/\-]{5,60})$",
        ]:
            for m in re.finditer(pattern, text, re.IGNORECASE | re.MULTILINE):
                candidate_title = m.group(1).strip().rstrip(".,;")
                # Strip trailing location info: ", Riyadh, KSA" / ", Dubai, UAE" etc.
                candidate_title = re.sub(r",\s*[A-Z][a-zA-Z\s]{2,20},\s*[A-Z]{2,5}\s*$", "", candidate_title).strip()
                candidate_title = re.sub(r",\s*[A-Z][a-zA-Z\s]{2,30}\s*$", "", candidate_title).strip() if "," in candidate_title else candidate_title
                if is_valid_title(candidate_title):
                    job_title = candidate_title
                    break
            if job_title:
                break

    # ── 6. Build skills list ───────────────────────────────────────────────────
    skills = []
    if nationality:
        skills.append(nationality)

    if "flight deck" in work_type:
        if position:
            skills.append(position)
        for auth in licence_auth[:2]:
            skills.append(auth)  # authority only — ATPL/CPL/PPL are licence types, not Tracker skills
        skills.extend(aircraft[:4])
        skills.extend(tri_tre)

    elif work_type == "cabin crew":
        if cabin_level:
            skills.append(cabin_level)

    elif work_type == "engineering":
        skills.extend(eng_licences)
        # No aircraft types for engineers — nationality + B1/B2 licence only

    # ── Groq LLM override ────────────────────────────────────────────────────────
    # LLM is far more reliable than regex for employer / job_title / work_type.
    # Skills and nationality are still built from regex (they work well).
    _llm = _groq_parse_cv(cv_text, candidate_name)
    if _llm:
        _llm_emp = clean_employer(_llm.get("employer", "") or "")  # run through same cleaner as regex
        _llm_jt  = _llm.get("job_title", "")
        _llm_wt  = _llm.get("work_type", "")
        if _llm_emp and _llm_emp != employer:
            print(f"  ℹ  LLM employer: '{_llm_emp}'" +
                  (f" (regex had: '{employer}')" if employer else ""))
        if _llm_jt and _llm_jt != job_title:
            print(f"  ℹ  LLM job title: '{_llm_jt}'" +
                  (f" (regex had: '{job_title}')" if job_title else ""))
        if _llm_wt and _llm_wt != work_type:
            print(f"  ℹ  LLM work type: '{_llm_wt}' (regex had: '{work_type}')")
        employer   = _llm_emp or employer
        job_title  = _llm_jt  or job_title
        work_type  = _llm_wt  or work_type

    return {
        "job_title":        job_title,
        "current_employer": employer,
        "work_type":        work_type,
        "skills":           [s for s in skills if s],
        "nationality":      nationality,
        "nat_explicit":     nat_explicit,
    }


# ── Tracker update ─────────────────────────────────────────────────────────────

def get_resource(jwt, resource_id):
    r = requests.get(f"{TRACKER_API}/api/v1/Resource/{resource_id}",
                     headers=h(jwt), timeout=15)
    r.raise_for_status()
    try:
        return r.json()
    except ValueError:
        raise RuntimeError(
            f"Empty/invalid JSON from Tracker for resource {resource_id} "
            f"(HTTP {r.status_code}, body: {r.text[:80]!r})"
        )

def fix_name_case(name):
    """Convert ALL-CAPS or all-lowercase names to Title Case."""
    if not name:
        return name
    if name == name.upper() or name == name.lower():
        return name.title()
    return name  # Mixed case — leave as-is (e.g. "van der Berg")

def update_resource(jwt, resource_id, job_title, employer, work_type_objs, skills_objs, first_name="", surname=""):
    # Accept either a single dict or a list
    if isinstance(work_type_objs, dict):
        work_type_objs = [work_type_objs]
    payload = {
        "jobTitle":      job_title,
        "currentClient": {"id": -1, "name": employer},
        "workTypes":     work_type_objs,
        "quickSkills":   skills_objs,
    }
    # Fix name casing if needed
    fixed_first = fix_name_case(first_name)
    fixed_sur   = fix_name_case(surname)
    if fixed_first and fixed_first != first_name:
        payload["firstName"] = fixed_first
        print(f"  ✎ Name fix: '{first_name}' → '{fixed_first}'")
    if fixed_sur and fixed_sur != surname:
        payload["surname"] = fixed_sur
        print(f"  ✎ Name fix: '{surname}' → '{fixed_sur}'")
    r = requests.patch(f"{TRACKER_API}/api/v1/Resource/{resource_id}",
                       json=payload, headers=h(jwt), timeout=15)
    return r.status_code, r.text

# ── Process one candidate ──────────────────────────────────────────────────────

def process_one(name, jwt, name_index, skills_lookup, email_cand=None, country_skills_set=None,
                nationality_ids=None, licence_country_ids=None, licence_country_lookup=None,
                force_reprocess=False):
    if country_skills_set is None:
        country_skills_set = set()
    if nationality_ids is None:
        nationality_ids = set()
    if licence_country_ids is None:
        licence_country_ids = set()
    if licence_country_lookup is None:
        licence_country_lookup = {}
    email_item = (email_cand or {}).get("item")  # backward compat
    print(f"\n{'='*60}")
    print(f"  {name}")
    print(f"{'='*60}")

    # 1. Find in Tracker
    print("\n[1/5] Looking up in Tracker...")
    result = find_candidate(name, name_index, jwt)

    NOT_FOUND_LOG = "tracker_not_found.txt"
    if os.path.dirname(NOT_FOUND_LOG):
        os.makedirs(os.path.dirname(NOT_FOUND_LOG), exist_ok=True)

    if result is None:
        print(f"  ✗ Not found in Tracker — logging and continuing.")
        with open(NOT_FOUND_LOG, "a", encoding="utf-8") as f:
            f.write(f"{name}\n")
        # Move to Not Found folder — keeps email visible but out of processing queue
        if email_cand and move_to_done(email_cand, dest_folder=EMAIL_NOT_FOUND_FOLDER):
            print(f"  ↪ Email moved to 'Not Found' folder (needs manual action)")
        return False
    elif isinstance(result, list):
        # Multiple matches — pick the closest one automatically
        print(f"  Multiple matches — picking best fit:")
        name_lower = name.strip().lower()
        best = None
        best_score = -1
        for item in result:
            if isinstance(item, int):
                # Plain integer ID (no name key) — use as fallback only
                if best is None:
                    best = (name_lower, item)
                continue
            if not isinstance(item, tuple) or len(item) < 2:
                continue  # guard against malformed entries
            key, rid = item
            # If rid is a nested list, take the first element
            if isinstance(rid, list):
                if not rid:
                    continue
                rid = rid[0]
            score = sum(p in key.lower() for p in name_lower.split())
            if score > best_score:
                best_score = score
                best = (key, rid)
        if best is None:
            print(f"  ✗ Multiple matches found but none could be scored — skipping")
            return False
        print(f"    → Auto-selected: {best[0]} (ID {best[1]})")
        resource_id = best[1]
    else:
        resource_id = result
        # Fetch the record briefly to confirm the name matches
        try:
            check = requests.get(f"{TRACKER_API}/api/v1/Resource/{resource_id}",
                                 headers=h(jwt), timeout=15).json()
            found_first = (check.get("firstName") or "").strip()
            found_sur   = (check.get("surname") or "").strip()
            found_name  = f"{found_first} {found_sur}".strip()
            print(f"  ✓ Found: {found_name} (ID {resource_id})")

            # ── Robust name-match check ────────────────────────────────────────────
            import unicodedata as _ucn_mc, difflib as _dlmc, re as _remc

            def _acc_mc(s):
                """Lower, strip accents, keep only letters/digits/spaces."""
                s = _ucn_mc.normalize("NFD", s).encode("ascii", "ignore").decode("ascii").lower()
                return _remc.sub(r"[^a-z0-9 ]", "", s)

            # Rank/title words to strip from the search name before comparison
            _MC_RANKS = [
                "air chief marshal", "air vice-marshal", "air vice marshal", "air marshal",
                "air commodore", "wing commander", "squadron leader", "flight lieutenant",
                "group captain", "commodore", "rear admiral", "vice admiral", "admiral",
                "lieutenant general", "major general", "brigadier general", "brigadier",
                "lieutenant colonel", "colonel", "lieutenant commander", "commander",
                "major", "captain", "lieutenant", "sergeant", "corporal",
                "retd", "ret'd", "retired", "dr", "prof", "professor", "sir",
                "dame", "lord", "lady", "mr", "mrs", "ms", "miss",
            ]
            _MC_PARTICLES = {
                "bin", "bint", "bte", "bt", "ibn", "van", "von", "de", "del", "di",
                "al", "el", "ul", "ur", "abu", "abd", "af", "av", "le", "la", "du",
            }

            # Strip rank suffixes from the search name (same logic as find_candidate)
            _search_clean = name.strip()
            _sc_lc = _search_clean.lower()
            for _rk in sorted(_MC_RANKS, key=len, reverse=True):
                if _sc_lc.endswith(" " + _rk):
                    _before = _search_clean[:-(len(_rk) + 1)].strip()
                    if len(_before.split()) >= 1:
                        _search_clean = _before
                        _sc_lc = _search_clean.lower()
                        break

            # Accent-normalised versions of both names
            _s_norm = _acc_mc(_search_clean)
            _f_norm = _acc_mc(found_name)
            _s_parts = _s_norm.split()
            _f_parts = _f_norm.split()

            # Significant parts: exclude particles and single/double-char tokens
            _s_sig = [p for p in _s_parts if len(p) > 2 and p not in _MC_PARTICLES]
            _f_sig = [p for p in _f_parts if len(p) > 2 and p not in _MC_PARTICLES]

            first_variants = FIRST_NAME_VARIANTS.get(_s_parts[0], [_s_parts[0]]) if _s_parts else []

            # Method A: all significant search-name parts appear in found name
            name_ok = bool(_s_sig) and all(p in _f_norm for p in _s_sig)

            # Method B: first-name variant substitution
            if not name_ok and _s_sig:
                for _v in first_variants:
                    _nv = _acc_mc(_v)
                    _rest = _s_sig[1:]
                    if _nv in _f_norm and all(p in _f_norm for p in _rest):
                        name_ok = True
                        break

            # Method C: all significant FOUND-name parts appear in search name
            # Handles: email has extra middle names Tracker doesn't store
            # e.g. "John Robert Smith" (email) → "John Smith" (Tracker)
            if not name_ok and _f_sig:
                if all(p in _s_norm for p in _f_sig):
                    name_ok = True
                # Also try substituting first-name aliases for the FOUND name's first word
                # e.g. found="Mohammad Wahid", search="Mohammed wahid Dar" — Mohammad≠Mohammed but variant matches
                if not name_ok and _f_parts:
                    _f_first_vars = FIRST_NAME_VARIANTS.get(_f_parts[0], [])
                    for _fv in _f_first_vars:
                        _fv_n = _acc_mc(_fv)
                        _f_sig_v = [_fv_n] + [p for p in _f_sig if p != _acc_mc(_f_parts[0])]
                        if all(p in _s_norm for p in _f_sig_v):
                            name_ok = True
                            break

            # Method D: surname parts only match (handles Mohammed/Muhammad/Mohd variants)
            if not name_ok and len(_s_sig) >= 2:
                if all(p in _f_norm for p in _s_sig[1:]):
                    name_ok = True

            # Method E: high fuzzy similarity on accent-normalised concatenated names
            # Handles typos like "Hehsam" vs "Hesham", "Camosa" vs "Canosa"
            if not name_ok and _s_norm and _f_norm:
                _sc2 = _s_norm.replace(" ", "")
                _fc2 = _f_norm.replace(" ", "")
                _sim = _dlmc.SequenceMatcher(None, _sc2, _fc2).ratio()
                if _sim >= 0.87:
                    # Guard: first word and last word must both be reasonably similar
                    _s0 = _s_parts[0] if _s_parts else ""
                    _f0 = _f_parts[0] if _f_parts else ""
                    _sL = _s_parts[-1] if _s_parts else ""
                    _fL = _f_parts[-1] if _f_parts else ""
                    _first_sim = _dlmc.SequenceMatcher(None, _s0, _f0).ratio()
                    _last_sim  = _dlmc.SequenceMatcher(None, _sL, _fL).ratio()
                    if _first_sim >= 0.72 and _last_sim >= 0.72:
                        name_ok = True
            # ──────────────────────────────────────────────────────────────────────

            if not name_ok:
                # Before giving up, try a direct Tracker keyword search for the original name.
                # Catches cases where find_candidate returned a wrong fuzzy/partial match
                # but the real candidate IS in Tracker (e.g. "Mohd Haziq Bin Halim").
                print(f"  ⚠  Name mismatch! '{name}' → '{found_name}' — trying direct search...")
                _retry_id = None
                _stop_words_r = {"bin","binti","bt","ibn","al","el","ul","abu","bte"}
                _sig_words  = [w for w in re.sub(r"[^a-z ]", "", name.lower()).split()
                               if w not in _stop_words_r and len(w) > 1]
                _parts_r = name.split()
                _first_r = _parts_r[0] if _parts_r else name
                _last_r  = _parts_r[-1] if len(_parts_r) > 1 else ""
                # Build search payloads — richest/most targeted first
                _retry_payloads = []
                # Bin-name split: e.g. "Mohd Haziq Bin Halim" → firstName="Mohd Haziq" surname="Halim"
                _bin_idx_r = next((i for i, w in enumerate([p.lower() for p in _parts_r])
                                   if w in _stop_words_r), -1)
                if _bin_idx_r > 0:
                    _pre_bin  = " ".join(_parts_r[:_bin_idx_r])
                    _post_bin = " ".join(_parts_r[_bin_idx_r+1:])
                    _bin_word = _parts_r[_bin_idx_r]
                    if _pre_bin and _post_bin:
                        # e.g. "Mohd Haziq Bin Halim" → firstName="Mohd Haziq" surname="Halim"
                        _retry_payloads.append({"pageSize": 50, "pageNumber": 1,
                                                "firstName": _pre_bin, "surname": _post_bin})
                        # Also try with particle in surname: firstName="Mohd Haziq" surname="Bin Halim"
                        _retry_payloads.append({"pageSize": 50, "pageNumber": 1,
                                                "firstName": _pre_bin, "surname": f"{_bin_word} {_post_bin}"})
                # firstName (first word) + surname (last word)
                if _first_r and _last_r:
                    _retry_payloads.append({"pageSize": 50, "pageNumber": 1,
                                            "firstName": _first_r, "surname": _last_r})
                # Surname only
                if _last_r and len(_last_r) > 2:
                    _retry_payloads.append({"pageSize": 50, "pageNumber": 1,
                                            "surname": _last_r})
                # First name only (catches "Gokulan S" for "Gokulan Sivakumar")
                if _first_r and len(_first_r) > 2:
                    _retry_payloads.append({"pageSize": 50, "pageNumber": 1,
                                            "firstName": _first_r})
                # Full name keyword/searchTerm
                _retry_payloads.append({"pageSize": 50, "pageNumber": 1, "keyword": name})
                _retry_payloads.append({"pageSize": 50, "pageNumber": 1, "searchTerm": name})
                for _rp in _retry_payloads:
                    try:
                        _rr = requests.post(f"{TRACKER_API}/api/v1/Resource/Search",
                                            json=_rp, headers=h(jwt), timeout=15)
                    except Exception as _re_err:
                        print(f"  ⛑ Retry search error: {_re_err}")
                        continue
                    if _rr.status_code != 200:
                        continue
                    try:
                        _rr_data = _rr.json()
                    except Exception:
                        continue
                    _rr_items = _rr_data if isinstance(_rr_data, list) \
                                else (_rr_data.get("items") or _rr_data.get("results") or [])
                    if not _rr_items:
                        continue
                    for _item in _rr_items:
                        _rf  = (_item.get("firstname") or _item.get("firstName") or "").strip()
                        _rs  = (_item.get("surname") or "").strip()
                        _rid = _item.get("resourceId") or _item.get("id")
                        if not _rid or _rid == resource_id:
                            continue  # skip the already-rejected match
                        _comb = re.sub(r"[^a-z ]", "", f"{_rf} {_rs}".lower())
                        # Accept: all significant words appear in result name
                        _word_match = bool(_sig_words and all(w in _comb for w in _sig_words))
                        # Accept: initial-surname match ("Gokulan S" for "Gokulan Sivakumar")
                        _init_match = (len(_rs) == 1 and len(_sig_words) >= 2
                                       and _rs.lower() == _sig_words[-1][0]
                                       and _rf.lower().startswith(_sig_words[0]))
                        if _word_match or _init_match:
                            print(f"  ~ Direct search found: \'{_rf} {_rs}\' (ID {_rid})")
                            _retry_id = _rid
                            break
                    if _retry_id:
                        break

                if _retry_id:
                    print(f"  ~ Retrying with direct search result (ID {_retry_id})")
                    resource_id = _retry_id
                    # Fall through: step 2 will fetch this record
                else:
                    with open(NOT_FOUND_LOG, "a", encoding="utf-8") as f:
                        f.write(f"MISMATCH: looking for '{name}', found '{found_name}' (ID {resource_id})\n")
                    return False
        except Exception:
            print(f"  ✓ Found resource ID {resource_id}")

    # Store tracker resource_id in the email candidate dict so the caller can
    # persist it to tracker_processed.json for later auditing.
    if email_cand is not None and isinstance(email_cand, dict):
        email_cand["tracker_id"] = resource_id

    # 2. Current record
    print("\n[2/5] Current Tracker record...")
    try:
        rec = get_resource(jwt, resource_id)
    except Exception as e:
        print(f"  ✗ Could not fetch record: {e}")
        return False

    rec_first    = rec.get("firstName", "")
    rec_sur      = rec.get("surname", "")
    rec_job      = rec.get("jobTitle") or ""
    rec_employer = (rec.get("currentClient") or {}).get("name") or ""
    # Fix double-encoded UTF-8 in employer name stored in Tracker
    # e.g. "AÃƒÂ©reas" → "Aéreas" (Latin-1 bytes re-decoded as UTF-8)
    if rec_employer and any(c in rec_employer for c in ("Ã", "Â", "Ä", "Ö", "Ü")):
        try:
            rec_employer = rec_employer.encode("latin-1").decode("utf-8")
        except (UnicodeDecodeError, UnicodeEncodeError):
            pass
    print(f"  Name:      {rec_first} {rec_sur}")
    print(f"  Job title: {rec_job or '(empty)'}")
    print(f"  Employer:  {rec_employer or '(empty)'}")
    print(f"  Work type: {rec.get('workTypes') or '(empty)'}")
    print(f"  Skills:    {[s.get('name') for s in rec.get('quickSkills', [])] or '(empty)'}")

    # ── FAST PATH: profile has job+worktype, only issue is a regional term skill ──
    # Note: NATIONALITY_ADJECTIVES removed from this check — Tracker now stores
    # nationalities as demonyms ("British", "Indian") so those are correct.
    # Don't re-parse the CV — just fix the skill name and leave everything else.
    if rec_job and rec.get("workTypes"):
        bad_skills = [s for s in rec.get("quickSkills", [])
                      if (s.get("name") or "").strip().lower() in REGIONAL_TERMS]
        if bad_skills:
            fixed_skills = []
            for s in rec.get("quickSkills", []):
                sname = (s.get("name") or "").strip()
                sname_lower = sname.lower()
                if sname_lower in REGIONAL_TERMS:
                    pass  # remove regional terms entirely
                else:
                    fixed_skills.append(s)
            # Verify the result actually passes completeness — if not (e.g. cabin
            # crew still has >1 country after removing regional terms), fall through
            # to the full CV-download path instead of accepting a broken fast-path fix.
            _temp_rec = dict(rec)
            _temp_rec["quickSkills"] = fixed_skills
            if not is_profile_complete(_temp_rec, country_skills_set=country_skills_set):
                pass  # fall through to full CV processing below
            else:
                print(f"\n[FAST] Profile is otherwise complete — cleaning up skills.")
                print(f"  Was:  {[s.get('name') for s in rec.get('quickSkills', [])]}")
                print(f"  Now:  {[s.get('name') for s in fixed_skills]}")
                if DRY_RUN:
                    print("  ⚠  DRY RUN — changes NOT written to Tracker.")
                else:
                    status, resp = update_resource(jwt, resource_id,
                                                   rec_job, rec_employer,
                                                   rec.get("workTypes"), fixed_skills,
                                                   first_name=rec_first, surname=rec_sur)
                    if 200 <= status < 300:
                        print(f"  ✓ Updated — ID {resource_id}")
                        if email_cand:
                            if move_to_done(email_cand):
                                print(f"  ✓ Email moved to '{EMAIL_DONE_FOLDER}'")
                    else:
                        print(f"  ✗ Update failed: {status} — {resp}")
                return True

    # 3. Download CV
    print("\n[3/5] Downloading CV...")
    cv_text = get_cv_text(jwt, resource_id)
    no_cv = False
    if cv_text == "CV_EXISTS_NO_TEXT":
        # Documents are attached in Tracker but couldn't be downloaded.
        # Mark as done so we don't retry on every run — CV parsing is best-effort.
        print("  ⚠  CV exists in Tracker but could not be downloaded — marking done, moving email.")
        if email_cand:
            move_to_done(email_cand)
        return True
    if not cv_text or not cv_text.strip():
        if not SEND_CV_REQUESTS:
            _display_name = f"{rec.get('firstName','')} {rec.get('surname','')}".strip() or name
            print(f"  ⚠  No CV found — marking done, moving email ({_display_name}).")
            _NO_CV_NAMES.append(_display_name)
            # Add to pending_cv.json so the daily no-CV report can track them.
            # We do NOT email the candidate — Emily reviews the list and emails manually.
            _cand_email = ((rec.get("contactDetails") or {}).get("email")
                           or rec.get("email") or "")
            save_pending_cv(resource_id, _display_name, _cand_email)
            if email_cand:
                move_to_done(email_cand)
            return True
        print("  ⚠  No CV found — auto-emailing candidate to request CV.")
        choice = "e"
        if False:  # manual paste mode disabled in auto-run
            lines = []
            while True:
                line = input()
                if line.strip().upper() == "END":
                    break
                lines.append(line)
            cv_text = "\n".join(lines)
            if not cv_text.strip():
                print("  Nothing pasted — treating as no CV.")
                no_cv = True
        elif choice == "e":
            no_cv = True
        else:
            print("  Skipped.")
            return False
    else:
        print(f"  ✓ CV downloaded ({len(cv_text)} characters)")

    # ── Hard protection: manually-corrected profiles must never be auto-updated ──
    if not force_reprocess and resource_id in PROTECTED_IDS:
        print(f"\n  🔒 ID {resource_id} is in PROTECTED_IDS — skipping (manually corrected).")
        if email_cand:
            if move_to_done(email_cand):
                print(f"  ✓ Email moved to '{EMAIL_DONE_FOLDER}'")
        return True

    # ── Auto-skip: profile already complete and CV is present ─────────────────
    if not no_cv and not force_reprocess and is_profile_complete(rec, country_skills_set=country_skills_set):
        print("\n  ✓ Profile already complete (job title, work type & skills all set).")
        print("  → Auto-skipping.")
        if email_cand:
            if move_to_done(email_cand):
                print(f"  ✓ Email moved to '{EMAIL_DONE_FOLDER}'")
        return True

    # ── No CV: email candidate and add to watchlist ────────────────────────────
    if no_cv:
        candidate_email = (rec.get("contactDetails") or {}).get("email") or rec.get("email") or ""
        candidate_name  = f"{rec.get('firstName','')} {rec.get('surname','')}".strip()

        if not candidate_email:
            print("  ⚠  No email address on their Tracker profile — cannot send email.")
            print("  Skipped.")
            return False

        print(f"\n  Will email: {candidate_email}")
        print(f"  Subject:    CV Request — {candidate_name}")
        confirm = "y"  # auto-confirm in auto-run mode

        sent = send_cv_request_email(candidate_email, candidate_name)
        if sent:
            save_pending_cv(resource_id, candidate_name, candidate_email)
            print(f"  ✓ Email sent to {candidate_email}")
            print(f"  ✓ Added to watchlist — profile will be deleted in 5 days if no CV received")
        else:
            print(f"  ✗ Email failed — candidate not added to watchlist")
        return sent

    # 4. AI parsing
    print("\n[4/5] Parsing CV with AI...")
    # Combine email name + Tracker record name so both sets of words are excluded from
    # title parsing — prevents "MOHAMMED" (Tracker first name) being treated as a job title
    # when the email name is "Muhammad Tariq" (different format).
    tracker_full_name = f"{rec.get('firstname') or rec.get('firstName') or ''} {rec.get('surname') or ''}".strip()
    combined_name = f"{name} {tracker_full_name}".strip()
    # ── OCR quality gate: reject garbage text before parsing ────────────────────
    # If OCR produced very little real text, or the ratio of non-alpha chars is too
    # high (random caps, noise glyphs), treat it like no CV rather than saving garbage.
    def _ocr_quality_ok(text):
        if not text or not text.strip():
            return False
        _clean = re.sub(r"\s+", " ", text.strip())
        _alpha = sum(1 for c in _clean if c.isalpha() or c.isspace())
        _total = max(len(_clean), 1)
        _ratio = _alpha / _total
        if len(_clean) < 30:               # fewer than 30 real characters = useless
            return False
        # Aviation CVs often have lots of numbers (hours, dates, aircraft types).
        # If the text contains recognisable aviation keywords, accept it even if
        # the alpha ratio is below the normal threshold.
        _AVIATION_KWS = re.compile(
            r"\b(captain|officer|pilot|airline|airways|aviation|airbus|boeing|"
            r"embraer|bombardier|atpl|cpl|ppl|b737|b747|b777|b787|a320|a330|"
            r"a340|a350|a380|a220|crj|erj|q400|atr|dash|hours|landings|"
            r"cycles|flight|command|co-pilot|copilot|first officer|second officer|"
            r"cabin crew|purser|engineer|maintenance|technician)\b",
            re.IGNORECASE
        )
        if _AVIATION_KWS.search(_clean) and _ratio >= 0.35:
            pass  # accept: recognisable aviation CV even if number-heavy
        elif _ratio < 0.45:                # below 45% alpha/space = garbage
            return False
        # Detect OCR word-salad: avg word length outside 2-15 chars = likely noise
        _words = [w for w in _clean.split() if w]
        if _words:
            _avg = sum(len(w) for w in _words) / len(_words)
            if _avg > 15 or _avg < 2:
                return False
        return True

    if not _ocr_quality_ok(cv_text):
        print("  ⚠  OCR text quality too low to parse reliably — preserving existing Tracker data.")
        # Keep existing skills/job title — do not overwrite with garbage
        _existing_qs = rec.get("quickSkills") or []
        _existing_wt = rec.get("workTypes") or []
        _existing_jt = rec.get("jobTitle") or ""
        _existing_emp = (rec.get("currentClient") or {}).get("name") or rec.get("employer") or ""
        if _existing_qs or _existing_wt or _existing_jt:
            print("  ℹ  Existing profile data retained (no changes made).")
            if email_cand:
                move_to_done(email_cand)
            return True
        # No existing data either — nothing we can do
        if email_cand:
            move_to_done(email_cand)
        return True

    # ── Fix common OCR word-splits before AI parsing ──────────────────────────
    # Scanned PDFs sometimes split words across lines: "Inst ructor" → "Instructor"
    _OCR_SPLITS = [
        (r"\bInst\s+ructor",   "Instructor"),
        (r"\bTech\s+nician",   "Technician"),
        (r"\bEng\s+ineer",     "Engineer"),
        (r"\bMaint\s+enance",  "Maintenance"),
        (r"\bMan\s+ager",      "Manager"),
        (r"\bDir\s+ector",     "Director"),
        (r"\bOper\s+ations",   "Operations"),
        (r"\bAss\s+istant",    "Assistant"),
        (r"\bCoord\s+inator",  "Coordinator"),
        (r"\bSup\s+ervisor",   "Supervisor"),
        (r"\bAdmin\s+istrator","Administrator"),
    ]
    for _pat, _fix in _OCR_SPLITS:
        cv_text = re.sub(_pat, _fix, cv_text, flags=re.IGNORECASE)

    result = parse_cv(cv_text, combined_name)
    if not result:
        print("  ✗ Could not parse. Skipping.")
        return False

    job_title    = result.get("job_title", "")
    # Never allow a country/nationality name to be a job title
    _jt_lc = job_title.strip().lower()
    if _jt_lc and (
        _jt_lc in KNOWN_COUNTRY_NAMES_EXTRA
        or _jt_lc in {v.lower() for v in COUNTRY_ALIASES.values()}
        or _jt_lc in NATIONALITY_ADJECTIVES
    ):
        print(f"  \u2139  Rejected country/adjective as job title: '{job_title}'")
        job_title = ""
    # Job title plausibility gate: if extracted title doesn't look like a real
    # professional title AND Tracker already has one, keep the Tracker value.
    if job_title and rec_job and not _is_plausible_job_title(job_title):
        print(f"  \u2139  Kept existing title '{rec_job}' "
              f"(extracted '{job_title}' doesn't look like a job title)")
        job_title = rec_job
    employer     = result.get("current_employer", "")
    # Safety net: if CV extracted something that is clearly a job title, discard it.
    # Catches cases where the fallback employer regex matches "Airline Pilot" etc.
    _EMPLOYER_IS_JOB_TITLE = re.compile(
        r"^(airline pilot|commercial pilot|senior first officer|first officer|"
        r"second officer|junior first officer|relief captain|training captain|"
        r"check captain|line captain|co-pilot|copilot|captain|pilot|"
        r"cabin crew|senior cabin crew|flight attendant|purser|chief purser|"
        r"aircraft engineer|maintenance engineer|avionics engineer|"
        r"flight engineer|ground engineer|certifying engineer)\b",
        re.IGNORECASE
    )
    if employer and _EMPLOYER_IS_JOB_TITLE.match(employer.strip()):
        print(f"  ℹ  Discarded employer '{employer}' — looks like a job title")
        employer = ""
    wtype_key    = result.get("work_type", "").lower().strip()
    skill_names  = result.get("skills", [])
    nationality  = result.get("nationality", "")
    nat_explicit = result.get("nat_explicit", False)  # True = found via explicit keyword

    # Override work type if existing Tracker title clearly shows engineering
    # (catches mismatch where candidate was registered as Flight Deck but is maintenance)
    _rec_lc = (rec_job or "").lower()
    _wtype_was_overridden = False  # True when we override based on Tracker title
    _ENG_TITLE_SIGNALS = [
        "b1 certif","b2 certif","cat b1","cat b2","b1/b2","b1b2",
        "lame","licensed aircraft maintenance","aircraft maintenance engineer",
        "base maintenance","line maintenance","mro engineer","avionics engineer",
        "a&p mechanic","powerplant mechanic","certifying staff","certifying support",
        # Broader aviation engineering titles
        "aerodynamic","aerospace engineer","systems engineer","flight engineer",
        "structural engineer","propulsion engineer","aircraft engineer",
        "aviation engineer","maintenance engineer","quality engineer",
        "design engineer","test engineer","stress engineer","camo",
        # Simple rule: any job title containing "engineer" is Engineering
        "engineer",
    ]
    if any(kw in _rec_lc for kw in _ENG_TITLE_SIGNALS):
        if "engineering" not in wtype_key:
            print(f"  ℹ  Work type overridden to Engineering (Tracker title: '{rec_job}')")
            wtype_key = "engineering"
            _wtype_was_overridden = True

    # If parse_cv detected Flight Deck but the existing Tracker job title is clearly
    # a ground/logistics/support role, downgrade to the appropriate work type.
    # Prevents aircraft types in a CV from overriding someone's actual current role.
    _NON_PILOT_TITLE_SIGNALS = [
        "material controller","materials controller","shipping","receiving",
        "supply chain","procurement","logistics","warehouse","stores",
        "inventory","purchasing","buyer","planner","scheduler",
        "dispatcher","load controller","ground handling","ramp agent",
        "ramp supervisor","ground crew","cargo agent","cargo handler",
    ]
    if wtype_key == "flight deck" and any(kw in _rec_lc for kw in _NON_PILOT_TITLE_SIGNALS):
        # Determine better work type from the title
        if any(kw in _rec_lc for kw in ["material controller","materials controller",
                                          "supply chain","procurement","logistics",
                                          "warehouse","stores","inventory","purchasing",
                                          "buyer","planner","scheduler"]):
            print(f"  ℹ  Work type overridden to Engineering (non-pilot Tracker title: '{rec_job}')")
            wtype_key = "engineering"
        else:
            print(f"  ℹ  Work type overridden to Operations (non-pilot Tracker title: '{rec_job}')")
            wtype_key = "operations"
        _wtype_was_overridden = True

    # Override work type if the existing Tracker title clearly says cabin crew
    # but parse_cv scored it as something else (e.g. head office / flight deck)
    # (runs after engineering override so cabin crew still wins if both match)
    if any(kw in _rec_lc for kw in ["cabin crew","flight attendant","purser",
                                     "cabin manager","inflight manager","air cabin"]):
        if "cabin" not in wtype_key:
            print(f"  ℹ  Work type overridden to Cabin Crew (Tracker title: '{rec_job}')")
            wtype_key = "cabin crew"
            _wtype_was_overridden = True

    # ── Fallback: if parse_cv + title overrides couldn't determine work type,
    # inherit the work type from the existing Tracker record (via work_type_objs).
    # This handles candidates like engineers whose CVs don't contain clear work-type
    # signals — without this, the engineering aircraft-type strip never fires.
    if not wtype_key and work_type_objs:
        _wtype_ids = {w.get("id") for w in work_type_objs}
        if 471 in _wtype_ids:
            wtype_key = "engineering"
        elif 472 in _wtype_ids:
            wtype_key = "flight deck"
        elif 469 in _wtype_ids:
            wtype_key = "cabin crew"
        elif 470 in _wtype_ids or 473 in _wtype_ids:
            wtype_key = "operations"
        if wtype_key:
            print(f"  ℹ  Work type inferred from existing Tracker record: '{wtype_key}'")

    # When work type was overridden away from flight deck, strip pilot-specific
    # skills that parse_cv extracted — they don't belong on this profile.
    if _wtype_was_overridden and "flight deck" not in wtype_key:
        _PILOT_ONLY_SKILLS = (
            {p.lower() for p in POSITIONS}
            | {"icao", "easa", "faa", "jaa", "tri", "tre", "sfi", "ire", "irr"}
        )
        _before = len(skill_names)
        skill_names = [s for s in skill_names
                       if s.lower() not in _PILOT_ONLY_SKILLS
                       and s not in AIRCRAFT_TYPES]
        if len(skill_names) < _before:
            print(f"  ℹ  Stripped {_before - len(skill_names)} pilot-specific skill(s) "
                  f"(work type overridden to {wtype_key})")

    def resolve_work_types(wtype_str):
        """Parse a comma-separated work type string into a list of Tracker objects."""
        objs = []
        for part in wtype_str.split(","):
            key = part.strip().lower()
            obj = WORK_TYPES.get(key)
            if not obj:
                for k, v in WORK_TYPES.items():
                    if k in key or key in k:
                        obj = v
                        break
            if obj and obj not in objs:
                objs.append(obj)
        return objs or [{"id": 0, "name": wtype_str.strip()}]

    # Map work type (may be multiple, comma-separated from parse_cv)
    work_type_objs = resolve_work_types(wtype_key)

    # Map skills
    skills_objs = resolve_skills(skill_names, skills_lookup)

    # ── Preserve existing Tracker values when parsed result is empty or worse ──
    # Never blank out a field that already has a good value in Tracker.

    # Keep existing job title unless it's clearly non-aviation for a known aviation role.
    # For cabin crew / flight deck: if the stored title doesn't look aviation-related,
    # use the CV-parsed title (or a sensible default) instead.
    _AVIATION_RE = re.compile(
        r"\b(cabin|crew|flight|attendant|purser|steward|captain|officer|pilot|"
        r"engineer|maintenance|aviation|airline|airways|dispatcher|controller|"
        r"instructor|ramp|ground|airport|inflight|in-flight|cargo)\b"
        r"|(?<![A-Za-z])F/?O(?![A-Za-z])"       # F/O or FO abbreviation
        r"|(?<![A-Za-z])S/?O(?![A-Za-z])"        # S/O (Second Officer)
        r"|(?<![A-Za-z])C/?O(?![A-Za-z])"        # C/O (Co-pilot)
        r"|(?<![A-Za-z])CPT(?![A-Za-z])"         # CPT (Captain)
        r"|(?<![A-Za-z])SFO(?![A-Za-z])",        # SFO (Senior First Officer)
        re.IGNORECASE)
    _ACTYPE_RE = re.compile(
        r"\b([AB]\d{3}|EMB\d{3}|CRJ\d*|ATR\d*|ERJ[-\d]*|E\d{3}|MD\d{2}|B74[78]|Q\d00|Dash|DHC|Saab|AW\d{3}|EC\d{3}|S\d{2}|Bell\d{3}|PC.?12|King Air|Caravan)\b",
        re.IGNORECASE)
    # An existing title is "aviation" if it has an aviation keyword OR an aircraft type
    if rec_job and (_AVIATION_RE.search(rec_job) or _ACTYPE_RE.search(rec_job)):
        # Existing title is aviation — keep it, UNLESS it has no aircraft prefix and
        # the CV-parsed title has one (Emily's rule: aircraft always in front of title).
        if not _ACTYPE_RE.search(rec_job) and job_title and _ACTYPE_RE.search(job_title):
            pass  # use the CV-parsed title which already has the aircraft prefix
        else:
            job_title = rec_job
    elif rec_job and not (_AVIATION_RE.search(rec_job) or _ACTYPE_RE.search(rec_job)):
        # Existing title is not aviation (e.g. "NURSE", "PERFUME ADVISOR").
        # Only keep CV-parsed title if it IS aviation — otherwise use work-type default.
        if job_title and _AVIATION_RE.search(job_title):
            pass  # use CV-parsed aviation title
        else:
            # Neither existing nor CV title is aviation
            if "cabin" in wtype_key:
                # Cabin crew with no aviation background: non-aviation Tracker title is meaningless
                job_title = "Unknown"
                if rec_job:
                    print(f"  ℹ  Non-aviation title '{rec_job}' replaced with 'Unknown' (cabin crew candidate, no aviation background)")
            else:
                # Management / operations etc.: keep the real Tracker title
                # But reject it if it looks like the candidate's own name.
                _jt_rec_words = [w.lower() for w in (rec_job or "").split() if len(w) > 2]
                _jt_surname   = (name or "").split()[-1].lower()
                _rec_is_name  = (
                    _jt_surname and len(_jt_surname) > 3 and _jt_rec_words
                    and any(_jt_surname == w or w.startswith(_jt_surname) for w in _jt_rec_words)
                )
                if _rec_is_name:
                    print(f"  ℹ  Rejected Tracker title '{rec_job}' — looks like candidate's name")
                    job_title = ""
                else:
                    job_title = rec_job  # don't overwrite with CV garbage (e.g. a country name)
            # flight deck: job_title already built from position+aircraft by parse_cv
    elif not job_title or len(job_title) < 4:
        # Also reject rec_job if it's the candidate's name
        _jt_rec_words2 = [w.lower() for w in (rec_job or "").split() if len(w) > 2]
        _jt_surname2   = (name or "").split()[-1].lower()
        _rec_is_name2  = (
            _jt_surname2 and len(_jt_surname2) > 3 and _jt_rec_words2
            and any(_jt_surname2 == w for w in _jt_rec_words2)
        )
        if rec_job and not _rec_is_name2:
            job_title = rec_job
        elif _rec_is_name2:
            print(f"  ℹ  Rejected Tracker title '{rec_job}' — looks like candidate's name")
            job_title = ""

    # Flight deck: if job title is aviation but has no aircraft prefix, add one from skills
    if "flight" in wtype_key:
        if job_title and not _ACTYPE_RE.search(job_title):
            # Look for an aircraft type in the existing skills to prepend
            _ac_prefix = next(
                (s.get("name","") for s in skills_objs
                 if (s.get("name","") or "") in AIRCRAFT_TYPES),
                ""
            )
            if _ac_prefix and _ac_prefix.lower() not in job_title.lower():
                job_title = f"{_ac_prefix} {job_title}"
                print(f"  ℹ  Prepended aircraft type to job title: '{job_title}'")

    # Keep existing title when neither CV nor existing is aviation (prevents 'Nigeria' etc.)
    # and existing is non-empty — but guard against rec_job being the candidate's own name.
    if not job_title and rec_job:
        _ft_words = [w.lower() for w in rec_job.split() if len(w) > 2]
        _ft_sname = (name or "").split()[-1].lower()
        _ft_is_name = (
            _ft_sname and len(_ft_sname) > 3 and _ft_words
            and any(_ft_sname == w or w.startswith(_ft_sname) for w in _ft_words)
        )
        if not _ft_is_name:
            job_title = rec_job
        else:
            print(f"  ℹ  Final fallback: rejected '{rec_job}' — looks like candidate's name")
    # Title-case if all-caps (PDF/CV formatting artefact)
    if job_title and job_title == job_title.upper() and len(job_title) > 3:
        job_title = job_title.title()
    # Reject garbled titles containing a lone uppercase letter ('C orrespondence')
    if job_title and any(re.match(r'^[A-Z]$', w) for w in job_title.split()):
        print(f"  ℹ  Rejected garbled job title: '{job_title}' — reverting to existing")
        job_title = rec_job or ""

    # Final bad-title filter — catches garbage from CV parse OR already-bad Tracker records.
    # Applied AFTER rec_job fallback, so it also cleans stale Tracker data.
    if job_title:
        _FINAL_BAD_TITLE_RE = re.compile(
            r"^("
            r"declaration|pilot\s+ratings?:?|pilot\s+rating\b|"  # CV/Tracker header noise
            r"informations?\s+personnelles?|donn[eé]es?\s+personnelles?|"  # French
            r"datos?\s+personales?|informaci[oó]n\s+personal|"  # Spanish
            r"persönliche\s+(?:daten|angaben)|"  # German
            r"informazioni\s+personali|dati\s+personali|"  # Italian
            r"informações\s+pessoais|dados\s+pessoais|"  # Portuguese
            r"(?:contribute|contribut|seek|provid|leverag|utiliz|"
            r"pursu|aspir|obtain|demonstrat|maximiz|ensur|apply|appli)"
            r"\w*\s+to\b|"  # objective-statement openers
            r"seu\s+\w|hoya\s+\w|"  # OCR garbage prefixes
            r"further\s+(more|information|details?|to\b)|"  # transition word fragments
            r"resume\b|curriculum\s+vitae\b|c\.v\.\b|"  # CV document headers
            r"\w*\s+cours[e]?\b|"  # training course names (e.g. "Differences Course")
            r"recurrent\b|refresher\b|conversion\b|differences?\b"  # standalone course words
            r")",
            re.IGNORECASE
        )
        _SCHOOL_FINAL_RE = re.compile(
            r"\b(patts|college\s+of\s+aeronaut|university\s+of\s+aero|"
            r"college\s+of\s+aviation|aviation\s+(?:colleg|academ|universit|instit)|"
            r"aeronautical\s+(?:universit|colleg|academ)|polytechnic)\b",
            re.IGNORECASE
        )
        # Strip trailing " - Qualifier/Company" suffix before word-count check.
        # e.g. "Aircraft Systems & Avionics Engineer - Airhub Aviation"
        #    → "Aircraft Systems & Avionics Engineer"
        # IMPORTANT: only apply to CV-extracted titles. If job_title == rec_job we are
        # using the existing Tracker value as fallback — don't strip it further.
        _jt_stripped = re.sub(r'\s+[-–]\s+\S+.*$', '', job_title).strip()
        if _jt_stripped and _jt_stripped != job_title and job_title != (rec_job or ""):
            job_title = _jt_stripped
        _jt_final_words = job_title.split()
        _title_is_garbage = (
            _FINAL_BAD_TITLE_RE.match(job_title)
            or _SCHOOL_FINAL_RE.search(job_title)
            or len(_jt_final_words) > 7
        )
        if _title_is_garbage:
            print(f"  ℹ  Cleared bad job title: '{job_title}'")
            job_title = ""

    # If work type was overridden away from flight deck, a CV title like "Pilot - EMB190"
    # or "B737 Captain" must not survive. Use the existing Tracker title instead.
    if _wtype_was_overridden and "flight deck" not in wtype_key and rec_job:
        if _ACTYPE_RE.search(job_title or "") or "pilot" in (job_title or "").lower():
            print(f"  ℹ  Reverting pilot CV title '{job_title}' → Tracker title '{rec_job}'")
            job_title = rec_job

    # Plausibility gate: if the CV extraction found something but it doesn't look like
    # a real company name, AND Tracker already has a good employer, keep the Tracker value.
    # This is the main defence against regex returning sentence fragments, course names,
    # cities, industry descriptions, etc. that clean_employer didn't catch.
    _AIRLINE_SIGNAL_RE = re.compile(
        # \bair\b covers "Air Haifa", "Air France", "Nok Air", "Air Asia" etc.
        # Previously used air\s+\w which broke due to the trailing \b not matching mid-word.
        r"\b(airlines?|airways|\bair\b|aviation\s+(?:limited|ltd|llc|plc|inc|group|services)|"
        r"helicopters?|aero(?:lines?|flot|mexico|peru|logic)|"
        r"lufthansa|emirates|ryanair|easyjet|wizz|flydubai|flyadeal|indigo|"
        r"airasia|jetblue|southwest|united\s+airlines|delta\s+air)\b",
        re.IGNORECASE
    )
    if employer and rec_employer:
        _validated_rec_emp = _clean_employer_simple(rec_employer)
        if _validated_rec_emp:
            # Hard rule: if existing employer is clearly an airline, don't replace with
            # a non-airline name (training school, institute, etc.)
            if (_AIRLINE_SIGNAL_RE.search(rec_employer)
                    and not _AIRLINE_SIGNAL_RE.search(employer)):
                print(f"  ℹ  Kept airline employer '{_validated_rec_emp}' "
                      f"(extracted '{employer}' is not an airline)")
                employer = _validated_rec_emp
            elif not _is_plausible_employer(employer):
                print(f"  ℹ  Kept existing employer '{_validated_rec_emp}' "
                      f"(extracted '{employer}' doesn't look like a company name)")
                employer = _validated_rec_emp
    elif not employer and rec_employer:
        # No employer found in CV — validate the stored employer before using it as
        # fallback (previous runs may have written garbage).
        _validated_rec_emp = _clean_employer_simple(rec_employer)
        if _validated_rec_emp:
            employer = _validated_rec_emp
    # Final safety: if extracted employer fails plausibility and there's no existing
    # employer to fall back to, clear it rather than saving garbage to Tracker.
    if employer and not rec_employer and not _is_plausible_employer(employer):
        print(f"  ℹ  Cleared implausible employer '{employer}' (no existing employer to fall back to)")
        employer = ""
    # Prefer CV employer (most recent); fall back to Tracker only when CV found nothing

    # Resolve cabin crew ↔ flight deck conflicts only.
    # For all other work type combinations (engineering, management, operations, etc.)
    # trust what parse_cv returned — don't revert to stale Tracker data.
    # Exception: if _wtype_was_overridden, the override already won; skip entirely.
    _FLIGHT_DECK_ID = 472
    _CABIN_CREW_ID  = 469
    if rec.get("workTypes") and not _wtype_was_overridden:
        email_wtype_ids    = {w.get("id") for w in work_type_objs}
        existing_wtype_ids = {w.get("id") for w in rec.get("workTypes")}
        if _CABIN_CREW_ID in email_wtype_ids:
            if _FLIGHT_DECK_ID in existing_wtype_ids:
                # Existing is Flight Deck — never downgrade to Cabin Crew based on CV text alone.
                # A pilot's CV will often mention cabin crew terms (safety demos, crew duties etc.)
                # and should remain classified as Flight Deck.
                work_type_objs = rec.get("workTypes")
            else:
                # Parse says cabin crew → strip any Flight Deck from existing
                corrected = [w for w in rec.get("workTypes") if w.get("id") != _FLIGHT_DECK_ID]
                work_type_objs = corrected if corrected else work_type_objs
        elif _FLIGHT_DECK_ID in email_wtype_ids:
            # Parse says flight deck → strip any Cabin Crew from existing
            corrected = [w for w in rec.get("workTypes") if w.get("id") != _CABIN_CREW_ID]
            work_type_objs = corrected if corrected else work_type_objs
        elif _CABIN_CREW_ID in existing_wtype_ids or _FLIGHT_DECK_ID in existing_wtype_ids:
            # Existing is cabin crew or flight deck but parse disagrees — keep existing
            # (registration category is authoritative for these two)
            work_type_objs = rec.get("workTypes")
        # else: neither existing nor parsed is cabin crew/flight deck
        # → trust parse_cv (engineering, management, operations etc. update freely)

    # Merge CV-extracted skills with existing Tracker skills.
    # Always keep existing skills; add new ones from CV on top.
    # Also convert any nationality adjectives in existing skills to country names.
    existing_qs = rec.get("quickSkills") or []
    existing_fixed = []
    _all_country_lower = (KNOWN_COUNTRY_NAMES_EXTRA
                          | {v.lower() for v in COUNTRY_ALIASES.values()})
    for s in existing_qs:
        # Strip leading/trailing whitespace from skill names (Tracker sometimes stores "Main Crew ")
        if s.get("name"):
            s = dict(s)
            s["name"] = s["name"].strip()
        sname = (s.get("name") or "").strip()
        # Detect comma-joined country skills (e.g. "Colombia,Italy") — split and resolve each part
        if "," in sname and sname.lower() not in skills_lookup:
            parts = [p.strip() for p in sname.split(",") if p.strip()]
            all_countries = all(p.lower() in _all_country_lower for p in parts)
            if all_countries and len(parts) > 1:
                print(f"  ℹ  Split corrupted multi-country skill '{sname}' → {parts}")
                for part in parts:
                    resolved = resolve_skills([part], skills_lookup)
                    existing_fixed.extend(resolved if resolved else [{"id": 0, "name": part}])
                continue
        alias = COUNTRY_ALIASES.get(sname.lower())
        if alias:
            resolved = resolve_skills([alias], skills_lookup)
            existing_fixed.extend(resolved if resolved else [{"id": 0, "name": alias}])
        elif sname:
            existing_fixed.append(s)
    # Merge: start with existing (nationality-fixed), add any newly extracted skills not already present
    existing_names_lower = {(s.get("name") or "").lower() for s in existing_fixed}
    # Don't add a position from CV if the existing Tracker skills already have one —
    # the existing position is authoritative (prevents Captain being added for a First Officer)
    _position_set = {p.lower() for p in POSITIONS}
    _has_existing_position = any((s.get("name") or "").lower() in _position_set
                                  for s in existing_fixed)
    for s in skills_objs:
        sname_lc = (s.get("name") or "").lower()
        if sname_lc in existing_names_lower:
            continue
        if sname_lc in _position_set and _has_existing_position:
            continue  # don't add conflicting position from CV
        existing_fixed.append(s)
    if existing_fixed:
        skills_objs = existing_fixed
    # If still nothing, keep whatever we extracted from the CV
    # (skills_objs already contains the CV result if no existing skills)

    # ── Operations/Management: nationality only — strip flight-deck skills ────────
    # Licence authorities (ICAO, EASA, FAA) and positions (Captain, FO) are
    # flight-deck-specific and should not appear on ops/management profiles.
    _FLIGHT_DECK_ONLY_SKILLS = {
        "icao", "easa", "faa", "casa", "sacaa", "caas", "gcaa", "dgca",
        "tcca", "hcaa", "dgac", "caac", "jcab",
        "captain", "first officer", "second officer", "co-pilot", "copilot",
        "purser", "chief purser", "senior purser",
    }
    _is_ops_mgmt = wtype_key and not any(
        kw in wtype_key for kw in ("flight deck", "cabin crew", "engineering")
    )
    if _is_ops_mgmt:
        _before = len(skills_objs)
        skills_objs = [s for s in skills_objs
                       if (s.get("name") or "").strip().lower() not in _FLIGHT_DECK_ONLY_SKILLS]
        if len(skills_objs) < _before:
            print(f"  ℹ  Ops/Management profile: stripped {_before - len(skills_objs)} flight-deck-only skill(s)")

    # ── Engineering: strip aircraft types — only nationality + B1/B2 needed ─────
    if "engineering" in wtype_key:
        _eng_aircraft_re = re.compile(
            # Airbus A-series (A220, A300, A310, A318-A321, A330, A340, A350, A380)
            r"^(a[23][0-9]{2}[-/]?\w*"
            # Boeing B7xx series
            r"|b7[0-9]{2}[-/]?\w*"
            # ATR types: ATR42, ATR72, ATR72-600, ATR42/72
            r"|atr[-\d/]*"
            # CRJ / ERJ regional jets
            r"|crj[-\d]*|erj[-\d]*"
            # Dash / Q-series turboprops
            r"|dash[-\s\d]*|q[2-9][0-9]{2}[-/]?\w*"
            # Other: EMB, BAe, Fokker, Saab, Avro
            r"|emb[-\d]*|bae[-\d]*|fokker[-\d]*|saab[-\d]*|avro[-\d]*"
            r")\b",
            re.IGNORECASE
        )
        # Also strip licence authorities (EASA, ICAO, FAA etc.) and positions —
        # engineers need nationality + B1/B2 only, nothing else.
        _eng_strip_names = {
            "easa", "icao", "faa", "jaa", "casa", "sacaa", "caas", "gcaa", "dgca",
            "tcca", "hcaa", "dgac", "caac", "jcab",
            "captain", "first officer", "second officer", "co-pilot", "copilot",
            "purser", "chief purser", "senior purser",
        }
        _before = len(skills_objs)
        skills_objs = [s for s in skills_objs
                       if not _eng_aircraft_re.match((s.get("name") or "").strip())
                       and (s.get("name") or "").strip().lower() not in _eng_strip_names]
        if len(skills_objs) < _before:
            print(f"  ℹ  Engineering profile: stripped {_before - len(skills_objs)} non-engineering skill(s)")

    # ── Safety net: never send an empty skills list if the candidate already
    # has skills in Tracker. Bad OCR / failed parse must not wipe real data.
    _original_qs = rec.get("quickSkills") or []
    if not skills_objs and _original_qs:
        print("  ℹ  CV parse produced no skills — preserving existing Tracker skills.")
        skills_objs = _original_qs

    # ── Flight deck: ensure position skill is always present ─────────────────────
    # If no position skill exists (CV OCR failed / too short to parse), extract it
    # from the existing Tracker job title as a reliable fallback.
    if "flight deck" in wtype_key:
        _cur_pos_in_skills = any(
            (s.get("name") or "").lower() in _position_set for s in skills_objs
        )
        if not _cur_pos_in_skills:
            # Try to extract position from existing job title
            _title_src = rec_job or job_title or ""
            for _pos in POSITIONS:
                if re.search(r"\b" + re.escape(_pos) + r"\b", _title_src, re.IGNORECASE):
                    _pos_resolved = resolve_skills([_pos], skills_lookup)
                    if _pos_resolved:
                        skills_objs = skills_objs + _pos_resolved
                        print(f"  ℹ  Added position '{_pos}' from job title (OCR couldn't parse)")
                    break

    # ── Deduplicate multiple country/nationality skills ────────────────────────
    # Tracker stores two types of country skills by skill ID:
    #   nationality_ids      (area 43) = the candidate's citizenship/nationality
    #   licence_country_ids  (area 39) = country where licence was issued
    # Cabin crew: 1 nationality, 0 licence countries
    # Pilots/engineers: 1 nationality + 1 licence country
    if country_skills_set and nationality_ids:
        non_country       = [s for s in skills_objs
                             if (s.get("name") or "").strip().lower() not in country_skills_set]
        nat_skills        = [s for s in skills_objs if s.get("id") in nationality_ids]
        lic_skills        = [s for s in skills_objs if s.get("id") in licence_country_ids]
        # Skills whose IDs don't appear in either set — treat as nationality by name
        uncat_country     = [s for s in skills_objs
                             if (s.get("name") or "").strip().lower() in country_skills_set
                             and s.get("id") not in nationality_ids
                             and s.get("id") not in licence_country_ids]
        # Names of ALL country-like skills present in the EXISTING Tracker record —
        # used to guard against overriding real data with CV fallback guesses.
        _all_country_names_lower = (KNOWN_COUNTRY_NAMES_EXTRA
                                    | {v.lower() for v in COUNTRY_ALIASES.values()})
        _existing_tracker_countries = {
            (s.get("name") or "").strip().lower()
            for s in skills_objs
            if (s.get("name") or "").strip().lower() in _all_country_names_lower
        }
        # Use CV nationality to pick the correct nationality skill
        cv_nat_lower = (nationality or "").strip().lower()
        # Use the CORRECTED work_type_objs (after work-type conflict resolution above),
        # not the raw Tracker value — e.g. Harry James had Flight Deck in Tracker but
        # was registered via a Cabin Crew email, so work_type_objs is now [Cabin Crew].
        work_type_ids = {w.get("id") for w in work_type_objs}
        CABIN_CREW_ID   = 469
        FLIGHT_DECK_ID2 = 472
        ENGINEERING_ID2 = 471
        is_cabin_crew   = CABIN_CREW_ID   in work_type_ids
        is_flight_deck  = FLIGHT_DECK_ID2 in work_type_ids
        is_engineering2 = ENGINEERING_ID2 in work_type_ids
        # Dispatchers get Flight Deck as a secondary work type for candidate matching,
        # but they are NOT pilots — suppress pilot-specific skill rules (ICAO/EASA
        # authority enforcement, local authority stripping).
        is_dispatcher = "dispatcher" in wtype_key or any(
            x in (job_title or "").lower() for x in ["dispatcher", "flight dispatch", "aircraft dispatch"]
        )
        if is_dispatcher:
            is_flight_deck = False
        # Only Flight Deck gets two country skills (nationality + FCL country).
        # Everyone else — including Engineering — gets 1 nationality only.
        is_single_nat   = not is_flight_deck

        # Flight deck: if no aircraft type in skills yet, scan the existing Tracker
        # job title (rec_job) — catches "Captain, Boeing 737 NG" style titles.
        if is_flight_deck:
            _has_ac = any((s.get("name") or "").strip() in AIRCRAFT_TYPES
                          for s in skills_objs)
            if not _has_ac and rec_job:
                _jt_norm = _norm_ac_text(rec_job)
                for _ac in AIRCRAFT_TYPES:
                    if re.search(r"\b" + re.escape(_ac) + r"\b", _jt_norm, re.IGNORECASE):
                        _ac_obj = next(({"id": v["id"], "name": v["name"]}
                                        for k, v in skills_lookup.items()
                                        if k == _ac.lower()), None)
                        if _ac_obj:
                            non_country.append(_ac_obj)
                            print(f"  ℹ  Added aircraft '{_ac}' from existing job title")
                            break

        # Flight deck: strip country-specific authority skills (DGCA, GCAA, SACAA etc).
        # Emily's rule: only EASA / FAA / ICAO are allowed as authority skills.
        # NOTE: strip also applied to non_country here so they don't sneak back in
        # when skills_objs is rebuilt from component parts below.
        if is_flight_deck:
            _local_auths = {"dgca","gcaa","caa","sacaa","jcab","anac","casa","tcca",
                            "caac","caam","caas","caav","kcaa","ncaa","bcaa","hcaa","bcaa"}
            skills_objs = [s for s in skills_objs
                           if (s.get("name") or "").strip().lower() not in _local_auths]
            non_country = [s for s in non_country
                           if (s.get("name") or "").strip().lower() not in _local_auths]

        # Flight deck: ensure the correct authority (EASA / FAA / ICAO) is present,
        # and strip any conflicting authorities derived from CV text.
        # The correct authority is DERIVED from the FCL country:
        #   EU/EEA country  → EASA
        #   United States   → FAA
        #   Everywhere else → ICAO
        if is_flight_deck:
            _TOP_AUTHS = {"icao", "easa", "faa"}
            _EASA_COUNTRIES = {
                "austria","belgium","bulgaria","croatia","cyprus","czech republic",
                "denmark","estonia","finland","france","germany","greece","hungary",
                "iceland","ireland","italy","latvia","liechtenstein","lithuania",
                "luxembourg","malta","netherlands","norway","poland","portugal",
                "romania","slovakia","slovenia","spain","sweden",
            }
            # Use lic_skills country first, fall back to nat_skills
            _fcl_cn = ((lic_skills[0].get("name") or "") if lic_skills
                       else (nat_skills[0].get("name") or "") if nat_skills
                       else "").strip().lower()
            if _fcl_cn in _EASA_COUNTRIES:
                _auth_key = "easa"
            elif _fcl_cn in {"united states", "usa", "us"}:
                _auth_key = "faa"
            else:
                _auth_key = "icao"
            # Strip any wrong top-authority skills from both buckets
            _wrong_auths = _TOP_AUTHS - {_auth_key}
            _before = len(skills_objs) + len(non_country)
            skills_objs = [s for s in skills_objs
                           if (s.get("name") or "").strip().lower() not in _wrong_auths]
            non_country = [s for s in non_country
                           if (s.get("name") or "").strip().lower() not in _wrong_auths]
            if len(skills_objs) + len(non_country) < _before:
                print(f"  ℹ  Stripped conflicting authority skills (keeping {_auth_key.upper()} for FCL country '{_fcl_cn}')")
            # Now ensure the correct authority is present
            _has_correct_auth = any((s.get("name") or "").strip().lower() == _auth_key
                                    for s in skills_objs + non_country)
            if not _has_correct_auth:
                _auth_obj = next(({"id": v["id"], "name": v["name"]}
                                  for k, v in skills_lookup.items() if k == _auth_key), None)
                if _auth_obj:
                    non_country.insert(0, _auth_obj)
                    print(f"  ℹ  Added {_auth_key.upper()} authority (derived from FCL country '{_fcl_cn}')")

        # Cabin Crew: enforce exactly one service level — Main Crew, VIP, or Business.
        # Strip any wrong/stale levels first, then ensure the correct one is present.
        if is_cabin_crew:
            _VALID_CABIN_LEVELS = {"main crew", "senior / cabin manager", "vip", "business"}
            _INVALID_CABIN_LEVELS = {
                "senior cabin crew", "senior", "cabin manager", "purser", "chief purser",
                "senior purser", "business class", "first class",
            }
            # Remove any invalid cabin level skills from all buckets
            for _bucket in (non_country, nat_skills, lic_skills, uncat_country):
                _bucket[:] = [s for s in _bucket
                               if (s.get("name") or "").strip().lower() not in _INVALID_CABIN_LEVELS]
            skills_objs = [s for s in skills_objs
                           if (s.get("name") or "").strip().lower() not in _INVALID_CABIN_LEVELS]
            # Ensure exactly one valid level — if >1 present, keep only the highest
            _level_priority = {"senior / cabin manager": 3, "vip": 3, "business": 2, "main crew": 1}
            _existing_levels = [s for s in skills_objs
                                 if (s.get("name") or "").strip().lower() in _VALID_CABIN_LEVELS]
            if len(_existing_levels) > 1:
                _keep = max(_existing_levels,
                            key=lambda s: _level_priority.get((s.get("name") or "").strip().lower(), 0))
                skills_objs = [s for s in skills_objs
                               if (s.get("name") or "").strip().lower() not in _VALID_CABIN_LEVELS]
                skills_objs.append(_keep)
                non_country = [s for s in non_country
                               if (s.get("name") or "").strip().lower() not in _VALID_CABIN_LEVELS]
                print(f"  ℹ  Multiple cabin levels — keeping highest: {_keep.get('name')}")
            # If no valid level present at all, add Main Crew as default
            _has_level = any((s.get("name") or "").strip().lower() in _VALID_CABIN_LEVELS
                              for s in skills_objs)
            if not _has_level:
                _mc_obj = next(({"id": v["id"], "name": v["name"]}
                                for k, v in skills_lookup.items() if k == "main crew"), None)
                if _mc_obj:
                    non_country.insert(0, _mc_obj)
                    print(f"  ℹ  Added Main Crew skill (cabin crew candidate)")

        # Strip regional terms (Europe, Middle East, Asia etc.) from both buckets —
        # they must never be chosen as a candidate's nationality.
        nat_skills    = [s for s in nat_skills
                         if (s.get("name") or "").strip().lower() not in REGIONAL_TERMS]
        uncat_country = [s for s in uncat_country
                         if (s.get("name") or "").strip().lower() not in REGIONAL_TERMS]

        # Reduce nationality skills to 1
        if len(nat_skills) + len(uncat_country) > 1:
            all_nat = nat_skills + uncat_country
            # If there are >4 country skills, this is almost certainly stale
            # work-destination data (e.g. supply chain candidate with 9 countries).
            # Don't try to pick one — clear all and let parse_cv nationality take over.
            if len(all_nat) > 4 and not cv_nat_lower:
                print(f"  ℹ  {len(all_nat)} country skills found — treating as stale destination "
                      f"data, clearing and using CV-detected nationality")
                nat_skills = []
                uncat_country = []
            else:
                # 1. Prefer exact match with CV-detected nationality
                chosen_nat = next((s for s in all_nat
                                   if (s.get("name") or "").strip().lower() == cv_nat_lower), None)
                # 2. Prefer one that appears in CV text in a non-route context
                if not chosen_nat and cv_text:
                    _cv_lc = cv_text.lower()
                    _ROUTE_CTX_NAT = re.compile(
                        r"\b(?:to|from|via|fly(?:ing)?|routes?|destination|based\s+in|"
                        r"operated?(?:\s+to)?|flights?\s+to|serv(?:es?|ing)|travel(?:l?ing)?)\s*$",
                        re.IGNORECASE
                    )
                    for _s in all_nat:
                        _sname = (_s.get("name") or "").strip().lower()
                        for _m in re.finditer(r"\b" + re.escape(_sname) + r"\b", _cv_lc):
                            _pre = _cv_lc[max(0, _m.start() - 60):_m.start()]
                            if _ROUTE_CTX_NAT.search(_pre.strip()):
                                continue  # skip route/destination mention
                            chosen_nat = _s
                            break
                        if chosen_nat:
                            break
                # 3. If none verified in non-route context, preserve the first existing
                #    Tracker nationality rather than wiping it (short CVs may not mention it)
                if chosen_nat:
                    print(f"  ℹ  Multiple nationality skills → keeping: {chosen_nat['name']}")
                    nat_skills = [chosen_nat]
                    uncat_country = []
                else:
                    _preserved = next(
                        (s for s in all_nat if (s.get("name") or "").strip().lower()
                         in _existing_tracker_countries),
                        all_nat[0] if all_nat else None
                    )
                    if _preserved:
                        print(f"  ℹ  Multiple nationality skills — none verified in CV "
                              f"→ preserving existing: {_preserved['name']}")
                        nat_skills = [_preserved]
                        uncat_country = []
                    else:
                        print(f"  ℹ  Multiple nationality skills — none verified in CV → clearing stale data")
                        nat_skills = []
                        uncat_country = []

        # Reduce licence-country skills to 0 (cabin crew / ops / mgmt / airport) or 1 (flight deck / engineering)
        if is_single_nat:
            if lic_skills:
                # If nat_skills and uncat_country are BOTH empty, lic_skills is the ONLY
                # nationality data we have. Stripping it would produce []. Instead,
                # try to resolve the country to an area-43 nationality skill; if that
                # fails, just keep it as-is rather than wiping the candidate's nationality.
                if not nat_skills and not uncat_country:
                    _rescued_from_lic = []
                    for _ls in lic_skills[:1]:  # promote at most 1
                        _lname = (_ls.get("name") or "").strip()
                        # Prefer the area-43 (nationality) version of this country name
                        _nat_obj = next(
                            ({"id": v["id"], "name": v["name"]}
                             for k, v in skills_lookup.items()
                             if k == _lname.lower() and v.get("id") in nationality_ids),
                            None
                        )
                        if _nat_obj:
                            _rescued_from_lic.append(_nat_obj)
                            print(f"  ℹ  Promoted licence country '{_lname}' to nationality skill "
                                  f"(was only country data for {', '.join(w['name'] for w in work_type_objs)} profile)")
                        else:
                            _rescued_from_lic.append(_ls)
                            print(f"  ℹ  Keeping '{_lname}' as nationality (only country data; "
                                  f"no area-43 nationality form found)")
                    nat_skills = _rescued_from_lic
                    lic_skills = []
                else:
                    print(f"  ℹ  {', '.join(w['name'] for w in work_type_objs)} — removing licence country skills: {[s['name'] for s in lic_skills]}")
                    lic_skills = []
            else:
                lic_skills = []
        elif len(lic_skills) > 1:
            chosen_lic = lic_skills[0]
            print(f"  ℹ  Multiple licence countries → keeping: {chosen_lic['name']}")
            lic_skills = [chosen_lic]

        # Dedup within non_country by name.
        # Allow at most 2 of the same name only when it matches cv_nat_lower —
        # that means nationality == licence country (e.g. Korea/Korea), which is valid.
        # Any other same-name duplicate is a data error: keep just the first.
        _seen_nc: dict = {}
        _nc_deduped = []
        for s in non_country:
            _n = (s.get("name") or "").strip().lower()
            _count = _seen_nc.get(_n, 0)
            if _count == 0:
                _nc_deduped.append(s)
                _seen_nc[_n] = 1
            elif _count == 1 and _n == cv_nat_lower:
                # nationality == licence country landing in non_country — keep both
                _nc_deduped.append(s)
                _seen_nc[_n] = 2
            # else: skip duplicate
        non_country = _nc_deduped

        # If cv_nat_lower explicitly matches a skill sitting in non_country
        # (Tracker stored it with wrong area ID), treat those non_country entries
        # as the authoritative nationality/licence-country and drop conflicting
        # entries from nat_skills / lic_skills.
        # Only promote CV nationality from the non-country bucket when:
        #   a) nationality was explicitly found via a keyword (not a fallback guess), OR
        #      there is truly no existing country data in the Tracker record at all.
        # This prevents e.g. "United States" (picked up from a header address) overriding
        # a real existing nationality like Egypt or Brazil that Tracker already holds.
        # Only trust existing Tracker countries that actually appear in the CV text.
        # If a country is in Tracker but nowhere in the CV, it's likely stale/wrong data —
        # don't let it block CV-detected nationality from being applied.
        _cv_full_lower = (cv_text or "").lower()
        _existing_tracker_countries_in_cv = {
            c for c in _existing_tracker_countries
            if re.search(r"\b" + re.escape(c) + r"\b", _cv_full_lower)
        }
        _has_existing_countries = bool(nat_skills or uncat_country or _existing_tracker_countries_in_cv)
        if cv_nat_lower and (nat_explicit or not _has_existing_countries) and not any(
            (s.get("name") or "").strip().lower() == cv_nat_lower
            for s in nat_skills + uncat_country + lic_skills
        ):
            matching_nc = [s for s in non_country
                           if (s.get("name") or "").strip().lower() == cv_nat_lower]
            if matching_nc:
                non_country = [s for s in non_country
                               if (s.get("name") or "").strip().lower() != cv_nat_lower]
                nat_skills = [matching_nc[0]]
                uncat_country = []
                if len(matching_nc) > 1 and not is_single_nat:
                    if not lic_skills:
                        lic_skills = [matching_nc[1]]
                else:
                    lic_skills = []
                print(f"  ℹ  CV nationality '{matching_nc[0]['name']}' found in non-country skills "
                      f"— promoted to nationality/licence country bucket.")

        # Country cleanup: remove any misclassified country names from non_country.
        # Flight deck: 1 nationality + 1 licence country, no extras allowed.
        # Cabin crew: 1 nationality, no extras at all.
        # Engineering/other: leave non_country alone — multi-country licensing is valid.
        _all_country_names = country_skills_set | KNOWN_COUNTRY_NAMES_EXTRA | {
            v.lower() for v in COUNTRY_ALIASES.values()
        }
        if is_flight_deck:
            extra_in_nc = [s for s in non_country
                           if (s.get("name") or "").strip().lower() in _all_country_names]
            if extra_in_nc:
                # If no licence country yet, try to promote one from non_country.
                # Only promote countries that are actually linked to a licence authority
                # (e.g. United Kingdom → CAA, Australia → CASA, India → DGCA).
                # Work-location countries (Saudi Arabia, Yemen, UAE etc.) must NOT become
                # the FCL country — they get stripped and FCL defaults to nationality instead.
                if not lic_skills:
                    nat_name_lc = (nat_skills[0].get("name") or "").strip().lower() if nat_skills else ""
                    _known_fcl_countries = (
                        {v.lower() for v in LICENCE_AUTH_TO_COUNTRY.values()} |
                        {"austria","belgium","bulgaria","croatia","cyprus","czech republic",
                         "denmark","estonia","finland","france","germany","greece","hungary",
                         "iceland","ireland","italy","latvia","liechtenstein","lithuania",
                         "luxembourg","malta","netherlands","norway","poland","portugal",
                         "romania","slovakia","slovenia","spain","sweden"}
                    )
                    chosen_lic = next(
                        (s for s in extra_in_nc
                         if (s.get("name") or "").strip().lower() in _known_fcl_countries
                         and (s.get("name") or "").strip().lower() != nat_name_lc),
                        None  # work-location countries are not valid FCL countries
                    )
                    if chosen_lic:
                        # Re-resolve via area-39 lookup so the skill gets the FCL country
                        # ID (not the nationality area-43 ID it may have from Tracker).
                        _chosen_name = (chosen_lic.get("name") or "").strip()
                        _fcl_resolved = licence_country_lookup.get(_chosen_name.lower())
                        lic_skills = [_fcl_resolved if _fcl_resolved else chosen_lic]
                        extra_in_nc = [s for s in extra_in_nc if s is not chosen_lic]
                        non_country = [s for s in non_country if s is not chosen_lic]
                        print(f"  ℹ  Flight deck: promoted '{_chosen_name}' to licence country "
                              f"(was in wrong skill bucket)")
                # If nationality still empty after promoting lic country, rescue one more
                # country from extra_in_nc rather than throwing it away.
                if not nat_skills and extra_in_nc:
                    lic_name_lc = (lic_skills[0].get("name") or "").strip().lower() if lic_skills else ""
                    chosen_nat_rescue = next(
                        (s for s in extra_in_nc
                         if (s.get("name") or "").strip().lower() != lic_name_lc),
                        extra_in_nc[0]
                    )
                    nat_skills = [chosen_nat_rescue]
                    extra_in_nc = [s for s in extra_in_nc if s is not chosen_nat_rescue]
                    non_country = [s for s in non_country if s is not chosen_nat_rescue]
                    print(f"  ℹ  Flight deck: rescued '{chosen_nat_rescue['name']}' as nationality "
                          f"(was in non-country bucket)")
                if extra_in_nc:
                    print(f"  ℹ  Flight deck: removing extra country skills from non-country bucket: "
                          f"{[s['name'] for s in extra_in_nc]}")
                    _nc_remove = {(s.get("name") or "").strip().lower() for s in extra_in_nc}
                    non_country = [s for s in non_country
                                   if (s.get("name") or "").strip().lower() not in _nc_remove]

            # Flight deck: infer FCL country from licence authority found in CV.
            # Only runs when lic_skills is not already set — prevents FAA/CASA found
            # incidentally in CV (e.g. type rating context) from overriding an FCL
            # country already established from the Tracker record or nationality.
            # Ambiguous authorities (EASA, ICAO) are not in LICENCE_AUTH_TO_COUNTRY.
            if cv_text and not lic_skills:
                import re as _re2
                # Authority must appear in a licence context — not just anywhere in the CV.
                # Prevents false matches from e.g. "FAA equivalent" or "FAA regulations".
                _LICENCE_CONTEXT_RE = _re2.compile(
                    r"\b(licence|license|atpl|cpl|ppl|type rating|rating|approved|issued|"
                    r"certified|certif|authoris|authoriz|validation|privileges)\b",
                    _re2.IGNORECASE
                )
                for _auth, _fcl_country in LICENCE_AUTH_TO_COUNTRY.items():
                    for _m in _re2.finditer(r"\b" + _re2.escape(_auth) + r"\b", cv_text, _re2.IGNORECASE):
                        # Check within 80 chars either side for a licence keyword
                        _window = cv_text[max(0, _m.start()-80): _m.end()+80]
                        if _LICENCE_CONTEXT_RE.search(_window):
                            _fcl_lc = _fcl_country.lower()
                            # Use area-39 (FCL country) lookup so the skill gets the
                            # correct FCL country ID, not the nationality (area-43) ID.
                            _fcl_obj = licence_country_lookup.get(_fcl_lc)
                            if not _fcl_obj:
                                _fcl_obj = skills_lookup.get(_fcl_lc)
                            if not _fcl_obj:
                                _fcl_obj = {"id": 0, "name": _fcl_country}
                            lic_skills = [_fcl_obj]
                            print(f"  ℹ  FCL country inferred from {_auth} authority: '{_fcl_country}'")
                            break
                    if lic_skills:
                        break
            # If still no FCL country, default to same country as nationality.
            # Use area-39 lookup specifically — the main skills_lookup stores the
            # area-43 ID for each country (nationality takes priority), so we must
            # use licence_country_lookup to get the distinct area-39 skill ID.
            if not lic_skills and nat_skills:
                _nat_name = (nat_skills[0].get("name") or "").strip()
                _fcl_obj  = licence_country_lookup.get(_nat_name.lower())
                if not _fcl_obj:
                    # fallback: any entry in skills_lookup with a different ID
                    _fcl_obj = next(({"id": v["id"], "name": v["name"]}
                                     for k, v in skills_lookup.items()
                                     if k == _nat_name.lower()
                                     and v["id"] != nat_skills[0].get("id")), None)
                if _fcl_obj:
                    lic_skills = [_fcl_obj]
                    print(f"  ℹ  FCL country defaulted to nationality '{_nat_name}' (no other FCL info found)")
        elif is_single_nat:
            extra_in_nc = [s for s in non_country
                           if (s.get("name") or "").strip().lower() in _all_country_names]
            if extra_in_nc:
                # If there are no nationality skills anywhere else, rescue one country from
                # non_country before stripping — prevents the profile ending up with [] skills.
                if not nat_skills and not uncat_country:
                    rescue = next(
                        (s for s in extra_in_nc
                         if cv_nat_lower and (s.get("name") or "").strip().lower() == cv_nat_lower),
                        extra_in_nc[0]
                    )
                    nat_skills = [rescue]
                    extra_in_nc = [s for s in extra_in_nc if s is not rescue]
                    non_country = [s for s in non_country if s is not rescue]
                    print(f"  ℹ  Rescued '{rescue['name']}' as nationality (was in wrong skill area)")
                wt_label = ", ".join(w["name"] for w in work_type_objs)
                if extra_in_nc:
                    print(f"  ℹ  {wt_label}: removing extra country skills from non-country bucket: "
                          f"{[s['name'] for s in extra_in_nc]}")
                    non_country = [s for s in non_country
                                   if (s.get("name") or "").strip().lower() not in _all_country_names]

        skills_objs = non_country + nat_skills + uncat_country + lic_skills

        # Safety net: nationality must survive for all profiles.
        # If the assembled list has no nationality/country skill at all, rescue one
        # from the ORIGINAL Tracker record.  The check uses nationality_ids (area-43 IDs)
        # so multi-word demonyms like 'Saudi Arabian' are caught even if they're not in
        # NATIONALITY_ADJECTIVES (which only has single-word forms like 'saudi').
        _has_nat_adj = any(
            s.get("id") in nationality_ids or  # area-43 skill by Tracker ID
            (s.get("name") or "").strip().lower() in NATIONALITY_ADJECTIVES
            for s in skills_objs
        )
        _has_country = any(
            (s.get("name") or "").strip().lower() in _all_country_names
            for s in skills_objs
        )
        if not _has_nat_adj and not _has_country:
            _rescue_nat = next(
                (s for s in _original_qs
                 if s.get("id") in nationality_ids  # area-43 skill
                 or (s.get("name") or "").strip().lower() in NATIONALITY_ADJECTIVES
                 or (s.get("name") or "").strip().lower() in country_skills_set),
                None
            )
            if _rescue_nat:
                skills_objs.append(_rescue_nat)
                print(f"  ℹ  Rescued nationality '{_rescue_nat['name']}' from original Tracker record")

        # Final dedup — by skill ID only. Same country can appear twice legitimately:
        # once as nationality (area 43) and once as licence country (area 39).
        # Free-text skills (id=0) dedup by name instead.
        # Pre-dedup: strip out skill entries with clearly nonsensical names
        # (OCR artefacts, English words that are not skills, None-containing values)
        _INVALID_SKILL_NAME_RE = re.compile(
            r"^\s*(valid|invalid|none|n/a|not\s+applicable|current|expired|"
            r"dual|multi|single|british\s+indian\s+ocean\s+territory)\s*$",
            re.IGNORECASE
        )
        _NONE_WORD_RE = re.compile(r"\bnone\b", re.IGNORECASE)
        _skills_objs_cleaned = []
        for _s in skills_objs:
            _raw_name = (_s.get("name") or "").strip()
            if not _raw_name:
                continue
            if _INVALID_SKILL_NAME_RE.match(_raw_name) or _NONE_WORD_RE.search(_raw_name):
                print(f"  ℹ  Discarding invalid skill name: '{_raw_name}'")
                continue
            _skills_objs_cleaned.append(_s)
        skills_objs = _skills_objs_cleaned

        _final_seen_ids: set = set()
        _final_seen_names: set = set()
        _final_deduped = []
        for _s in skills_objs:
            _sn  = (_s.get("name") or "").strip().lower()
            _sid = _s.get("id")
            if not _sn:
                continue
            # Always dedup by name — same country can appear with different area IDs
            # (e.g. Greece as nationality area-43 AND as FCL country area-39).
            # Keep the first occurrence (which will be the correctly resolved one).
            if _sn in _final_seen_names:
                continue
            if _sid and _sid != 0:
                if _sid not in _final_seen_ids:
                    _final_deduped.append(_s)
                    _final_seen_ids.add(_sid)
                    _final_seen_names.add(_sn)
            else:  # free-text (id=0) — dedup by name only
                _final_deduped.append(_s)
                _final_seen_names.add(_sn)
        skills_objs = _final_deduped

    # Auto-inject 'Flight instructor' skill for Flight Deck candidates whose
    # job title (current or existing) indicates they are a flight instructor.
    # Rule: "Flight instructors should be treated the same as pilots;
    #        flight instructor should also be added to their skill."
    if is_flight_deck:
        _has_fi_skill = any(
            (s.get("name") or "").strip().lower() in ("flight instructor", "flight instructor ")
            for s in skills_objs
        )
        if not _has_fi_skill:
            _fi_title_check = ((job_title or "") + " " + (rec_job or "")).lower()
            _fi_keywords = ("flight instructor", "flying instructor", " cfi", " tki", " dfi")
            if any(kw in _fi_title_check for kw in _fi_keywords):
                _fi_obj = skills_lookup.get("flight instructor")
                if _fi_obj:
                    skills_objs.append({"id": _fi_obj["id"], "name": _fi_obj["name"]})
                    print("  ℹ  Auto-added 'Flight instructor' skill (job title indicates flight instructor)")

    # 5. Review & confirm
    wt_display = ", ".join(w["name"] for w in work_type_objs)
    print("\n[5/5] Proposed changes:")
    print(f"  Job title:  {job_title}{'  (kept existing)' if job_title == rec_job else ''}")
    print(f"  Employer:   {employer}{'  (kept existing)' if employer == rec_employer else ''}")
    print(f"  Work type:  {wt_display}")
    print(f"  Skills:     {[s['name'] for s in skills_objs]}")

    confirm = "y"  # auto-apply in auto-run mode

    if confirm == "e":
        # Show parsed value; fall back to current Tracker value if Enter pressed
        jt_default  = job_title  or rec_job
        emp_default = employer   or rec_employer
        job_title = input(f"  Job title [{jt_default}]: ").strip() or jt_default
        employer  = input(f"  Employer  [{emp_default}]: ").strip() or emp_default
        wt = input(f"  Work type(s) — comma-separated for multiple [{wt_display}]: ").strip()
        if wt:
            work_type_objs = resolve_work_types(wt)
        sk = input(f"  Skills (comma-separated) [{','.join(s['name'] for s in skills_objs)}]: ").strip()
        if sk:
            skills_objs = resolve_skills([x.strip() for x in sk.split(",")], skills_lookup)
        confirm = "y"

    if confirm != "y":
        print("  Skipped.")
        return False

    if DRY_RUN:
        print("  ⚠  DRY RUN — changes NOT written to Tracker.")
        print("     (Set DRY_RUN = False in the script once parsing looks correct)")
        return True

    # ── Final fallback: write "Unknown" if job title or employer still empty ─────
    # An empty field is worse than "Unknown" — it leaves recruiters with nothing.
    if not job_title:
        job_title = "Unknown"
        print("  ℹ  Job title could not be determined — setting to 'Unknown'")
    if not employer:
        employer = "Unknown"
        print("  ℹ  Employer could not be determined — setting to 'Unknown'")

    status, resp = update_resource(jwt, resource_id, job_title, employer, work_type_objs, skills_objs,
                                   first_name=rec_first, surname=rec_sur)
    if status in (200, 204):
        print(f"  ✓ Updated! (HTTP {status})")
        candidate_full_name = f"{rec_first} {rec_sur}".strip() or name
        log_daily_update(resource_id, candidate_full_name, job_title, employer, work_type_objs, skills_objs)
        if email_cand:
            if move_to_done(email_cand):
                print(f"  ✓ Email moved to '{EMAIL_DONE_FOLDER}'")
        return True
    else:
        print(f"  ✗ Update failed (HTTP {status}): {resp[:300]}")
        return False

# ── Email reading & moving ─────────────────────────────────────────────────────

def _get_support_inbox():
    """Return the Inbox folder of the support@aeroprofessional.com mailbox."""
    import win32com.client
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns      = outlook.GetNamespace("MAPI")

    # Method 1: scan Stores (support mailbox added as shared account)
    for store in ns.Stores:
        dn = (store.DisplayName or "").lower()
        if "support" in dn or SUPPORT_MAILBOX.lower() in dn:
            try:
                root = store.GetRootFolder()
                for inbox_name in ["Inbox", "Postvak IN"]:
                    try:
                        return root.Folders[inbox_name]
                    except Exception:
                        pass
                return root.Folders.Item(1)
            except Exception:
                pass

    # Fallback: default inbox
    return ns.GetDefaultFolder(6)

def _find_folder_recursive(parent, target_name, depth=0):
    """Search all folders in the hierarchy — handles unsynced/hidden folders."""
    if depth > 6:
        return None
    try:
        count = parent.Folders.Count
        for i in range(1, count + 1):
            try:
                f = parent.Folders.Item(i)
                if f.Name.strip().upper() == target_name.strip().upper():
                    return f
                result = _find_folder_recursive(f, target_name, depth + 1)
                if result:
                    return result
            except Exception:
                continue
    except Exception:
        pass
    return None

def _get_source_folder():
    if EMAIL_SOURCE != "subfolder":
        return _get_support_inbox()

    import win32com.client
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns      = outlook.GetNamespace("MAPI")

    # Method 1: check ALL open Outlook windows (not just the active one)
    try:
        for i in range(1, outlook.Explorers.Count + 1):
            try:
                exp = outlook.Explorers.Item(i)
                cf  = exp.CurrentFolder
                if cf and cf.Name.strip().upper() == EMAIL_SUBFOLDER.strip().upper():
                    print(f"  ✓ Found '{EMAIL_SUBFOLDER}' via open Outlook window")
                    return cf
            except Exception:
                continue
    except Exception:
        pass

    # Method 2: direct string-key access (different MAPI path — bypasses count cache)
    for store in ns.Stores:
        dn = (store.DisplayName or "").lower()
        if "support" in dn or SUPPORT_MAILBOX.lower() in dn:
            try:
                root = store.GetRootFolder()
                for inbox_name in ["Inbox", "INBOX"]:
                    try:
                        inbox  = root.Folders[inbox_name]
                        folder = inbox.Folders[EMAIL_SUBFOLDER]
                        print(f"  ✓ Found '{EMAIL_SUBFOLDER}' via direct string-key access")
                        return folder
                    except Exception:
                        pass
            except Exception:
                pass

    # Method 3: recursive index search through entire mailbox hierarchy
    for store in ns.Stores:
        dn = (store.DisplayName or "").lower()
        if "support" in dn or SUPPORT_MAILBOX.lower() in dn:
            try:
                root   = store.GetRootFolder()
                result = _find_folder_recursive(root, EMAIL_SUBFOLDER)
                if result:
                    print(f"  ✓ Found '{EMAIL_SUBFOLDER}' via recursive search")
                    return result
            except Exception:
                pass

    # Method 4: ns.Folders top-level iteration (different MAPI path to stores)
    try:
        for i in range(1, ns.Folders.Count + 1):
            try:
                top = ns.Folders.Item(i)
                dn  = (getattr(top, "Name", "") or "").lower()
                if "support" in dn or SUPPORT_MAILBOX.lower() in dn:
                    result = _find_folder_recursive(top, EMAIL_SUBFOLDER)
                    if result:
                        print(f"  ✓ Found '{EMAIL_SUBFOLDER}' via ns.Folders traversal")
                        return result
                    # Also try direct string-key from this root
                    for inbox_name in ["Inbox", "INBOX"]:
                        try:
                            inbox  = top.Folders[inbox_name]
                            folder = inbox.Folders[EMAIL_SUBFOLDER]
                            print(f"  ✓ Found '{EMAIL_SUBFOLDER}' via ns.Folders string key")
                            return folder
                        except Exception:
                            pass
            except Exception:
                continue
    except Exception:
        pass

    # Method 5: search from default inbox
    try:
        inbox  = ns.GetDefaultFolder(6)
        result = _find_folder_recursive(inbox, EMAIL_SUBFOLDER)
        if result:
            print(f"  ✓ Found '{EMAIL_SUBFOLDER}' in default inbox")
            return result
    except Exception:
        pass

    raise Exception(f"Subfolder '{EMAIL_SUBFOLDER}' not found via COM")

def _get_done_folder(folder_path=None):
    if folder_path is None:
        folder_path = EMAIL_DONE_FOLDER
    inbox  = _get_support_inbox()
    folder = inbox
    for part in folder_path.split("/"):
        part = part.strip()
        try:
            folder = folder.Folders[part]
        except Exception:
            # Folder doesn't exist yet — create it
            try:
                folder = folder.Folders.Add(part)
                print(f"  ℹ  Created mail folder: '{part}'")
            except Exception as _fe:
                raise Exception(f"Cannot create/find folder '{part}': {_fe}")
    return folder

def _strip_html(html):
    """Strip HTML tags and normalise whitespace to plain text."""
    text = re.sub(r'<br\s*/?>', '\n', html, flags=re.IGNORECASE)
    text = re.sub(r'<p[^>]*>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'&nbsp;', ' ', text)
    text = re.sub(r'&amp;', '&', text)
    text = re.sub(r'&lt;', '<', text)
    text = re.sub(r'&gt;', '>', text)
    text = re.sub(r'&#\d+;', '', text)
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n\s*\n+', '\n', text)
    return text.strip()

def parse_email_for_name(subject, body):
    """Extract candidate name from a registration or Jobs+ email body."""
    # Strip HTML if present
    clean = _strip_html(body) if '<' in body else body

    # Type 1: "Candidate Registration through Website"
    #   "Hi Admin,\n\nJohn Smith\n\nhas registered themselves"
    m = re.search(r"Hi Admin[,.]?\s+(.+?)\s+has registered themselves", clean, re.IGNORECASE | re.DOTALL)
    if m:
        name = m.group(1).strip()
        # Reject if name captured is unreasonably long (HTML leak)
        if len(name) < 80:
            return name

    # Type 2: "New Candidate Application - Jobs+"
    #   "John Smith has applied for Job" or "John Smith\nhas applied for Job"
    m = re.search(r"([A-Za-z][^\n]{1,60}?)\s*\n?\s*has applied for Job", clean, re.IGNORECASE)
    if m:
        name = m.group(1).strip()
        if len(name) < 80:
            return name

    return None

OWA_SESSION_FILE = os.path.join(os.path.expanduser("~"), "tracker_owa_session.json")
CACHE_FILE_GLOBAL = "tracker_cache.json"
OWA_TOKEN_MAX_MINS = 240   # OWA tokens last ~4 hours


def _save_owa_token(token):
    """Save the OWA bearer token to the cache file so the browser can be skipped next run."""
    try:
        cache = {}
        if os.path.exists(CACHE_FILE_GLOBAL):
            with open(CACHE_FILE_GLOBAL) as f:
                cache = json.load(f)
        cache["owa_token"]    = token
        cache["owa_token_ts"] = time.time()
        with open(CACHE_FILE_GLOBAL, "w") as f:
            json.dump(cache, f)
    except Exception:
        pass


def _load_cached_owa_token():
    """Return a cached OWA token if it's still fresh, otherwise None."""
    try:
        if not os.path.exists(CACHE_FILE_GLOBAL):
            return None
        with open(CACHE_FILE_GLOBAL) as f:
            cache = json.load(f)
        token = cache.get("owa_token")
        ts    = cache.get("owa_token_ts", 0)
        age_mins = (time.time() - ts) / 60
        if token and age_mins < OWA_TOKEN_MAX_MINS:
            return token
    except Exception:
        pass
    return None


def _fetch_emails_with_token(token):
    """
    Fetch all emails from the target folder using a known OWA bearer token.
    Returns list of raw message dicts, or raises on failure.
    """
    hdrs = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    mb   = SUPPORT_MAILBOX
    base = "https://outlook.cloud.microsoft/api/beta"

    # Find the subfolder ID
    fid = None
    for folders_url in [
        f"{base}/users/{mb}/MailFolders?$top=200&$select=Id,DisplayName",
        f"{base}/users/{mb}/MailFolders/Inbox/ChildFolders?$top=200&$select=Id,DisplayName",
    ]:
        r = requests.get(folders_url, headers=hdrs, timeout=30)
        if r.status_code == 200:
            for folder in r.json().get("value", []):
                if folder.get("DisplayName", "").lower() == EMAIL_SUBFOLDER.lower():
                    fid = folder["Id"]
                    break
        if fid:
            break

    if not fid:
        raise Exception(f"Folder '{EMAIL_SUBFOLDER}' not found")

    # Page through all messages
    all_messages = []
    url = (f"{base}/users/{mb}/MailFolders/{fid}/Messages"
           f"?$top=1000&$select=Id,Subject,Body,BodyPreview,ReceivedDateTime")
    while url:
        r = requests.get(url, headers=hdrs, timeout=60)
        if r.status_code != 200:
            raise Exception(f"Messages API returned {r.status_code}")
        data = r.json()
        all_messages.extend(data.get("value", []))
        url = data.get("@odata.nextLink") or data.get("odata.nextLink")
    return all_messages, fid


def _read_emails_via_playwright():
    """
    Read emails by automating OWA in a real browser.
    Intercepts the internal API calls OWA makes when loading the email list,
    so we get all email bodies without clicking each one individually.
    Requires: pip install playwright && playwright install chromium
    """
    # ── Fast path: skip the browser if we have a fresh cached OWA token ──────
    cached_token = _load_cached_owa_token()
    if cached_token:
        print("  ✓ Using cached OWA token — no browser needed.")
        try:
            messages, _src_fid = _fetch_emails_with_token(cached_token)
            print(f"  ✓ Retrieved {len(messages)} emails via cached token")
            cands = []
            seen_names_c = set()
            for msg in messages:
                body = (msg.get("Body") or msg.get("body") or {})
                body_text = body.get("Content") or body.get("content") or msg.get("BodyPreview") or msg.get("bodyPreview") or ""
                subj = msg.get("Subject") or msg.get("subject") or ""
                name = parse_email_for_name(subj, body_text)
                if name:
                    key = name.strip().lower()
                    if key not in seen_names_c:
                        seen_names_c.add(key)
                        cands.append({
                            "name":           name.strip(),
                            "item":           None,
                            "email_id":       msg.get("Id") or msg.get("id"),
                            "graph_token":    cached_token,
                            "owa_base":       "https://outlook.cloud.microsoft/api/beta",
                            "mailbox":        SUPPORT_MAILBOX,
                            "src_folder_id":  _src_fid,
                        })
            print(f"  ✓ {len(cands)} unique candidate(s) found")
            return cands
        except Exception as e:
            print(f"  ⚠  Cached token failed ({e}) — falling back to browser...")

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise Exception("playwright not installed — run: pip install playwright && playwright install chromium")

    # Suppress "Future exception was never retrieved / TargetClosedError" noise.
    # Asyncio prints these directly to stderr — must intercept via the exception handler.
    import asyncio as _asyncio, sys as _sys
    class _StderrFilter:
        """Wrapper around stderr that drops Playwright TargetClosedError noise."""
        def __init__(self, wrapped): self._w = wrapped
        def write(self, s):
            if ("TargetClosedError" in s or
                "Future exception was never retrieved" in s or
                "Target page, context or browser has been closed" in s):
                return
            self._w.write(s)
        def flush(self): self._w.flush()
        def __getattr__(self, a): return getattr(self._w, a)
    _sys.stderr = _StderrFilter(_sys.stderr)

    captured = []   # raw message dicts from intercepted API responses
    seen_ids = set()
    owa_api_token = [None]   # token used by OWA for outlook.cloud.microsoft/api calls
    owa_api_base  = ["https://outlook.cloud.microsoft/api/beta"]

    def on_request(req):
        """Capture the Bearer token OWA uses for its own API calls."""
        try:
            url = req.url.lower()
            if "outlook.cloud.microsoft/api" in url:
                auth = req.headers.get("authorization", "")
                if auth.startswith("Bearer "):
                    # Prefer tokens from mailbox/user API calls (not calendar/compliance/policy)
                    is_mailbox_call = any(x in url for x in [
                        "/users/", "/me/messages", "/me/mailfolders",
                        "mailfolder", "/mail", "/messages"
                    ])
                    is_junk_call = any(x in url for x in [
                        "calendar", "compliancepolicy", "label", "gettimeprofile",
                        "getpersonasettings", "getclassifrules"
                    ])
                    token = auth.replace("Bearer ", "")
                    if not owa_api_token[0]:
                        owa_api_token[0] = token
                    elif is_mailbox_call and not is_junk_call:
                        owa_api_token[0] = token
        except Exception:
            pass

    def on_response(resp):
        """Capture email data from OWA's internal API calls."""
        try:
            if resp.status != 200:
                return
            url = resp.url.lower()
            if not any(h in url for h in [
                "outlook.cloud.microsoft", "outlook.office365.com",
                "outlook.office.com", "graph.microsoft.com"
            ]):
                return
            data = resp.json()
            for m in data.get("value", []):
                mid = m.get("id", "")
                if mid and mid not in seen_ids:
                    seen_ids.add(mid)
                    captured.append(m)
            # (token capture — no debug output needed)
        except Exception:
            pass

    print("  Opening OWA in browser — leave it alone while it loads...")

    # Persistent profile dir — keeps login session and Edge extensions (incl. device compliance)
    PLAYWRIGHT_PROFILE = os.path.join(os.path.expanduser("~"), "tracker_edge_profile")

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=PLAYWRIGHT_PROFILE,
            channel="msedge",
            headless=False,
            slow_mo=100,
            args=["--disable-blink-features=AutomationControlled"],
            service_workers="block",  # force real network calls so we can capture Bearer token
        )
        page = context.new_page()
        page.on("request",  on_request)
        page.on("response", on_response)

        # Load OWA — persistent session means we'll be logged in already
        # Open OWA in the context of the support shared mailbox.
        # This ensures the folder tree and captured API tokens relate to
        # support@aeroprofessional.com, not Emily's personal mailbox.
        page.goto(f"https://outlook.office365.com/owa/{SUPPORT_MAILBOX}/", timeout=60000)
        page.wait_for_load_state("domcontentloaded")

        # Auto-detect login — wait up to 90s for OWA inbox to appear
        # If still on a login/auth page after that, prompt once for manual login
        print("  Waiting for OWA to load...")
        for _ in range(18):  # 18 × 5s = 90s max
            page.wait_for_timeout(5000)
            url = page.url.lower()
            if any(x in url for x in ["outlook.cloud.microsoft/mail", "outlook.office365.com/mail",
                                       "outlook.office.com/mail"]):
                print("  ✓ OWA loaded — already signed in.")
                break
            if any(x in url for x in ["login", "microsoftonline", "signin"]):
                # Auto-fill credentials if on a Microsoft login page
                try:
                    email_input = page.query_selector('input[type="email"], input[name="loginfmt"]')
                    if email_input and email_input.is_visible():
                        email_input.fill(EMAIL_LOGIN_USER)
                        page.keyboard.press("Enter")
                        page.wait_for_timeout(2000)
                    pwd_input = page.query_selector('input[type="password"], input[name="passwd"]')
                    if pwd_input and pwd_input.is_visible():
                        pwd_input.fill(EMAIL_PASSWORD)
                        page.keyboard.press("Enter")
                        page.wait_for_timeout(3000)
                    # Handle "Stay signed in?" prompt
                    kmsi = page.query_selector('input[value="Yes"], #idBtn_Back')
                    if kmsi and kmsi.is_visible():
                        kmsi.click()
                        page.wait_for_timeout(2000)
                    print("  ℹ  Auto-filled login credentials.")
                except Exception:
                    pass
                continue
        else:
            # Still not at inbox after 90s — try auto-login one more time then proceed
            try:
                pwd_input = page.query_selector('input[type="password"], input[name="passwd"]')
                if pwd_input and pwd_input.is_visible():
                    pwd_input.fill(EMAIL_PASSWORD)
                    page.keyboard.press("Enter")
                    page.wait_for_timeout(5000)
                    print("  ℹ  Auto re-authenticated.")
            except Exception:
                pass
            page.wait_for_timeout(10000)

        # Wait for OWA to settle and fire its mailbox API calls so we can capture the token
        # (calendar/compliance requests arrive first; mailbox requests arrive shortly after)
        page.wait_for_timeout(6000)

        print(f"  Current URL: {page.url}")

        # ── Fast path: if we already have the OWA API token, use it directly ──────
        # No need to navigate the folder tree — just call the API
        if owa_api_token[0]:
            print("  ✓ Have OWA token — fetching emails via API directly (no folder navigation)...")
            context.close()
            token = owa_api_token[0]
            _save_owa_token(token)  # cache it so next run skips the browser entirely
            hdrs  = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
            mb    = SUPPORT_MAILBOX
            base  = owa_api_base[0]
            # Find the folder
            fid = None
            for folders_url in [
                f"{base}/users/{mb}/MailFolders?$top=200&$select=Id,DisplayName",
                f"{base}/users/{mb}/MailFolders/Inbox/ChildFolders?$top=200&$select=Id,DisplayName",
            ]:
                try:
                    r = requests.get(folders_url, headers=hdrs, timeout=30)
                    if r.status_code == 200:
                        for f in r.json().get("value", []):
                            if f.get("DisplayName","").upper() == EMAIL_SUBFOLDER.upper():
                                fid = f.get("Id")
                                break
                    if fid:
                        break
                except Exception:
                    pass
            if fid:
                all_msgs = []
                url = (f"{base}/users/{mb}/MailFolders/{fid}/Messages"
                       f"?$top=500&$select=Id,Subject,Body,BodyPreview")
                while url:
                    try:
                        r = requests.get(url, headers=hdrs, timeout=30)
                        r.raise_for_status()
                        d = r.json()
                        all_msgs.extend(d.get("value", []))
                        url = d.get("@odata.nextLink") or d.get("odata.nextLink")
                    except Exception as e:
                        print(f"  ⚠ Fetch error: {e}")
                        break
                print(f"  ✓ Retrieved {len(all_msgs)} emails via OWA API")
                cands = []
                seen_names_local = set()
                for msg in all_msgs:
                    body_obj  = msg.get("Body") or msg.get("body") or {}
                    body_text = body_obj.get("Content","") if isinstance(body_obj, dict) else ""
                    body_text = body_text or msg.get("BodyPreview","") or msg.get("bodyPreview","")
                    subj = msg.get("Subject","") or msg.get("subject","")
                    name = parse_email_for_name(subj, body_text)
                    if name:
                        key = name.strip().lower()
                        if key not in seen_names_local:
                            seen_names_local.add(key)
                            cands.append({"name": name.strip(), "item": None,
                                          "email_id": msg.get("Id") or msg.get("id"),
                                          "graph_token": token, "owa_base": base,
                                          "mailbox": mb, "folder_id": fid})
                print(f"  ✓ {len(cands)} unique candidate(s) found")
                return cands
            else:
                print(f"  ⚠ Could not find '{EMAIL_SUBFOLDER}' via API — falling back to folder navigation")

        # Give OWA time to render the folder tree
        page.wait_for_timeout(3000)

        # Find the NEW REGS TO ACTION folder and click it
        print(f"  Looking for '{EMAIL_SUBFOLDER}' folder...")
        folder_clicked = False

        # Step 1: Find and expand the support shared mailbox first
        # Shared mailboxes appear at the bottom of the folder tree — scroll to find them
        support_short = SUPPORT_MAILBOX.split("@")[0].lower()  # e.g. "support"
        print(f"  Expanding shared mailbox ({SUPPORT_MAILBOX})...")
        for scroll_n in range(30):
            try:
                # Look for the shared mailbox label and click it to expand
                expanded = page.evaluate(f"""() => {{
                    const all = document.querySelectorAll('[aria-label], [title]');
                    for (const el of all) {{
                        const label = (el.getAttribute('aria-label') || el.getAttribute('title') || '').toLowerCase();
                        if (label.includes('{support_short}') && label.includes('@')) {{
                            // Click to expand if it has an expand arrow
                            const btn = el.querySelector('[aria-expanded]') || el.closest('[aria-expanded]');
                            if (btn && btn.getAttribute('aria-expanded') === 'false') {{
                                btn.click();
                                return 'expanded';
                            }}
                            return 'found';
                        }}
                    }}
                    return null;
                }}""")
                if expanded:
                    print(f"  ✓ Support mailbox {expanded}")
                    page.wait_for_timeout(1500)
                    break
            except Exception:
                pass
            # Scroll folder panel down to reveal more folders
            page.evaluate("""() => {
                const panels = document.querySelectorAll('nav, [role="navigation"], [role="tree"], [class*="folder"], [class*="nav"]');
                panels.forEach(p => { p.scrollTop += 300; });
                window.scrollBy(0, 200);
            }""")
            page.wait_for_timeout(400)

        # Step 2: Scroll within the folder panel until NEW REGS TO ACTION appears, then click it
        selectors = [
            f"[aria-label*='{EMAIL_SUBFOLDER}']",
            f"[title*='{EMAIL_SUBFOLDER}']",
            f"span:text-is('{EMAIL_SUBFOLDER}')",
            f"div:text-is('{EMAIL_SUBFOLDER}')",
            f"button:text-is('{EMAIL_SUBFOLDER}')",
        ]
        for scroll_attempt in range(40):
            # Try every selector at current scroll position
            for selector in selectors:
                try:
                    el = page.locator(selector).first
                    el.wait_for(timeout=500, state="visible")
                    el.scroll_into_view_if_needed()
                    el.click()
                    folder_clicked = True
                    print(f"  ✓ Clicked folder via selector: {selector}")
                    break
                except Exception:
                    continue
            if folder_clicked:
                break
            # Also try text match
            try:
                page.get_by_text(EMAIL_SUBFOLDER, exact=True).first.click(timeout=500)
                folder_clicked = True
                print(f"  ✓ Clicked folder via text match")
                break
            except Exception:
                pass
            # Scroll all navigational panels down
            page.evaluate("""() => {
                document.querySelectorAll('nav, [role="navigation"], [role="tree"], [role="group"]').forEach(el => {
                    el.scrollTop += 150;
                });
            }""")
            page.wait_for_timeout(300)

        if not folder_clicked:
            # Show what IS visible to help diagnose
            visible = page.evaluate("""() => {
                const all = document.querySelectorAll('[aria-label], [title]');
                const matches = [];
                all.forEach(el => {
                    const t = el.getAttribute('aria-label') || el.getAttribute('title') || '';
                    if (t.toUpperCase().includes('REGS') || t.toUpperCase().includes('ACTION') ||
                        t.toLowerCase().includes('support')) {
                        matches.push(t.trim());
                    }
                });
                return [...new Set(matches)].slice(0, 20);
            }""")
            context.close()
            raise Exception(
                f"Could not find '{EMAIL_SUBFOLDER}' in OWA. "
                "Make sure the support mailbox is visible in the left panel."
            )

        print("  Folder found — waiting for OWA to load folder contents...")

        # With service workers blocked, OWA MUST hit the network to load the folder.
        # Actively wait (up to 20s) for any outlook.cloud.microsoft API response.
        # This also fires on_request which captures the Bearer token.
        try:
            page.wait_for_response(
                lambda r: "outlook.cloud.microsoft/api" in r.url.lower() and r.status == 200,
                timeout=20000
            )
            print("  ✓ OWA API response received — token captured")
        except Exception:
            # No API response in 20s — OWA may have used in-memory data.
            # Force a folder reload by clicking it again.
            print("  ⚠ No OWA response in 20s — forcing folder reload...")
            try:
                for sel in [f"[title*='{EMAIL_SUBFOLDER}']",
                            f"[aria-label*='{EMAIL_SUBFOLDER}']"]:
                    try:
                        page.locator(sel).first.click(timeout=3000)
                        break
                    except Exception:
                        pass
                page.wait_for_response(
                    lambda r: "outlook.cloud.microsoft/api" in r.url.lower() and r.status == 200,
                    timeout=15000
                )
                print("  ✓ OWA API response received on second attempt")
            except Exception:
                print("  ⚠ Still no OWA API response — proceeding with MSAL fallback")

        # Scroll the email list to load all emails (OWA uses virtual scrolling)
        for scroll_n in range(200):   # cap at 200 scrolls ≈ ~10,000 emails
            prev = len(captured)
            page.evaluate("""
                () => {
                    const list = document.querySelector('[role="list"], [role="listbox"]');
                    if (list) list.scrollTop += list.clientHeight * 2;
                }
            """)
            page.wait_for_timeout(800)
            if len(captured) == prev and scroll_n > 2:
                break   # no new emails loaded — we're at the bottom

        # ── Use OWA API token captured from request headers ───────────────────
        print("  Checking for OWA API token...")
        use_token   = owa_api_token[0]
        use_base    = owa_api_base[0]
        use_mailbox = f"support@aeroprofessional.com"

        if not use_token:
            # Fall back to MSAL localStorage — prefer OWA token over Graph token
            print("  No OWA request token — trying MSAL localStorage...")
            use_token = page.evaluate("""() => {
                let owaToken = null, graphToken = null;
                let owaExp = 0, graphExp = 0;
                for (const store of [localStorage, sessionStorage]) {
                    try {
                        for (const key of Object.keys(store)) {
                            if (!key.toLowerCase().includes('accesstoken')) continue;
                            try {
                                const v = JSON.parse(store.getItem(key));
                                if (!v || !v.secret) continue;
                                const target = (v.target || '').toLowerCase();
                                const exp = parseInt(v.expiresOn || '0');
                                if (target.includes('outlook.cloud.microsoft') ||
                                    target.includes('outlook.office')) {
                                    if (exp > owaExp) { owaExp = exp; owaToken = v.secret; }
                                } else if (target.includes('graph.microsoft.com')) {
                                    if (exp > graphExp) { graphExp = exp; graphToken = v.secret; }
                                }
                            } catch(e) {}
                        }
                    } catch(e) {}
                }
                // Return OWA token if found (works for support mailbox), else Graph
                if (owaToken) return JSON.stringify({token: owaToken, type: 'owa'});
                if (graphToken) return JSON.stringify({token: graphToken, type: 'graph'});
                return null;
            }""")
            if use_token:
                try:
                    parsed = json.loads(use_token)
                    if parsed.get("type") == "owa":
                        use_token = parsed["token"]
                        use_base  = "https://outlook.cloud.microsoft/api/beta"
                        print("  ✓ Got OWA token from MSAL cache")
                    else:
                        use_token = parsed["token"]
                        use_base  = "https://graph.microsoft.com/v1.0"
                        print("  ✓ Got Graph token from MSAL cache (may 403)")
                except Exception:
                    use_base = "https://graph.microsoft.com/v1.0"

        if use_token and owa_api_token[0]:
            _save_owa_token(owa_api_token[0])  # cache intercepted token for next run

        if use_token:
            is_owa_api = "outlook.cloud.microsoft" in use_base
            api_label  = "OWA internal API" if is_owa_api else "Graph API"
            print(f"  ✓ Got token — reading emails via {api_label}...")
            # Keep context open — we may need DOM fallback if API 403s

            hdrs = {"Authorization": f"Bearer {use_token}",
                    "Accept": "application/json"}
            mb   = "support@aeroprofessional.com"

            if is_owa_api:
                # ── OWA internal API (outlook.cloud.microsoft/api/beta) ──────────
                # First find the folder ID
                folders_url = f"{use_base}/users/{mb}/MailFolders?$top=200&$select=Id,DisplayName"
                fid = None
                try:
                    r = requests.get(folders_url, headers=hdrs, timeout=30)
                    r.raise_for_status()
                    for f in r.json().get("value", []):
                        if f.get("DisplayName","").upper() == EMAIL_SUBFOLDER.upper():
                            fid = f.get("Id")
                            break
                    if not fid:
                        # Try child folders of Inbox
                        inbox_url = f"{use_base}/users/{mb}/MailFolders/Inbox/ChildFolders?$top=200&$select=Id,DisplayName"
                        r = requests.get(inbox_url, headers=hdrs, timeout=30)
                        r.raise_for_status()
                        for f in r.json().get("value", []):
                            if f.get("DisplayName","").upper() == EMAIL_SUBFOLDER.upper():
                                fid = f.get("Id")
                                break
                except Exception as e:
                    print(f"  ⚠ OWA folder search error: {e}")

                if not fid:
                    # Also try alternative paths (root folders, msgfolderroot, v2.1)
                    for alt_base in ["https://outlook.cloud.microsoft/api/v2.1",
                                     "https://graph.microsoft.com/v1.0"]:
                        for alt_path in [
                            f"{alt_base}/users/{mb}/MailFolders?$top=200&$select=Id,DisplayName",
                            f"{alt_base}/users/{mb}/MailFolders/msgfolderroot/ChildFolders?$top=200&$select=Id,DisplayName",
                            f"{alt_base}/users/{mb}/MailFolders/Inbox/ChildFolders?$top=200&$select=Id,DisplayName",
                        ]:
                            try:
                                r2 = requests.get(alt_path, headers=hdrs, timeout=20)
                                if r2.status_code == 200:
                                    for f2 in r2.json().get("value", []):
                                        if f2.get("DisplayName","").upper() == EMAIL_SUBFOLDER.upper():
                                            fid = f2.get("Id")
                                            use_base = alt_base
                                            break
                                if fid:
                                    break
                            except Exception:
                                pass
                        if fid:
                            break

                if not fid:
                    print(f"  ⚠ Could not find folder via OWA API — trying captured DOM items")
                    context_closed = True
                else:
                    # Fetch all messages from the folder
                    all_msgs = []
                    url = (f"{use_base}/users/{mb}/MailFolders/{fid}/Messages"
                           f"?$top=500&$select=Id,Subject,Body,BodyPreview")
                    while url:
                        try:
                            r = requests.get(url, headers=hdrs, timeout=30)
                            r.raise_for_status()
                            d = r.json()
                            all_msgs.extend(d.get("value", []))
                            url = d.get("@odata.nextLink") or d.get("odata.nextLink")
                        except Exception as e:
                            print(f"  ⚠ OWA message fetch error: {e}")
                            break
                    print(f"  ✓ Retrieved {len(all_msgs)} emails via {api_label}")
                    cands = []
                    seen_names_local = set()
                    for msg in all_msgs:
                        body = (msg.get("Body") or msg.get("body") or {})
                        body_text = body.get("Content","") if isinstance(body, dict) else ""
                        body_text = body_text or msg.get("BodyPreview","") or msg.get("bodyPreview","")
                        subj = msg.get("Subject","") or msg.get("subject","")
                        name = parse_email_for_name(subj, body_text)
                        if name:
                            key = name.strip().lower()
                            if key not in seen_names_local:
                                seen_names_local.add(key)
                                cands.append({
                                    "name":        name.strip(),
                                    "item":        None,
                                    "email_id":    msg.get("Id") or msg.get("id"),
                                    "graph_token": use_token,
                                    "owa_base":    use_base,
                                    "mailbox":     mb,
                                    "folder_id":   fid,
                                })
                    print(f"  ✓ {len(cands)} unique candidate(s) found")
                    return cands
            else:
                # ── Graph API path ───────────────────────────────────────────────
                fid = _graph_find_folder(use_token, display_name=EMAIL_SUBFOLDER, mailbox=mb)
                if not fid:
                    raise Exception(f"Graph API: folder '{EMAIL_SUBFOLDER}' not found")
                print(f"  ✓ Found folder via Graph API")
                all_msgs = []
                url = (f"https://graph.microsoft.com/v1.0/users/{mb}"
                       f"/mailFolders/{fid}/messages"
                       f"?$top=999&$select=id,subject,body,bodyPreview")
                while url:
                    r = requests.get(url, headers=hdrs, timeout=30)
                    r.raise_for_status()
                    d = r.json()
                    all_msgs.extend(d.get("value", []))
                    url = d.get("@odata.nextLink")
                print(f"  ✓ Retrieved {len(all_msgs)} emails via Graph API (OWA token)")
                cands = []
                seen_names_local = set()
                for msg in all_msgs:
                    body = (msg.get("body") or {}).get("content", "") or msg.get("bodyPreview", "")
                    name = parse_email_for_name(msg.get("subject", ""), body)
                    if name:
                        key = name.strip().lower()
                        if key not in seen_names_local:
                            seen_names_local.add(key)
                            cands.append({
                                "name":        name.strip(),
                                "item":        None,
                                "email_id":    msg.get("id"),
                                "graph_token": use_token,
                            })
                print(f"  ✓ {len(cands)} unique candidate(s) found")
                return cands
        else:
            print("  ⚠ Could not extract token — will parse from DOM previews only")
            context.close()

    # Parse candidate names from captured messages
    candidates = []
    seen_names = set()
    for msg in captured:
        body = (msg.get("body") or {}).get("content", "") or msg.get("bodyPreview", "")
        subj = msg.get("subject", "")
        name = parse_email_for_name(subj, body)
        if name:
            key = name.strip().lower()
            if key not in seen_names:
                seen_names.add(key)
                candidates.append({
                    "name":        name.strip(),
                    "item":        None,
                    "email_id":    msg.get("id"),
                    "graph_token": None,   # can't move via Graph without consent
                })

    return candidates


def _read_emails_via_imap():
    """
    Read registration emails from the REGS folder via IMAP.
    Works headlessly on Linux/GitHub Actions.
    Requires Basic Authentication to be enabled for the M365 account.
    """
    import imaplib
    import email as _em_lib
    from email.header import decode_header as _dh

    IMAP_SERVER = "outlook.office365.com"
    IMAP_PORT   = 993

    # Try direct login first, then delegate format for shared mailbox
    conn = None
    last_err = None
    for username in [EMAIL_LOGIN_USER, f"{EMAIL_LOGIN_USER}\\{SUPPORT_MAILBOX}"]:
        try:
            _c = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
            _c.login(username, EMAIL_PASSWORD)
            conn = _c
            break
        except Exception as e:
            last_err = e
    if conn is None:
        raise Exception(f"IMAP login failed: {last_err}")

    # Try common folder path formats
    selected_folder = None
    for path in [EMAIL_SUBFOLDER, f"INBOX/{EMAIL_SUBFOLDER}", f"Inbox/{EMAIL_SUBFOLDER}",
                 f"INBOX.{EMAIL_SUBFOLDER}"]:
        try:
            typ, _ = conn.select(f'"{path}"')
            if typ == "OK":
                selected_folder = path
                break
        except Exception:
            pass
    if selected_folder is None:
        conn.logout()
        raise Exception(
            f"IMAP: could not find folder '{EMAIL_SUBFOLDER}'. "
            "Check that Basic Auth IMAP is enabled in M365 admin and that the folder exists."
        )

    typ, data = conn.search(None, "ALL")
    if typ != "OK" or not data[0]:
        conn.logout()
        return []  # empty folder

    candidates = []
    seen_names = set()
    for uid_bytes in data[0].split():
        try:
            uid = uid_bytes.decode()
            typ2, msg_data = conn.fetch(uid_bytes, "(RFC822)")
            raw = msg_data[0][1]
            msg = _em_lib.message_from_bytes(raw)

            # Decode subject
            subj = ""
            for part, charset in _dh(msg.get("Subject", "") or ""):
                if isinstance(part, bytes):
                    subj += part.decode(charset or "utf-8", errors="replace")
                else:
                    subj += str(part)

            # Decode body (prefer plain text)
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        try:
                            body = part.get_payload(decode=True).decode(
                                part.get_content_charset() or "utf-8", errors="replace")
                            break
                        except Exception:
                            pass
            else:
                try:
                    body = msg.get_payload(decode=True).decode(
                        msg.get_content_charset() or "utf-8", errors="replace")
                except Exception:
                    pass

            name = parse_email_for_name(subj, body)
            if name:
                key = name.strip().lower()
                if key not in seen_names:
                    seen_names.add(key)
                    candidates.append({
                        "name":          name.strip(),
                        "item":          None,
                        "email_id":      f"imap:{uid}",
                        "imap_uid":      uid,
                        "imap_folder":   selected_folder,
                        "graph_token":   None,
                    })
        except Exception:
            continue

    conn.logout()
    print(f"  ✓ {len(candidates)} unique candidate(s) found via IMAP")
    return candidates


def _move_email_via_imap(imap_uid, src_folder, dest_folder):
    """Move an email by UID from src_folder to dest_folder via IMAP."""
    import imaplib
    IMAP_SERVER = "outlook.office365.com"
    conn = imaplib.IMAP4_SSL(IMAP_SERVER, 993)
    try:
        conn.login(EMAIL_LOGIN_USER, EMAIL_PASSWORD)
        conn.select(f'"{src_folder}"')
        # Exchange Online supports IMAP MOVE extension
        try:
            conn.uid("MOVE", imap_uid, f'"{dest_folder}"')
            return True
        except imaplib.IMAP4.error:
            pass
        # Fallback: COPY + mark deleted + expunge
        typ, _ = conn.uid("COPY", imap_uid, f'"{dest_folder}"')
        if typ == "OK":
            conn.uid("STORE", imap_uid, "+FLAGS", r"(\Deleted)")
            conn.expunge()
            return True
        return False
    finally:
        try:
            conn.logout()
        except Exception:
            pass


def read_candidates_from_email():
    """
    Read ALL emails from the configured source folder.
    Returns list of dicts: [{"name": str, "item": COM_item_or_None, "email_id": str_or_None, "graph_token": str_or_None}]
    Deduplicates by name.

    Priority:
      0. IMAP (headless — used on Linux/GitHub Actions)
      1. Microsoft Graph API (works for any Exchange folder, even uncached ones)
      2. Playwright OWA browser automation (no admin needed)
      3. Outlook COM (local cache — may miss some folders)
    """
    # ── Method 0: IMAP (headless — GitHub Actions / Linux / non-Windows) ────────
    import platform as _platform
    if _platform.system() != "Windows":
        try:
            candidates = _read_emails_via_imap()
            if candidates:
                return candidates
        except Exception as _imap_e:
            print(f"  ⚠  IMAP: {_imap_e}")

    # ── Method 1: Microsoft Graph API ─────────────────────────────────────────
    if GRAPH_TENANT_ID and GRAPH_CLIENT_ID:
        try:
            messages, token, src_fid = _read_emails_via_graph()
            candidates = []
            seen = set()
            for msg in messages:
                body = (msg.get("body") or {}).get("content", "") or msg.get("bodyPreview", "")
                subj = msg.get("subject", "")
                name = parse_email_for_name(subj, body)
                if name:
                    key = name.strip().lower()
                    if key not in seen:
                        seen.add(key)
                        candidates.append({
                            "name":           name.strip(),
                            "item":           None,
                            "email_id":       msg["id"],
                            "graph_token":    token,
                            "src_folder_id":  src_fid,   # used to find/create Done subfolder
                        })
            print(f"  ✓ {len(candidates)} unique candidate(s) found")
            return candidates
        except Exception as e:
            print(f"  ⚠  Graph API: {e}")

    # ── Method 2: Playwright OWA browser automation ────────────────────────────
    try:
        candidates = _read_emails_via_playwright()
        seen = set()
        deduped = []
        for c in candidates:
            key = c["name"].lower()
            if key not in seen:
                seen.add(key)
                deduped.append(c)
        if deduped:
            print(f"  ✓ {len(deduped)} unique candidate(s) found via OWA")
            return deduped
    except Exception as e:
        print(f"  ⚠  OWA browser: {e}")

    # ── Method 3: Outlook COM ──────────────────────────────────────────────────
    try:
        folder = _get_source_folder()
        items  = folder.Items
        candidates = []
        seen = set()
        for item in items:
            try:
                name = parse_email_for_name(item.Subject or "", item.Body or "")
                if name:
                    key = name.strip().lower()
                    if key not in seen:
                        seen.add(key)
                        try:
                            _eid  = item.EntryID
                            _sid  = item.Parent.StoreID
                        except Exception:
                            _eid, _sid = None, None
                        candidates.append({"name": name.strip(), "item": item,
                                           "entry_id": _eid, "store_id": _sid,
                                           "email_id": None, "graph_token": None})
            except Exception:
                continue
        print(f"  ✓ Found {len(candidates)} candidate(s) via Outlook COM")
        return candidates
    except Exception as e:
        print(f"  ⚠  Outlook COM: {e}")

    return []

# ── Microsoft Graph API helpers ────────────────────────────────────────────────

def get_graph_token():
    """
    Authenticate to Microsoft Graph.

    Two modes:
      • App-only (GitHub Actions): when GRAPH_CLIENT_SECRET is set, uses
        ConfidentialClientApplication with client credentials — no user interaction.
        Requires application permissions (Mail.Read, Mail.ReadWrite, Mail.Send)
        granted in Azure AD (not delegated permissions).
      • Device code (local): when GRAPH_CLIENT_SECRET is not set, prompts the
        user to sign in once via browser. Token is cached for future runs.
    """
    try:
        import msal
    except ImportError:
        raise Exception("msal not installed — run: pip install msal --break-system-packages")
    if not GRAPH_TENANT_ID or not GRAPH_CLIENT_ID:
        raise Exception("GRAPH_TENANT_ID / GRAPH_CLIENT_ID not configured (check GitHub Secrets)")

    # ── App-only / client credentials (GitHub Actions) ────────────────────────
    if GRAPH_CLIENT_SECRET:
        app = msal.ConfidentialClientApplication(
            GRAPH_CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}",
            client_credential=GRAPH_CLIENT_SECRET,
        )
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if result and "access_token" in result:
            return result["access_token"]
        raise Exception(f"Graph client-credentials auth failed: {result.get('error_description', result)}")

    # ── Device code flow (interactive — local first-time setup) ───────────────
    cache = msal.SerializableTokenCache()
    if os.path.exists(GRAPH_TOKEN_FILE):
        cache.deserialize(open(GRAPH_TOKEN_FILE).read())
    app = msal.PublicClientApplication(
        GRAPH_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}",
        token_cache=cache,
    )
    GRAPH_SCOPES = ["https://graph.microsoft.com/Mail.Read",
                    "https://graph.microsoft.com/Mail.ReadWrite",
                    "https://graph.microsoft.com/Mail.Send"]
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
        if result and "access_token" in result:
            if cache.has_state_changed:
                open(GRAPH_TOKEN_FILE, "w").write(cache.serialize())
            return result["access_token"]
    # Device code flow — user signs in once via browser
    flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)
    if "user_code" not in flow:
        raise Exception(f"Device flow failed: {flow}")
    print(f"\n  Microsoft sign-in required:")
    print(f"  1. Open: {flow['verification_uri']}")
    print(f"  2. Enter code: {flow['user_code']}")
    print(f"  Waiting for you to sign in (up to 15 minutes)...", flush=True)
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise Exception(f"Auth failed: {result.get('error_description', result)}")
    if cache.has_state_changed:
        open(GRAPH_TOKEN_FILE, "w").write(cache.serialize())
    return result["access_token"]

def _graph_find_folder(token, parent_folder_id=None, display_name=None, mailbox=None):
    """Return the ID of a named child folder under parent_folder_id (or inbox root).
    When no parent is given, searches inbox children AND top-level root folders."""
    mb   = mailbox or SUPPORT_MAILBOX
    hdrs = {"Authorization": f"Bearer {token}"}
    target = (display_name or "").strip().upper()

    def _search_url(url):
        while url:
            try:
                r = requests.get(url, headers=hdrs, timeout=15)
                r.raise_for_status()
                data = r.json()
            except Exception:
                break
            for f in data.get("value", []):
                if f.get("displayName", "").strip().upper() == target:
                    return f["id"]
            url = data.get("@odata.nextLink")
        return None

    if parent_folder_id:
        return _search_url(
            f"https://graph.microsoft.com/v1.0/users/{mb}/mailFolders/{parent_folder_id}/childFolders?$top=100"
        )
    # No parent — try inbox children first, then root-level folders
    fid = _search_url(
        f"https://graph.microsoft.com/v1.0/users/{mb}/mailFolders/inbox/childFolders?$top=100"
    )
    if fid:
        return fid
    # Try top-level mailbox folders (Inbox, Sent, and siblings like "New regs")
    return _search_url(
        f"https://graph.microsoft.com/v1.0/users/{mb}/mailFolders?$top=100"
    )

def _graph_ensure_folder(token, path_parts, mailbox=None):
    """Navigate/create a folder path (list of names from inbox root). Returns final folder ID."""
    mb   = mailbox or SUPPORT_MAILBOX
    hdrs = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    parent_id = None
    for part in path_parts:
        fid = _graph_find_folder(token, parent_folder_id=parent_id, display_name=part, mailbox=mb)
        if not fid:
            base = (f"https://graph.microsoft.com/v1.0/users/{mb}/mailFolders/"
                    + (f"{parent_id}/childFolders" if parent_id else "inbox/childFolders"))
            r = requests.post(base, json={"displayName": part}, headers=hdrs, timeout=15)
            r.raise_for_status()
            fid = r.json()["id"]
        parent_id = fid
    return parent_id

def _read_emails_via_graph():
    """Read all emails from EMAIL_SUBFOLDER in the support mailbox. Returns (messages, token, folder_id)."""
    token = get_graph_token()
    hdrs  = {"Authorization": f"Bearer {token}"}
    mb    = SUPPORT_MAILBOX
    fid   = _graph_find_folder(token, display_name=EMAIL_SUBFOLDER)
    if not fid:
        raise Exception(f"Folder '{EMAIL_SUBFOLDER}' not found under inbox in {mb}")
    print(f"  ✓ Found '{EMAIL_SUBFOLDER}' via Microsoft Graph API")
    messages = []
    url = (f"https://graph.microsoft.com/v1.0/users/{mb}"
           f"/mailFolders/{fid}/messages"
           f"?$top=999&$select=id,subject,body,bodyPreview")
    while url:
        r = requests.get(url, headers=hdrs, timeout=30)
        r.raise_for_status()
        data = r.json()
        messages.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    print(f"  ✓ Retrieved {len(messages)} emails from folder")
    return messages, token, fid

def _move_email_via_graph(token, message_id, owa_base=None, mailbox=None, folder_id=None, dest_folder=None):
    """Move an email to dest_folder (defaults to EMAIL_DONE_FOLDER) via Graph API or OWA internal API."""
    if dest_folder is None:
        dest_folder = EMAIL_DONE_FOLDER
    mb   = mailbox or SUPPORT_MAILBOX
    hdrs = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    if owa_base and "outlook.cloud.microsoft" in owa_base:
        # OWA internal API — find destination folder then move
        parts = [p.strip() for p in dest_folder.split("/")]
        dest_fid = None

        def _owa_list_folders(url):
            """Return list of folder dicts from a MailFolders URL, or []."""
            try:
                r = requests.get(url, headers=hdrs, timeout=15)
                if r.status_code == 200:
                    return r.json().get("value", [])
            except Exception:
                pass
            return []

        def _owa_name_match(f_name, target):
            f, t = f_name.lower(), target.lower()
            return f == t or t in f or f in t

        def _owa_find_child(parent_id, target_name):
            url = f"{owa_base}/users/{mb}/MailFolders/{parent_id}/ChildFolders?$top=200&$select=Id,DisplayName,ChildFolderCount"
            for cf in _owa_list_folders(url):
                if _owa_name_match(cf.get("DisplayName", ""), target_name):
                    return cf.get("Id")
            return None

        # Comprehensive search: top-level, Inbox children, then all top-level children
        # Collect all folders we might want to search for the FIRST path part
        all_search_urls = [
            f"{owa_base}/users/{mb}/MailFolders?$top=200&$select=Id,DisplayName,ChildFolderCount",
            f"{owa_base}/users/{mb}/MailFolders/Inbox/ChildFolders?$top=200&$select=Id,DisplayName,ChildFolderCount",
            f"{owa_base}/users/{mb}/MailFolders/msgfolderroot/ChildFolders?$top=200&$select=Id,DisplayName,ChildFolderCount",
        ]
        parent_fid = None
        for url in all_search_urls:
            for f in _owa_list_folders(url):
                if _owa_name_match(f.get("DisplayName", ""), parts[0]):
                    parent_fid = f.get("Id")
                    break
            if parent_fid:
                break

        if parent_fid:
            if len(parts) == 1:
                dest_fid = parent_fid
            else:
                dest_fid = _owa_find_child(parent_fid, parts[1])
                if not dest_fid:
                    # Try searching grandchildren of top-level too
                    for url in all_search_urls:
                        for f in _owa_list_folders(url):
                            child_id = _owa_find_child(f.get("Id",""), parts[1])
                            if child_id:
                                dest_fid = child_id
                                break
                        if dest_fid:
                            break

        _subfolder_name = dest_folder.split("/")[-1].strip() if dest_folder else "Done"
        if not dest_fid and folder_id:
            # Destination folder not found — fall back to subfolder under source folder
            dest_fid = _owa_find_child(folder_id, _subfolder_name)
            if not dest_fid:
                # Create subfolder under the source folder via OWA API
                try:
                    _cr = requests.post(
                        f"{owa_base}/users/{mb}/MailFolders/{folder_id}/ChildFolders",
                        json={"DisplayName": _subfolder_name},
                        headers=hdrs, timeout=15,
                    )
                    if _cr.status_code in (200, 201):
                        dest_fid = _cr.json().get("Id")
                        print(f"  ℹ  Created '{_subfolder_name}' folder under source folder (OWA)")
                except Exception:
                    pass

        if not dest_fid:
            return None  # silently — caller will log to pending_moves

        r = requests.post(
            f"{owa_base}/users/{mb}/Messages/{message_id}/move",
            json={"DestinationId": dest_fid},
            headers=hdrs, timeout=15,
        )
        r.raise_for_status()
        return True
    else:
        # Standard Graph API
        # If we know the source folder ID, create/find "Done" directly under it
        # to avoid hunting for EMAIL_DONE_FOLDER which may not exist.
        _sf_name = dest_folder.split("/")[-1].strip() if dest_folder else "Done"
        if folder_id:
            dest_fid = _graph_find_folder(token, parent_folder_id=folder_id,
                                           display_name=_sf_name, mailbox=mb)
            if not dest_fid:
                # Create it
                try:
                    r = requests.post(
                        f"https://graph.microsoft.com/v1.0/users/{mb}/mailFolders/{folder_id}/childFolders",
                        json={"displayName": _sf_name},
                        headers=hdrs, timeout=15,
                    )
                    r.raise_for_status()
                    dest_fid = r.json()["id"]
                    print(f"  ℹ  Created '{_sf_name}' subfolder under source folder")
                except Exception as _ce:
                    raise
        else:
            parts    = [p.strip() for p in dest_folder.split("/")]
            dest_fid = _graph_ensure_folder(token, parts, mailbox=mb)
        r = requests.post(
            f"https://graph.microsoft.com/v1.0/users/{mb}/messages/{message_id}/move",
            json={"destinationId": dest_fid},
            headers=hdrs, timeout=15,
        )
        r.raise_for_status()
        return True
def move_to_done(cand, dest_folder=None):
    """Move a processed email to dest_folder (defaults to EMAIL_DONE_FOLDER) via OWA API then COM fallback."""
    if dest_folder is None:
        dest_folder = EMAIL_DONE_FOLDER
    email_id    = cand.get("email_id")    if isinstance(cand, dict) else None
    graph_token = cand.get("graph_token") if isinstance(cand, dict) else None
    owa_base    = cand.get("owa_base")    if isinstance(cand, dict) else None
    mailbox     = cand.get("mailbox")     if isinstance(cand, dict) else None
    folder_id   = cand.get("folder_id")  if isinstance(cand, dict) else None
    email_item  = cand.get("item")        if isinstance(cand, dict) else cand

    src_folder_id = cand.get("src_folder_id") if isinstance(cand, dict) else None
    if email_id and graph_token:
        for _attempt in range(2):  # attempt 0 = original token, attempt 1 = fresh token
            _token_to_use = graph_token
            if _attempt == 1:
                # Token likely expired — try to get a fresh one silently via MSAL cache
                try:
                    _token_to_use = get_graph_token()
                    if isinstance(cand, dict):
                        cand["graph_token"] = _token_to_use  # update for future moves
                    print(f"  ℹ  Token refreshed — retrying move...")
                except Exception:
                    break  # can't refresh, give up on API move
            try:
                result = _move_email_via_graph(_token_to_use, email_id,
                                               owa_base=owa_base, mailbox=mailbox,
                                               folder_id=folder_id or src_folder_id,
                                               dest_folder=dest_folder)
                if result is True:
                    return True   # API move worked
                # result is None = folder not found — no point retrying with fresh token
                print(f"  ⚠ Graph move: folder not found — trying COM fallback")
                break
            except Exception as _mv_ex:
                _is_auth_error = "401" in str(_mv_ex) or "403" in str(_mv_ex) or "Unauthorized" in str(_mv_ex)
                if _attempt == 0 and _is_auth_error:
                    continue  # token expired — loop back and refresh
                print(f"  ⚠ Graph move failed: {_mv_ex} — trying COM fallback")
                break

    # COM fallback (works when Outlook is open and email was from COM source)
    entry_id = cand.get("entry_id") if isinstance(cand, dict) else None
    store_id = cand.get("store_id") if isinstance(cand, dict) else None
    if email_item or entry_id:
        try:
            try:
                import pythoncom as _pc
                _pc.CoInitialize()
            except Exception:
                pass
            with _outlook_lock:
                # Re-acquire item from EntryID to avoid stale COM references
                if entry_id:
                    try:
                        import win32com.client as _wc
                        _ns = _wc.Dispatch("Outlook.Application").GetNamespace("MAPI")
                        email_item = _ns.GetItemFromID(entry_id, store_id)
                    except Exception:
                        pass  # fall through to original item reference
                if email_item:
                    email_item.Move(_get_done_folder(folder_path=dest_folder))
                    return True
        except Exception as e:
            print(f"  ⚠ COM move failed: {e}")

    # IMAP fallback (headless — GitHub Actions / Linux)
    imap_uid    = cand.get("imap_uid")    if isinstance(cand, dict) else None
    imap_folder = cand.get("imap_folder") if isinstance(cand, dict) else None
    if imap_uid and imap_folder:
        try:
            return _move_email_via_imap(imap_uid, imap_folder, dest_folder)
        except Exception as _ie:
            print(f"  ⚠ IMAP move failed: {_ie}")

    print(f"  ⚠ Could not move email to done folder (no working method)")
    return False

# Regional terms that should never appear as a candidate's nationality skill
REGIONAL_TERMS = {
    "europe", "european", "asia", "middle east", "africa", "americas",
    "north america", "south america", "caribbean", "oceania", "australasia",
    "gcc", "mena", "apac", "emea", "latam", "southeast asia", "central asia",
    "eastern europe", "western europe", "north africa", "sub-saharan africa",
    "gulf", "levant", "scandinavia", "balkans", "worldwide", "global",
    "international", "all", "any",
}

KNOWN_COUNTRY_NAMES_EXTRA = {
    # Countries that may not appear in country_skills_set if Tracker stored them
    # under a different area ID, but we still need to recognise as country names.
    "malta", "cyprus", "luxembourg", "liechtenstein", "monaco", "andorra",
    "canada", "australia", "new zealand", "fiji", "maldives", "iceland",
    "switzerland", "austria", "portugal", "finland", "denmark", "norway",
    "sweden", "belgium", "netherlands", "poland", "czech republic",
    "slovakia", "hungary", "romania", "bulgaria", "serbia", "serbia and montenegro",
    "croatia", "slovenia", "bosnia and herzegovina", "north macedonia",
    "albania", "kosovo", "montenegro", "moldova", "belarus", "latvia",
    "lithuania", "estonia", "georgia", "armenia", "azerbaijan",
    "kazakhstan", "uzbekistan", "kyrgyzstan", "tajikistan", "turkmenistan",
    "mongolia", "myanmar", "cambodia", "laos", "brunei", "timor-leste",
    "papua new guinea", "solomon islands", "vanuatu", "tonga", "samoa",
    "kiribati", "nauru", "tuvalu", "palau", "micronesia", "marshall islands",
    "cuba", "haiti", "dominican republic", "jamaica", "trinidad and tobago",
    "barbados", "bahamas", "belize", "guatemala", "honduras", "el salvador",
    "nicaragua", "costa rica", "panama", "colombia", "venezuela", "ecuador",
    "peru", "bolivia", "paraguay", "uruguay", "chile", "argentina", "brazil",
    "guyana", "suriname", "mauritius", "seychelles", "comoros", "djibouti",
    "eritrea", "somalia", "rwanda", "burundi", "malawi", "zambia", "mozambique",
    "zimbabwe", "botswana", "namibia", "lesotho", "swaziland", "eswatini",
    "cameroon", "senegal", "mali", "burkina faso", "guinea", "sierra leone",
    "liberia", "ivory coast", "ghana", "togo", "benin", "niger", "chad",
    "central african republic", "congo", "democratic republic of congo",
    "gabon", "equatorial guinea", "sao tome and principe", "cape verde",
    "gambia", "guinea-bissau", "angola", "madagascar", "reunion", "mayotte",
    "spain", "france", "germany", "italy", "united kingdom", "ireland",
    "united states", "usa", "united arab emirates", "uae", "saudi arabia",
    "kuwait", "qatar", "bahrain", "oman", "jordan", "lebanon", "syria",
    "iraq", "iran", "turkey", "israel", "palestine", "egypt", "libya",
    "algeria", "morocco", "tunisia", "sudan", "south sudan", "ethiopia",
    "kenya", "tanzania", "uganda", "nigeria", "ghana", "south africa",
    "india", "pakistan", "bangladesh", "nepal", "sri lanka", "china",
    "japan", "south korea", "north korea", "korea (south)", "korea (north)",
    "taiwan", "hong kong", "philippines", "indonesia", "malaysia", "singapore",
    "thailand", "vietnam", "laos", "cambodia", "russia", "ukraine",
    "mexico", "colombia", "venezuela", "peru", "chile", "argentina",
}

NATIONALITY_ADJECTIVES = {
    "british", "english", "scottish", "welsh", "irish",
    "american", "emirati", "saudi", "saudi arabian", "indian", "pakistani",
    "bangladeshi", "filipino", "indonesian", "malaysian",
    "australian", "canadian", "egyptian", "jordanian",
    "lebanese", "kuwaiti", "qatari", "bahraini", "omani",
    "yemeni", "iraqi", "turkish", "iranian", "russian",
    "chinese", "japanese", "korean", "south korean", "north korean", "thai", "vietnamese",
    "french", "german", "spanish", "italian", "greek",
    "dutch", "belgian", "swedish", "norwegian", "danish",
    "polish", "romanian", "ukrainian", "moroccan", "algerian",
    "tunisian", "nigerian", "kenyan", "ghanaian", "ugandan",
    "ethiopian", "tanzanian", "south african", "zimbabwean",
    "nepalese", "nepali", "sri lankan", "singaporean",
    "palestinian", "syrian", "libyan", "sudanese",
    "serbian", "bosnian", "montenegrin", "macedonian",
    "croatian", "slovenian", "albanian", "kosovar",
    "azerbaijani", "kazakhstani", "uzbek", "georgian", "armenian",
    "new zealander", "trinidadian", "jamaican", "barbadian",
    "colombian", "venezuelan", "ecuadorian", "peruvian", "chilean", "argentinian", "brazilian",
    "ghanaian", "zambian", "rwandan", "ugandan", "senegalese", "cameroonian",
    "hong konger",
}


def is_profile_complete(rec, country_skills_set=None):
    """Return True if the profile already has job title, work type, and correctly-formatted skills."""
    job_title_str = (rec.get("jobTitle") or "").strip()

    # Placeholder / generic titles are treated as missing
    _PLACEHOLDER_TITLES = {
        "unknown", "n/a", "tbd", "to be determined", "candidate",
        "cabin crew candidate", "crew candidate", "pilot candidate",
        "flight deck candidate", "aviation candidate", "applicant",
        "seeking", "open to work",
    }
    if not job_title_str or job_title_str.lower() in _PLACEHOLDER_TITLES:
        return False

    if not rec.get("workTypes"):
        return False
    skills = rec.get("quickSkills") or []
    if not skills:
        return False

    work_type_ids = {w.get("id") for w in (rec.get("workTypes") or [])}
    FLIGHT_DECK_ID = 472

    # Flight deck must not contain cabin crew classification skills
    _CC_CLASSIFICATION = {"main crew", "main crew ", "senior cabin crew", "purser", "vip"}
    if FLIGHT_DECK_ID in work_type_ids:
        if any((s.get("name") or "").strip().lower() in _CC_CLASSIFICATION for s in skills):
            return False  # cabin crew label on a flight deck profile — needs fixing

    for skill in skills:
        name = (skill.get("name") or "").strip().lower()
        # Note: NATIONALITY_ADJECTIVES check removed — Tracker now stores nationalities
        # as demonyms ("British", "Indian" etc.) so these are the correct form, not a bug.
        if name in REGIONAL_TERMS:
            return False  # regional term instead of specific country — needs fixing

    # Country / nationality skill checks — include both country names AND nationality adjectives
    if country_skills_set:
        ENGINEERING_ID = 471
        _all_country_names = country_skills_set | KNOWN_COUNTRY_NAMES_EXTRA | {
            v.lower() for v in COUNTRY_ALIASES.values()
        }
        # Count actual country names (France, Martinique)
        country_hits = [s for s in skills
                        if (s.get("name") or "").strip().lower() in _all_country_names]
        # Also count nationality adjectives (French, Portuguese) — rule is ONE of each
        nationality_hits = [s for s in skills
                            if (s.get("name") or "").strip().lower() in NATIONALITY_ADJECTIVES]
        # Combined: if ANY category has >1, the profile is over-specified and needs fixing
        if len(nationality_hits) > 1:
            return False  # multiple nationality adjectives — must reduce to one
        is_fd  = FLIGHT_DECK_ID in work_type_ids
        is_sn  = not is_fd
        if is_sn and len(country_hits) < 1 and len(nationality_hits) < 1:
            return False  # cabin crew / other must have at least 1 nationality skill
        if is_sn and (len(country_hits) + len(nationality_hits)) > 1:
            return False  # cabin crew: exactly 1 nationality/country skill
        if is_fd and len(country_hits) > 2:
            return False  # flight deck: max nationality + licence country
        if is_fd and len(country_hits) < 2:
            return False  # flight deck must have both nationality AND licence country

    # Pilots must have aircraft type in job title (e.g. "A320 First Officer")
    _AIRCRAFT_PREFIX_RE = re.compile(
        r'\b([AB]\d{3}|CRJ|Dash|ATR|Q\d00|ERJ|E\d{3}|MD\d{2}|DC\d|F\d{2,3}|'
        r'PC-12|C\d{3}|DHC|L\d{3}|Do\d{2}|Fokker|Saab)\b', re.IGNORECASE)
    if FLIGHT_DECK_ID in work_type_ids and not _AIRCRAFT_PREFIX_RE.search(job_title_str):
        return False  # flight deck job title missing aircraft type prefix
    return True

# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print("\n" + "="*60)
    print("  AERO PROFESSIONAL — TRACKER PROFILE UPDATER  v2")
    print("="*60)

    # Authenticate
    print("\nAuthenticating to Tracker...")
    try:
        jwt = get_jwt()
        print("  ✓ Connected")
    except Exception as e:
        print(f"  ✗ Auth failed: {e}")
        send_run_summary_email(0, 0, 0, [("AUTH FAILURE", str(e))])
        return

    # ── Load from cache if available (avoids ~60s index rebuild) ─────────────────
    CACHE_FILE = "tracker_cache.json"
    CACHE_MAX_MINS = 1440   # rebuild every 24 hours

    cache_loaded = False
    licence_country_lookup = {}
    tracker_incomplete = []   # candidates with incomplete profiles (from Tracker scan)
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE) as f:
                cache = json.load(f)
            age_mins = (time.time() - cache.get("ts", 0)) / 60
            if age_mins < CACHE_MAX_MINS:
                skills_lookup = cache["skills"]
                country_skills_set   = set(cache.get("country_skills", []))
                nationality_ids      = set(cache.get("nationality_ids", []))
                licence_country_ids  = set(cache.get("licence_country_ids", []))
                licence_country_lookup = cache.get("licence_country_lookup", {})
                name_index    = cache["names"]
                tracker_incomplete = cache.get("incomplete_candidates", [])
                if not country_skills_set or not nationality_ids or not licence_country_ids:
                    print(f"\n  Cache missing country skill data — forcing rebuild...")
                else:
                    print(f"\n  ✓ Loaded from cache ({int(age_mins)}m old) — {len(name_index)} candidates, {len(skills_lookup)} skills")
                    print(f"    (Delete tracker_cache.json to force a full rebuild)")
                    cache_loaded = True
            else:
                print(f"\n  Cache is {int(age_mins)}m old — rebuilding...")
        except Exception:
            pass

    if not cache_loaded:
        # Load skills
        print("\nLoading skills reference data...")
        skills_lookup, country_skills_set, nationality_ids, licence_country_ids, licence_country_lookup = load_all_skills(jwt)
        print(f"  ✓ {len(skills_lookup)} skills loaded")

        # Build candidate index (full scan of all records)
        print("\nBuilding candidate index...")
        name_index, extra_skills, tracker_incomplete = build_candidate_index(jwt)
        skills_lookup.update(extra_skills)
        print(f"  ✓ Index ready. Total skills in lookup: {len(skills_lookup)}")

        # Save cache (preserve any existing OWA token)
        try:
            existing_cache = {}
            if os.path.exists(CACHE_FILE):
                try:
                    with open(CACHE_FILE) as f:
                        existing_cache = json.load(f)
                except Exception:
                    pass
            with open(CACHE_FILE, "w") as f:
                json.dump({
                    "ts":        time.time(),
                    "skills":    skills_lookup,
                    "country_skills":    list(country_skills_set),
                    "nationality_ids":    list(nationality_ids),
                    "licence_country_ids": list(licence_country_ids),
                    "licence_country_lookup": licence_country_lookup,
                    "names":     name_index,
                    "incomplete_candidates": tracker_incomplete,
                    "owa_token": existing_cache.get("owa_token"),
                    "owa_token_ts": existing_cache.get("owa_token_ts", 0),
                }, f)
            print(f"  ✓ Cache saved — next restart will be instant")
        except Exception as e:
            print(f"  ⚠  Could not save cache: {e}")

    # ── Read candidates from Outlook (or Tracker directly on GitHub Actions) ────
    ON_GITHUB_ACTIONS = os.environ.get("GITHUB_ACTIONS") == "true"
    candidates_to_process = []

    if not ON_GITHUB_ACTIONS:
        # Local: read from Outlook email folder as normal
        folder_label = EMAIL_SUBFOLDER if EMAIL_SOURCE == "subfolder" else "Support Inbox"
        print(f"\nReading candidates from Outlook ({folder_label})...")
        email_candidates = read_candidates_from_email()
        if email_candidates:
            print(f"  ✓ {len(email_candidates)} unique candidate(s) found")
            candidates_to_process = email_candidates
        else:
            print("  No candidates loaded — could not read emails from any source.")
            send_run_summary_email(0, 0, 0, [("EMAIL SOURCE FAILURE", "Could not read candidates from any email source.")])
            return
    else:
        # GitHub Actions: Microsoft 365 basic auth is blocked — read from Tracker directly.
        # The build_candidate_index scan already identified candidates with incomplete profiles.
        print(f"\nRunning on GitHub Actions — reading incomplete profiles from Tracker scan...")
        if not tracker_incomplete:
            print("  No incomplete profiles found in Tracker — nothing to do.")
            return "done"
        print(f"  ✓ {len(tracker_incomplete)} candidate(s) with incomplete profiles found")
        candidates_to_process = tracker_incomplete

    # ── Load processed-candidate tracking (prevents re-processing) ────────────
    PROCESSED_FILE = "tracker_processed.json"
    try:
        with open(PROCESSED_FILE) as f:
            _proc = json.load(f)
        processed_names    = set(_proc.get("names", []))
        processed_ids      = set(_proc.get("ids",   []))
        # resource_ids: dict of {name_lower → tracker_resource_id}
        # Used by audit_profiles.py to look up profiles directly without name search
        processed_resource_ids = dict(_proc.get("resource_ids", {}))
    except Exception:
        processed_names, processed_ids, processed_resource_ids = set(), set(), {}

    def _save_processed():
        try:
            with open(PROCESSED_FILE, "w") as f:
                json.dump({
                    "names":        list(processed_names),
                    "ids":          list(processed_ids),
                    "resource_ids": processed_resource_ids,   # name → tracker ID
                }, f)
        except Exception:
            pass

    # Filter out already-processed candidates (by name or email_id)
    unprocessed = []
    already_done_count = 0
    for c in candidates_to_process:
        nm  = " ".join(c["name"].strip().split()).lower()  # normalise internal whitespace
        eid = c.get("email_id") or ""
        if nm in processed_names or (eid and eid in processed_ids):
            already_done_count += 1
        else:
            unprocessed.append(c)

    candidates_to_process = unprocessed
    total = len(candidates_to_process)

    if already_done_count:
        print(f"  → {already_done_count} candidate(s) already processed on a previous run — skipping.")
    if total == 0:
        print("\n  All candidates already processed — nothing to do.")
        return "done"

    if TEST_MODE:
        candidates_to_process = candidates_to_process[:TEST_MODE]
        total = len(candidates_to_process)
        print(f"\n⚠  TEST MODE — processing first {total} candidate(s) only.")
        print("   Check Tracker after this run, then set TEST_MODE = 0 for the full run.\n")
    else:
        print(f"\nReady to process {total} candidate(s) — running fully automatically.")

    PENDING_MOVES_FILE = os.path.join(os.path.expanduser("~"), "tracker_pending_moves.txt")
    ERROR_LOG = "run_errors.txt"
    if os.path.dirname(ERROR_LOG):
        os.makedirs(os.path.dirname(ERROR_LOG), exist_ok=True)

    done = skipped = 0
    error_summary = []  # (name, reason) for every candidate that didn't complete
    _counter_lock = threading.Lock()

    # Populate shared JWT holder so worker threads always have the latest token
    _jwt_holder.clear()
    _jwt_holder.append(jwt)

    def _run_one(idx_cand):
        """Process one candidate — safe to call from a worker thread."""
        idx, cand = idx_cand
        name = cand["name"]
        # Initialise COM for this thread (safe to call multiple times)
        try:
            import pythoncom as _pc2
            _pc2.CoInitialize()
        except Exception:
            pass
        print(f"\n--- Candidate {idx} of {total} ---")
        try:
            result = process_one(
                name, _jwt_holder[0], name_index, skills_lookup,
                email_cand=cand,
                country_skills_set=country_skills_set,
                nationality_ids=nationality_ids,
                licence_country_ids=licence_country_ids,
                licence_country_lookup=licence_country_lookup,
            )
        except Exception as exc:
            import traceback as _tb
            tb = _tb.format_exc()
            with open(ERROR_LOG, "a", encoding="utf-8") as _ef:
                _ef.write(f"\n{'='*60}\n{name}\n{tb}\n")
            return name, False, cand, f"CRASHED: {exc}"
        return name, result, cand, None

    from concurrent.futures import ThreadPoolExecutor, as_completed as _as_completed
    WORKERS = 3   # process 3 candidates simultaneously

    with ThreadPoolExecutor(max_workers=WORKERS) as _pool:
        # Submit all candidates; refresh JWT periodically on main thread
        _futures = {}
        for _batch_i, _cand in enumerate(candidates_to_process, 1):
            # Refresh JWT every 50 submissions on the main thread
            if _batch_i % 50 == 1 and _batch_i > 1:
                try:
                    _jwt_holder[0] = get_jwt()
                except Exception:
                    pass
            _futures[_pool.submit(_run_one, (_batch_i, _cand))] = _cand

        for _fut in _as_completed(_futures):
            try:
                name, result, cand, err = _fut.result()
            except Exception as _fe:
                skipped += 1
                continue

            if err:
                print(f"\n  ✗ Unexpected error for '{name}': {err}")
                print(f"  → Logged to run_errors.txt — continuing with next candidate.")
                error_summary.append((name, err))
                skipped += 1
                continue

            if result is True:
                done += 1
                with _processed_lock:
                    _nm = " ".join(name.strip().split()).lower()
                    processed_names.add(_nm)
                    eid = cand.get("email_id") or ""
                    if eid:
                        processed_ids.add(eid)
                    rid = cand.get("tracker_id")
                    if rid:
                        processed_resource_ids[_nm] = rid
                    _save_processed()
            else:
                skipped += 1
                _reason = "Not found in Tracker / skipped"
                error_summary.append((name, _reason))
                _cand_email = cand if isinstance(cand, dict) else None
                if _cand_email and _cand_email.get("email_id"):
                    try:
                        move_to_done(_cand_email, dest_folder=EMAIL_NOT_FOUND_FOLDER)
                    except Exception:
                        pass

    print(f"\n{'='*60}")
    print(f"  Complete. Processed: {done} | Not found/failed: {skipped}")
    if already_done_count:
        print(f"  Already done from previous runs: {already_done_count}")
    if error_summary:
        print(f"\n  ⚠  The following {len(error_summary)} candidate(s) did NOT complete:")
        for cname, reason in error_summary:
            print(f"    • {cname}  —  {reason}")
        print(f"\n  Full error details saved to: run_errors.txt")
    if _NO_CV_NAMES:
        print(f"\n  ℹ  {len(_NO_CV_NAMES)} candidate(s) had no CV — skipped without emailing:")
        for _ncv in _NO_CV_NAMES:
            print(f"    • {_ncv}")
        _ncv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "no_cv_candidates.txt")
        try:
            with open(_ncv_path, "w", encoding="utf-8") as _f:
                _f.write(f"Candidates with no CV — run {datetime.datetime.now():%Y-%m-%d %H:%M}\n")
                _f.write("\n".join(_NO_CV_NAMES) + "\n")
            print(f"\n  → Full list saved to: no_cv_candidates.txt")
        except Exception:
            pass
    print(f"{'='*60}")
    if done > 0 or error_summary:
        send_run_summary_email(done, skipped, already_done_count, error_summary)


def check_pending_cvs():
    """
    Check all candidates in pending_cv.json.
    If they've uploaded a CV → remove from watchlist.
    If 5+ days with no CV → delete their Tracker profile.
    """
    print("\n" + "="*60)
    print("  PENDING CV CHECKER")
    print("="*60)

    try:
        with open(PENDING_CV_FILE, "r") as f:
            pending = json.load(f)
    except FileNotFoundError:
        print("\n  No pending_cv.json found — nothing to check.")
        return
    except json.JSONDecodeError:
        print("\n  Cannot read pending_cv.json.")
        return

    if not pending:
        print("\n  Watchlist is empty — nothing to check.")
        return

    jwt = get_jwt()
    today = datetime.date.today()
    still_pending = []
    removed = []
    deleted = []

    for entry in pending:
        rid       = entry.get("resourceId")
        name      = entry.get("name", "Unknown")
        added_str = entry.get("addedDate", "")

        try:
            added_date = datetime.date.fromisoformat(added_str)
        except Exception:
            added_date = today

        days_waiting = (today - added_date).days

        has_cv = False
        try:
            r = requests.get(f"{TRACKER_API}/api/v1/Resource/{rid}/Documents",
                             headers=h(jwt), timeout=10)
            if r.status_code == 200:
                _dj = r.json()
                if isinstance(_dj, list):
                    has_cv = len(_dj) > 0
                elif isinstance(_dj, dict):
                    has_cv = bool(_dj.get("items") or _dj.get("documents") or
                                  _dj.get("value") or _dj.get("results"))
        except Exception:
            has_cv = False

        if has_cv:
            print(f"  CV uploaded for {name} — removing from watchlist.")
            removed.append(name)
        elif days_waiting >= 5:
            print(f"  {name} — {days_waiting} days, no CV. Deleting profile.")
            try:
                requests.delete(f"{TRACKER_API}/api/v1/Resource/{rid}",
                                headers=h(jwt), timeout=10)
                deleted.append(name)
            except Exception as e:
                print(f"    Warning: could not delete {name}: {e}")
                still_pending.append(entry)
        else:
            print(f"  {name} — {days_waiting} day(s) waiting.")
            still_pending.append(entry)

    with open(PENDING_CV_FILE, "w") as f:
        json.dump(still_pending, f, indent=2)

    print(f"\n  Done. Removed: {len(removed)} | Deleted: {len(deleted)} | "
          f"Still waiting: {len(still_pending)}")
    print("=" * 60)


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "check_cvs":
        check_pending_cvs()
    else:
        # ── Lock file: prevent multiple instances running simultaneously ────────
        LOCK_FILE = "update_tracker.lock"
        if os.path.exists(LOCK_FILE):
            print("  ⚠  Another instance is already running (lock file exists). Exiting.")
            sys.exit(0)
        try:
            with open(LOCK_FILE, "w") as _lf:
                _lf.write(str(os.getpid()))
        except Exception:
            pass

        try:
            if DRY_RUN or TEST_MODE > 0 or os.environ.get("GITHUB_ACTIONS") == "true":
                # GitHub Actions: run once and exit — the scheduler handles re-running
                main()
            else:
                run_number = 0
                while True:
                    run_number += 1
                    result = main()
                    if result == "done":
                        print("\n  Nothing left. Sleeping 5 minutes...")
                        time.sleep(300)
                        run_number = 0
                    else:
                        time.sleep(60)
        finally:
            try:
                os.remove(LOCK_FILE)
            except Exception:
                pass
