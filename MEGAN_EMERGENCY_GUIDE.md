# AeroTracker — Megan's Emergency Guide

This document covers everything you might need to do while Emily is on leave.
The system runs automatically — you only need to act if something goes wrong.

---

## How you'll know something is wrong

The system emails **support@aeroprofessional.com** automatically if there is a problem.
Watch for emails with subjects starting with:

- **⚠ AeroTracker — Tracker API token has EXPIRED** → follow Section 1 below
- **⚠ Tracker Update — X candidate(s) need attention** → follow Section 2 below

If the system is working normally you'll receive a daily email at 3pm listing any
candidates who registered without uploading a CV. No action needed from you on those
unless you want to follow up with the candidate manually.

---

## Section 1 — Tracker bearer token has expired

You will receive an alert email telling you the token has expired or stopped working.
This means the system cannot connect to Tracker RMS until you renew it.

> **Note:** The token does not have a fixed expiry date — it will only stop working if someone accidentally clicks Generate on the API page in Tracker (which creates a new token and invalidates the old one). Do not click Generate unless you intend to replace it.

**Step 1 — Get a new token from Tracker**
1. Log into Tracker RMS as normal
2. Go to **Tools & Settings** (top right menu)
3. Find **API Access** or **REST API**
4. Click **Generate new token** (or similar — the exact wording may vary)
5. Copy the token — it will look something like `b28cae06af044958afb45fa8b1445fa7`

Full instructions from Tracker: **https://academy.tracker-rms.com/Home/Lesson/1273**

**Step 2 — Update the token in GitHub**
1. Go to **https://github.com** and log in with your account
2. Open the **aerotracker** repository (Emily will have shared this with you)
3. Click the **Settings** tab
4. Left sidebar → **Secrets and variables** → **Actions**
5. Find **TRACKER_BEARER** and click the pencil icon next to it
6. Paste the new token into the value box
7. Click **Update secret**

**Step 3 — Trigger a manual run to confirm it works**
1. Click the **Actions** tab
2. Left sidebar → **AeroTracker — Process Registrations**
3. Click **Run workflow** → **Run workflow**
4. Wait 2–3 minutes and check that it completes with a green tick ✓
5. You should receive a summary email at support@ shortly after

---

## Section 2 — Candidates need attention / run errors

You'll receive an email listing candidates that the system could not process.
These are candidates who registered but whose profile could not be updated automatically.

**What to do:**
1. Log into Tracker RMS
2. Search for the candidate by name
3. Check their profile — if their CV is uploaded, update their details manually
4. If their profile looks fine already, no action needed

The most common reason is the candidate's name in the registration email
didn't match their name in Tracker. These will need to be done manually.

---

## Section 3 — Check if the system is running

If you're unsure whether the system is working:

1. Go to **https://github.com** and log in
2. Open the **aerotracker** repository
3. Click the **Actions** tab
4. You'll see a list of recent runs — green tick = success, red cross = failed
5. Click any run to see its full log

The system runs automatically every hour. If the last run was more than 2 hours ago
and shows a red cross, something may be wrong — check the log for the error message.

---

## Section 4 — Trigger a manual run

If you want to run the system immediately (e.g. after renewing the token):

1. Actions tab → **AeroTracker — Process Registrations** (left sidebar)
2. Click **Run workflow** → **Run workflow**
3. A green tick means it ran successfully

---

## Section 5 — No-CV report candidates

Every day at 3pm you'll receive an email listing candidates who registered
but haven't uploaded a CV. The system does NOT contact them automatically.

Your action:
- Email each candidate listed to request their CV
- If no CV is received after 3 days, delete their Tracker profile manually
- Once a candidate uploads their CV, the system will process them automatically
  on the next hourly run and remove them from future reports

---

## Section 6 — Running scripts manually on your computer

This is only needed in rare cases where a specific candidate profile needs to be corrected manually.

**One-time setup (do this once):**
1. Install Python from **https://python.org/downloads** — tick **"Add Python to PATH"** during install
2. Install Tesseract OCR from **https://digi.bib.uni-mannheim.de/tesseract/** — download `tesseract-ocr-w64-setup-v5.3.0.20221214.exe` and run it with all defaults
3. Log into GitHub → open the **aerotracker** repository → green **Code** button → **Download ZIP**
4. Unzip and move the folder to `C:\AeroTracker` (rename it if it downloads as `aerotracker-main`)
5. Open Command Prompt (search "cmd" in Start menu)
6. Type `cd C:\AeroTracker` and press Enter
7. Type `pip install -r requirements.txt --break-system-packages` and press Enter — let it finish

**To run a manual fix:**
1. Open Command Prompt
2. Type `cd C:\AeroTracker` and press Enter
3. Type `py reprocess_batch.py` and press Enter
4. It will connect to Tracker and show the status of the protected profiles

---

## Section 7 — Adding a new team member to GitHub

If a new starter needs access to the repository:

1. Ask them to create a free GitHub account at **https://github.com**
2. Get their GitHub username
3. Go to the **aerotracker** repository → **Settings** → **Collaborators**
4. Click **Add people** → enter their username → set role to **Write** → click **Add**
5. They will receive an email invitation to accept

---

## Key links

- GitHub repository: **https://github.com/emily123-png/aerotracker**
- Tracker RMS: **https://evouk.tracker-rms.com**
- Support inbox: **support@aeroprofessional.com**

---

## Who to contact if you're stuck

If something isn't covered here and you can't reach Emily, use **Claude AI**:

1. Go to **https://claude.ai** in your browser (free to create an account)
2. Start a new conversation and say something like:
   *"I'm the emergency contact for an automated system called AeroTracker at Aero Professional.
   It processes candidate registrations into Tracker RMS. [describe the problem you're seeing]."*
3. Download `update_tracker.py` and `reprocess_batch.py` from the GitHub repository
   and attach them to the conversation — Claude will read them and help you diagnose the issue.

Claude has no memory of previous conversations, so the more detail you give it about
what's wrong (copy/paste any error messages), the better it can help.
