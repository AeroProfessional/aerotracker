# AeroTracker — GitHub Setup Guide

Follow these steps once to get AeroTracker running automatically on GitHub.
Estimated time: 15–20 minutes.

---

## Step 1 — Create a GitHub account

1. Go to **https://github.com**
2. Click **Sign up** (top right)
3. Use your work email: emily.walton@aeroprofessional.com
4. Choose a username (e.g. `aeroprofessional-emily`)
5. Complete email verification

---

## Step 2 — Create a private repository

1. Once logged in, click the **+** icon (top right) → **New repository**
2. Fill in:
   - **Repository name**: `aerotracker`
   - **Visibility**: ✅ **Private** (important — keeps scripts and state files private)
   - Leave everything else unchecked
3. Click **Create repository**

---

## Step 3 — Upload the script files

You'll see a page with upload options. Click **uploading an existing file**.

Upload ALL files from your `C:\AeroTracker` folder:
- `update_tracker.py`
- `reprocess_batch.py`
- `no_cv_report.py`
- `requirements.txt`
- The `.github` folder (drag the whole folder — GitHub will pick up the workflow inside)

> **Note:** The `.github` folder may be hidden on Windows.  
> In File Explorer: View → Show → Hidden items, then drag it.

Click **Commit changes** at the bottom.

---

## Step 4 — Add secrets (credentials)

Secrets are encrypted values that the scripts read as environment variables.
GitHub never displays them after you save them.

1. Go to your repository → **Settings** tab
2. Left sidebar → **Secrets and variables** → **Actions**
3. Click **New repository secret** and add each of these:

| Secret name      | Value                                   |
|------------------|-----------------------------------------|
| `TRACKER_BEARER` | Your Tracker RMS bearer token           |
| `EMAIL_PASSWORD` | `PurpleAutumn96?`                       |

> **When the Tracker token expires:** Go back to this page, click the pencil icon next to `TRACKER_BEARER`, paste the new token, save. That's all — no code changes needed.

---

## Step 5 — Enable GitHub Actions

1. Go to your repository → **Actions** tab
2. If prompted "Workflows aren't running", click **I understand my workflows, go ahead and enable them**

The workflow will now run automatically every hour.

---

## Step 6 — Test it manually

1. Actions tab → click **AeroTracker — Process Registrations** (left sidebar)
2. Click **Run workflow** → **Run workflow**
3. Watch the run complete (takes 2–5 minutes)
4. Check that a summary email arrives at support@aeroprofessional.com

---

## Step 7 — Set TEST_MODE to 0

Before the full live run, you need to remove the test limit.

1. In the repository, click on `update_tracker.py`
2. Click the pencil icon (Edit)
3. Find the line: `TEST_MODE = 5`
4. Change it to: `TEST_MODE = 0`
5. Click **Commit changes**

The next scheduled run will process all candidates with no limit.

---

## Giving your colleague access

1. Repository → **Settings** → **Collaborators**
2. Click **Add people** → enter their GitHub username or email
3. Set role to **Write** (lets them edit scripts and secrets)

They will get an email invitation to accept.

---

## What runs automatically

| Time           | What happens                                              |
|----------------|-----------------------------------------------------------|
| Every hour     | Checks for new registration emails, processes any found   |
| Every day 3pm  | Sends no-CV report to support@aeroprofessional.com        |

---

## If something goes wrong

- **Token expired**: You'll receive an alert email at support@. Update the `TRACKER_BEARER` secret.
- **Run failed**: Go to Actions tab → click the failed run → read the error log.
- **IMAP blocked**: Your M365 admin may need to enable IMAP Basic Auth for your account. Contact your IT admin.
- **Candidates not processing**: Check `run_errors.txt` in the repository for details.

---

## Important notes

- The `.github/workflows/aerotracker.yml` file controls the schedule. Do not delete it.
- `pending_cv.json` and `tracker_processed.json` are updated automatically after each run and committed back to the repo. Do not delete them.
- `tracker_cache.json` is the Tracker candidate index (rebuilt every 24 hours). Safe to delete if you need to force a rebuild.
