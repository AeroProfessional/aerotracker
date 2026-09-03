"""Generate a daily spot-check HTML report and save to shared team folder.
Auto-pulls latest data from GitHub before generating so the report is always current.
Includes:
  - Section 1: Sample of profiles updated today (spot-check)
  - Section 2: All candidates currently waiting for a CV (pending_cv.json)
"""
import json, random, datetime, os, webbrowser

SCRIPT_DIR      = os.path.dirname(os.path.abspath(__file__))
DAILY_LOG_FILE  = os.path.join(SCRIPT_DIR, "daily_updates.json")
PENDING_CV_FILE = os.path.join(SCRIPT_DIR, "pending_cv.json")
SAMPLE_SIZE     = 15
SHARED_FOLDER   = os.path.join(os.path.expanduser("~"), "Aeroprofessional Limited", "Aeroprofessional - Documents", "Aeropro", "Admin", "Admin Team", "Claude")

# ── Load daily updates ────────────────────────────────────────────────────────
try:
    with open(DAILY_LOG_FILE) as f:
        log = json.load(f)
except (FileNotFoundError, json.JSONDecodeError):
    log = {}

today = datetime.date.today().isoformat()
updates = log.get(today, [])
check_date = today
if not updates:
    for days_back in range(1, 8):
        d = (datetime.date.today() - datetime.timedelta(days=days_back)).isoformat()
        if log.get(d):
            updates = log[d]
            check_date = d
            break

# Show the most recently registered candidates (highest Tracker IDs) first
updates_sorted = sorted(updates, key=lambda c: int(c.get("id", 0)), reverse=True)
sample = updates_sorted[:SAMPLE_SIZE]

# ── Load pending CV list ──────────────────────────────────────────────────────
try:
    with open(PENDING_CV_FILE) as f:
        pending_cv = json.load(f)
except (FileNotFoundError, json.JSONDecodeError):
    pending_cv = []

# Sort by how long they've been waiting (oldest first)
pending_cv.sort(key=lambda x: x.get("addedDate", ""))

# ── Build HTML: Section 1 — updated profiles ──────────────────────────────────
update_rows = ""
for i, c in enumerate(sample, 1):
    skills = ", ".join(c.get("skills") or []) or "(none)"
    update_rows += f"""
    <tr>
      <td>{i}</td>
      <td><strong>{c['name']}</strong><br><small>ID: {c['id']}</small></td>
      <td>{c.get('job_title','')}</td>
      <td>{c.get('employer','')}</td>
      <td>{c.get('work_type','')}</td>
      <td>{skills}</td>
    </tr>"""

if not update_rows:
    update_rows = '<tr><td colspan="6" style="color:#999;padding:12px">No profiles updated yet today.</td></tr>'

# ── Build HTML: Section 2 — no CV candidates ─────────────────────────────────
no_cv_rows = ""
for i, c in enumerate(pending_cv, 1):
    added   = c.get("addedDate", "")
    email   = c.get("email", "") or "(no email)"
    rid     = c.get("resourceId", "")
    name    = c.get("name", "")
    # Flag if waiting more than 7 days
    try:
        days_waiting = (datetime.date.today() - datetime.date.fromisoformat(added)).days
    except Exception:
        days_waiting = 0
    flag = " ⚠️" if days_waiting >= 7 else ""
    row_style = ' style="background:#fff8f0;"' if days_waiting >= 7 else ""
    no_cv_rows += f"""
    <tr{row_style}>
      <td>{i}</td>
      <td><strong>{name}</strong>{flag}<br><small>ID: {rid}</small></td>
      <td>{email}</td>
      <td>{added}</td>
      <td>{days_waiting} days</td>
    </tr>"""

if not no_cv_rows:
    no_cv_rows = '<tr><td colspan="5" style="color:#999;padding:12px">No candidates currently waiting for a CV.</td></tr>'

# ── Assemble full HTML ─────────────────────────────────────────────────────────
html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>AeroTracker Daily Report {check_date}</title>
<style>
  body {{ font-family: Arial, sans-serif; padding: 24px; max-width: 1100px; }}
  h1 {{ color: #1a3c5e; margin-bottom: 4px; }}
  h2 {{ color: #1a3c5e; margin-top: 40px; margin-bottom: 8px; font-size: 1.1em; }}
  table {{ border-collapse: collapse; width: 100%; margin-bottom: 32px; }}
  th {{ background: #1a3c5e; color: white; padding: 10px; text-align: left; }}
  td {{ padding: 8px 10px; border-bottom: 1px solid #ddd; vertical-align: top; font-size: 0.92em; }}
  tr:hover {{ background: #f5f9ff; }}
  .meta {{ color: #666; margin-bottom: 24px; font-size: 0.9em; }}
  .warn {{ color: #c0392b; font-weight: bold; }}
  .badge {{ display:inline-block; background:#1a3c5e; color:white; border-radius:4px; padding:2px 8px; font-size:0.8em; margin-left:6px; }}
</style></head><body>
<h1>AeroTracker Daily Report</h1>
<p class="meta">Date: <strong>{check_date}</strong></p>

<h2>✅ Section 1 — Profiles updated today <span class="badge">{len(updates)} total · showing {len(sample)}</span></h2>
<p style="color:#555;font-size:0.9em">Random sample. If anything looks wrong, note the candidate ID and flag to your manager.</p>
<table>
  <tr><th>#</th><th>Name</th><th>Job Title</th><th>Employer</th><th>Work Type</th><th>Skills</th></tr>
  {update_rows}
</table>

<h2>📋 Section 2 — Candidates waiting for a CV <span class="badge">{len(pending_cv)} total</span></h2>
<p style="color:#555;font-size:0.9em">These candidates registered but haven't uploaded a CV. AeroTracker cannot update their profile until a CV is available.
⚠️ = waiting 7+ days.</p>
<table>
  <tr><th>#</th><th>Name</th><th>Email</th><th>Added</th><th>Waiting</th></tr>
  {no_cv_rows}
</table>

</body></html>"""

# ── Save ──────────────────────────────────────────────────────────────────────
filename = f"AeroTracker_DailyReport_{check_date}.html"
shared_out = os.path.join(SHARED_FOLDER, filename)
try:
    os.makedirs(SHARED_FOLDER, exist_ok=True)
    with open(shared_out, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✓ Saved to shared folder: {shared_out}")
    out = shared_out
except Exception as e:
    out = os.path.join(SCRIPT_DIR, filename)
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"⚠ Shared folder not reachable ({e})")
    print(f"  Saved locally: {out}")

webbrowser.open(f"file:///{out}")
