"""Generate a spot-check HTML report from daily_updates.json and open it."""
import json, random, datetime, os, webbrowser

DAILY_LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "daily_updates.json")
SAMPLE_SIZE = 15

with open(DAILY_LOG_FILE) as f:
    log = json.load(f)

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

sample = random.sample(updates, min(SAMPLE_SIZE, len(updates)))

rows = ""
for i, c in enumerate(sample, 1):
    skills = ", ".join(c.get("skills") or []) or "(none)"
    rows += f"""
    <tr>
      <td>{i}</td>
      <td><strong>{c['name']}</strong><br><small>ID: {c['id']}</small></td>
      <td>{c.get('job_title','')}</td>
      <td>{c.get('employer','')}</td>
      <td>{c.get('work_type','')}</td>
      <td>{skills}</td>
    </tr>"""

html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>Spot-Check {check_date}</title>
<style>
  body {{ font-family: Arial, sans-serif; padding: 20px; }}
  h1 {{ color: #1a3c5e; }}
  table {{ border-collapse: collapse; width: 100%; }}
  th {{ background: #1a3c5e; color: white; padding: 10px; text-align: left; }}
  td {{ padding: 8px 10px; border-bottom: 1px solid #ddd; vertical-align: top; }}
  tr:hover {{ background: #f5f9ff; }}
  .meta {{ color: #666; margin-bottom: 20px; }}
</style></head><body>
<h1>AeroTracker Spot-Check</h1>
<p class="meta">Date: <strong>{check_date}</strong> &nbsp;|&nbsp;
Total updated: <strong>{len(updates)}</strong> &nbsp;|&nbsp;
Showing: <strong>{len(sample)}</strong> random profiles</p>
<table>
  <tr><th>#</th><th>Name</th><th>Job Title</th><th>Employer</th><th>Work Type</th><th>Skills</th></tr>
  {rows}
</table>
</body></html>"""

SHARED_FOLDER = r"C:\Users\EmilyWalton\Aeroprofessional Limited\Aeroprofessional - Documents\Aeropro\Admin\Admin Team"

# Save to shared team folder (OneDrive/SharePoint sync)
shared_out = os.path.join(SHARED_FOLDER, f"AeroTracker_SpotCheck_{check_date}.html")
try:
    os.makedirs(SHARED_FOLDER, exist_ok=True)
    with open(shared_out, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✓ Saved to shared folder: {shared_out}")
    out = shared_out
except Exception as e:
    # Fallback: save locally if shared folder not reachable
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"AeroTracker_SpotCheck_{check_date}.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"⚠ Shared folder not reachable ({e})")
    print(f"  Saved locally instead: {out}")

webbrowser.open(f"file:///{out}")
