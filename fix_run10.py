"""
fix_run10.py — One-time fixes for run 10 profile issues.

  131142  Mark Zammit     — skills had ICAO + European; correct to EASA only, nationality Maltese
  140139  YOEDHI HENDRA   — employer "have had many aviation courses" → clear
  140134  Ali Lsallum     — jobTitle "Licensing" (CV section header) → Unknown
  61719   Awsam Farjo     — employer "aerospace supply sector" → restore "AVIATOR SOLUTIONS LIMITED"

Run:  py fix_run10.py
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import requests
from update_tracker import get_jwt, h, TRACKER_API

jwt = get_jwt()

def patch(resource_id, payload, label):
    r = requests.patch(
        f"{TRACKER_API}/api/v1/Resource/{resource_id}",
        headers=h(jwt), json=payload, timeout=15
    )
    print(f"  {'✓' if r.status_code in (200, 204) else '✗'} [{resource_id}] {label}  (HTTP {r.status_code})")
    return r.status_code in (200, 204)

def get_skills(resource_id):
    r = requests.get(
        f"{TRACKER_API}/api/v1/Resource/{resource_id}",
        headers=h(jwt), timeout=15
    )
    if r.status_code == 200:
        return r.json().get("quickSkills", [])
    return []

print("=" * 60)
print("  FIX RUN 10 PROFILES")
print("=" * 60)

# 1. Mark Zammit (131142) — remove ICAO, remove European, keep EASA + Maltese
print("\n[1/4] Mark Zammit (131142)")
current_skills = get_skills(131142)
print(f"  Current skills: {[s.get('name') for s in current_skills]}")
# Remove ICAO and European, keep everything else
fixed_skills = [s for s in current_skills
                if (s.get("name") or "").strip().lower() not in {"icao", "european"}]
# Check if Maltese is present; if not we can't add it without the skill ID
# but we can at least strip the wrong ones
has_maltese = any((s.get("name") or "").strip().lower() == "maltese" for s in fixed_skills)
print(f"  Fixed skills: {[s.get('name') for s in fixed_skills]}")
print(f"  Maltese present: {has_maltese}")
patch(131142, {"quickSkills": fixed_skills}, "removed ICAO + European")

# 2. YOEDHI HENDRA (140139) — clear garbled employer
print("\n[2/4] YOEDHI HENDRA (140139)")
patch(140139, {"employer": ""}, "employer → clear")

# 3. Ali Lsallum (140134) — fix job title
print("\n[3/4] Ali Lsallum (140134)")
patch(140134, {"jobTitle": "Unknown"}, "jobTitle → 'Unknown'")

# 4. Awsam Farjo (61719) — restore employer
print("\n[4/4] Awsam Farjo (61719)")
patch(61719, {"employer": "AVIATOR SOLUTIONS LIMITED"}, "employer → 'AVIATOR SOLUTIONS LIMITED'")

print("\n" + "=" * 60)
print("  Done.")
print("=" * 60)
