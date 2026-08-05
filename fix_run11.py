"""
fix_run11.py — One-time fixes for run 11 profile issues.

  140131  Harikesh Tripathi  — employer "Module 10– Aviation Legislation" → restore existing
  140141  Hayden Tunmer      — jobTitle "B737NG B737NG Captain" → "B737NG Captain"
                               employer "South African Express Airways - CRJ200" → "South African Express Airways"
                               remove "United States" from skills

Run:  py fix_run11.py
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

def get_resource(resource_id):
    r = requests.get(
        f"{TRACKER_API}/api/v1/Resource/{resource_id}",
        headers=h(jwt), timeout=15
    )
    if r.status_code == 200:
        return r.json()
    return {}

print("=" * 60)
print("  FIX RUN 11 PROFILES")
print("=" * 60)

# 1. Harikesh Tripathi (140131) — restore employer to existing pre-update value
print("\n[1/2] Harikesh Tripathi (140131)")
patch(140131, {"employer": "AIRCRAFT REDELIVERY PROJECT - GMR HYDRABAD"},
      "employer → 'AIRCRAFT REDELIVERY PROJECT - GMR HYDRABAD'")

# 2. Hayden Tunmer (140141) — fix job title, employer, remove United States from skills
print("\n[2/2] Hayden Tunmer (140141)")
rec = get_resource(140141)
current_skills = rec.get("quickSkills", [])
print(f"  Current skills: {[s.get('name') for s in current_skills]}")
fixed_skills = [s for s in current_skills
                if (s.get("name") or "").strip().lower() != "united states"]
print(f"  Fixed skills:   {[s.get('name') for s in fixed_skills]}")
patch(140141, {
    "jobTitle":   "B737NG Captain",
    "employer":   "South African Express Airways",
    "quickSkills": fixed_skills,
}, "jobTitle, employer, removed United States from skills")

print("\n" + "=" * 60)
print("  Done.")
print("=" * 60)
