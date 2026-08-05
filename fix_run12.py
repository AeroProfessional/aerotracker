"""
fix_run12.py — One-time fixes for run 12 profile issues.

  140127  Zohair Ahmed Siddiqui  — employer "Aviation Technology (UniKL-MIAT)" → restore "SMARTLYNX AIRLINES LIMITED"
                                    jobTitle "Unknown" → "Aircraft Systems & Avionics Engineer"
  139657  Mahmood Mohammed       — jobTitle "Unknown" → "B1 Certifying Aircraft Maintenance Engineer"

Run:  py fix_run12.py
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

print("=" * 60)
print("  FIX RUN 12 PROFILES")
print("=" * 60)

print("\n[1/2] Zohair Ahmed Siddiqui (140127)")
patch(140127, {
    "employer":  "SMARTLYNX AIRLINES LIMITED",
    "jobTitle":  "Aircraft Systems & Avionics Engineer",
}, "employer → 'SMARTLYNX AIRLINES LIMITED', jobTitle → 'Aircraft Systems & Avionics Engineer'")

print("\n[2/2] Mahmood Mohammed (139657)")
patch(139657, {
    "jobTitle": "B1 Certifying Aircraft Maintenance Engineer",
}, "jobTitle → 'B1 Certifying Aircraft Maintenance Engineer'")

print("\n" + "=" * 60)
print("  Done.")
print("=" * 60)
