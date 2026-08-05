"""
fix_run7.py — One-time fixes for run 7 employer extraction bugs.

  60701   Léon Deivalassane Lamartine  — employer "Personal information Aviation Expérience" → restore "Sky Express"
  140056  Tony Malala                  — employer "including Emirates and Air France" → restore "Kenya Airways"
  140052  Rohan Raj                    — employer "Ltd. Baramati (M.H), Assistant Flight" → restore "Bihar Flying Institute"

Run:  py fix_run7.py
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

print("=" * 60)
print("  FIX RUN 7 PROFILES")
print("=" * 60)

# 1. Léon Deivalassane Lamartine (60701) — restore employer to Sky Express
print("\n[1/3] Léon Deivalassane Lamartine (60701)")
patch(60701, {"employer": "Sky Express"}, "employer → 'Sky Express'")

# 2. Tony Malala (140056) — restore employer to Kenya Airways
print("\n[2/3] Tony Malala (140056)")
patch(140056, {"employer": "Kenya Airways"}, "employer → 'Kenya Airways'")

# 3. Rohan Raj (140052) — restore employer to Bihar Flying Institute
print("\n[3/3] Rohan Raj (140052)")
patch(140052, {"employer": "Bihar Flying Institute"}, "employer → 'Bihar Flying Institute'")

print("\n" + "=" * 60)
print("  Done.")
print("=" * 60)
