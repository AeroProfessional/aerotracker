"""
fix_run9.py — One-time fixes for run 9 employer extraction bugs.

  99934   PANKAJ KUMAR JHA     — employer "Doha" (city) → restore "Qatar Airways"
  140066  Omar Nasser Afify    — employer "aviation principles" → restore "ALA Aerospace & Logistics"
  140067  Turki Hakami         — employer "fefueling,air shows" (OCR garbage) → clear

Run:  py fix_run9.py
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
print("  FIX RUN 9 PROFILES")
print("=" * 60)

# 1. PANKAJ KUMAR JHA (99934) — restore employer to Qatar Airways
print("\n[1/3] PANKAJ KUMAR JHA (99934)")
patch(99934, {"employer": "Qatar Airways"}, "employer → 'Qatar Airways'")

# 2. Omar Nasser Afify (140066) — restore employer to ALA Aerospace & Logistics
print("\n[2/3] Omar Nasser Afify (140066)")
patch(140066, {"employer": "ALA Aerospace & Logistics"}, "employer → 'ALA Aerospace & Logistics'")

# 3. Turki Hakami (140067) — clear garbled employer
print("\n[3/3] Turki Hakami (140067)")
patch(140067, {"employer": ""}, "employer → clear (OCR garbage)")

print("\n" + "=" * 60)
print("  Done.")
print("=" * 60)
