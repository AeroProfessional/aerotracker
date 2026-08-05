"""
fix_run8.py — One-time fixes for run 8 profile issues.

  140051  Inder Mohan Singh Virdi  — job title "Specialised Skills" → Unknown
                                     employer "from Emirates Aviation" → clear
  140049  SEHAR KHAN               — employer "Oct 2021" → restore "Marhaba Gelato"
  134798  Buhari Rehiman           — employer "SITSCO WLL, Bahrain" → "SITSCO WLL"

Run:  py fix_run8.py
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
print("  FIX RUN 8 PROFILES")
print("=" * 60)

# 1. Inder Mohan Singh Virdi (140051)
print("\n[1/3] Inder Mohan Singh Virdi (140051)")
patch(140051, {"jobTitle": "Unknown", "employer": ""}, "jobTitle → 'Unknown', employer → clear")

# 2. SEHAR KHAN (140049) — restore employer to existing value
print("\n[2/3] SEHAR KHAN (140049)")
patch(140049, {"employer": "Marhaba Gelato"}, "employer → 'Marhaba Gelato'")

# 3. Buhari Rehiman (134798) — strip ", Bahrain" from employer
print("\n[3/3] Buhari Rehiman (134798)")
patch(134798, {"employer": "SITSCO WLL"}, "employer → 'SITSCO WLL'")

print("\n" + "=" * 60)
print("  Done.")
print("=" * 60)
