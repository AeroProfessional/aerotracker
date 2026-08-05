"""
fix_reprocess.py — Correct the 4 profiles that reprocess_employers.py got wrong.

Employers confirmed from CVs:
  139982 Meram Mohamed      → Emcon PPS Steel Engineering
  140112 Carlo Rizza         → Heston Airlines
  139769 Giorgio Micoli      → AELIA
  137816 Ananthan Joshua     → LUFTAVIA LTD.
"""
import sys, os, requests
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from update_tracker import get_jwt, h, TRACKER_API

jwt = get_jwt()

profiles = [
    (139982, "Meram Mohamed",    "Emcon PPS Steel Engineering"),
    (140112, "Carlo Rizza",      "Heston Airlines"),
    (139769, "Giorgio Micoli",   "AELIA"),
    (137816, "Ananthan Joshua",  "LUFTAVIA LTD."),
]

print("=" * 65)
print("  FIX REPROCESS — CORRECT EMPLOYERS")
print("=" * 65)

passed = 0
for rid, name, employer in profiles:
    r = requests.patch(
        f"{TRACKER_API}/api/v1/Resource/{rid}",
        headers=h(jwt),
        json={"currentClient": {"id": -1, "name": employer}},
        timeout=15
    )
    ok = r.status_code in (200, 204)
    passed += ok
    print(f"  {'✓' if ok else '✗'} [{rid}] {name} → '{employer}'  (HTTP {r.status_code})")
    if not ok:
        print(f"      {r.text[:200]}")

print(f"\n  {passed}/{len(profiles)} patched successfully.")
print("=" * 65)
