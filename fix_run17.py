"""
fix_run17.py — Corrections for the fifth 50-candidate batch.
"""
import sys, os, requests
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from update_tracker import get_jwt, h, TRACKER_API

jwt = get_jwt()

def patch(resource_id, payload, label):
    r = requests.patch(
        f"{TRACKER_API}/api/v1/Resource/{resource_id}",
        headers=h(jwt), json=payload, timeout=15
    )
    ok = r.status_code in (200, 204)
    print(f"  {'✓' if ok else '✗'} [{resource_id}] {label}  (HTTP {r.status_code})")
    if not ok:
        print(f"      {r.text[:200]}")
    return ok

print("=" * 65)
print("  FIX RUN 17 PROFILES")
print("=" * 65)

# 1. MD INBESHAT HASSAN AARZOO — employer is a description, not a company
print("\n[1/4] MD INBESHAT HASSAN AARZOO (140204)")
patch(140204, {
    "currentClient": {"id": -1, "name": "Unknown"},
}, "employer → Unknown (was 'Total Aviation Work Experience: 13 years' — description, not a company)")

# 2. José Miguel Nogueira — employer is aircraft registration, not a company
print("\n[2/4] José Miguel Nogueira de Freitas Carvalho (45002)")
patch(45002, {
    "currentClient": {"id": -1, "name": "Avolon Aero"},
}, "employer → restored 'Avolon Aero' (was 'G450, MSN 4088, Reg.: B' — aircraft registration)")

# 3. Trpimir Brkic — location suffix on employer
print("\n[3/4] Trpimir Brkic (131224)")
patch(131224, {
    "currentClient": {"id": -1, "name": "Pilatus Aircraft Ltd"},
}, "employer → 'Pilatus Aircraft Ltd' (was 'Pilatus Aircraft Ltd — Qatar' — location stripped)")

# 4. Charoonpan Khaothiemsang — job title suffix on employer
print("\n[4/4] Charoonpan Khaothiemsang (139870)")
patch(139870, {
    "currentClient": {"id": -1, "name": "Etihad Airways"},
}, "employer → 'Etihad Airways' (was 'ETIHAD AIRWAYS – Captain A320' — job title suffix stripped)")

print("\n" + "=" * 65)
print("  Done.")
print("=" * 65)
