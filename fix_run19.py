"""
fix_run19.py — Corrections for the seventh 50-candidate batch.
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
print("  FIX RUN 19 PROFILES")
print("=" * 65)

# ── Saudi Arabian nationality wipeout (6 profiles) ───────────────────────────
# All had existing 'Saudi Arabian' skill which was wiped when the code stripped
# 'Saudi Arabia' (licence country) and left no nationality behind.

print("\n[1/14] Huzefa Shoaib (134753)")
patch(134753, {
    "skills": [{"id": -1, "name": "Saudi Arabian"}],
}, "skills → ['Saudi Arabian'] (nationality wiped by lic-country strip bug)")

print("\n[2/14] Abdullah Alloqman (139828)")
patch(139828, {
    "skills": [{"id": -1, "name": "Saudi Arabian"}],
}, "skills → ['Saudi Arabian'] (nationality wiped by lic-country strip bug)")

print("\n[3/14] Raed Al Ghamdi (139814)")
patch(139814, {
    "skills": [{"id": -1, "name": "Saudi Arabian"}],
}, "skills → ['Saudi Arabian'] (nationality wiped by lic-country strip bug)")

print("\n[4/14] Faleh Alsulami (139827)")
patch(139827, {
    "skills": [{"id": -1, "name": "Saudi Arabian"}],
}, "skills → ['Saudi Arabian'] (nationality wiped by lic-country strip bug)")

print("\n[5/14] Aamir Rizwan (139825)")
patch(139825, {
    "skills": [{"id": -1, "name": "Saudi Arabian"}],
}, "skills → ['Saudi Arabian'] (nationality wiped by lic-country strip bug)")

print("\n[6/14] Madawi Alhakami (139789)")
patch(139789, {
    "skills": [{"id": -1, "name": "Saudi Arabian"}],
}, "skills → ['Saudi Arabian'] (nationality wiped by lic-country strip bug)")

# ── Ghala Alotaibi — name as job title, generic employer, wiped nationality ──

print("\n[7/14] Ghala Alotaibi (139820)")
patch(139820, {
    "jobTitle": "Unknown",
    "currentClient": {"id": -1, "name": "Unknown"},
    "skills": [{"id": -1, "name": "Saudi Arabian"}],
}, "job title → Unknown (was candidate name); employer → Unknown (was 'Logistics'); skills → ['Saudi Arabian']")

# ── Unclosed parenthesis in employer names ────────────────────────────────────

print("\n[8/14] ODORICO ZAMPROGNO (139830)")
patch(139830, {
    "currentClient": {"id": -1, "name": "Airbus Helicopters"},
}, "employer → Airbus Helicopters (was 'Airbus Helicopters (Helibras' — unclosed parenthesis)")

print("\n[9/14] Robert Miles Ringsell (139812)")
patch(139812, {
    "currentClient": {"id": -1, "name": "Express Airways GmbH"},
}, "employer → Express Airways GmbH (was 'Express Airways GmbH)' — stray trailing parenthesis)")

# ── Generic single-word employers ─────────────────────────────────────────────

print("\n[10/14] Souban Farooqui (139808)")
patch(139808, {
    "currentClient": {"id": -1, "name": "Unknown"},
}, "employer → Unknown (was 'Logistics' — single generic word, not a company name)")

print("\n[11/14] Shihas P B (139806)")
patch(139806, {
    "currentClient": {"id": -1, "name": "Unknown"},
}, "employer → Unknown (was 'CARGO WAREHOUSE SENIOR AGENT' — job title extracted as employer)")

# ── Aircraft suffix in employer / Reza Param ─────────────────────────────────

print("\n[12/14] Reza Param (134755)")
patch(134755, {
    "currentClient": {"id": -1, "name": "Cargojet Airways"},
}, "employer → Cargojet Airways (was 'Cargojet Airways ( B767/757 )' — aircraft type suffix)")

# ── F/O titles overwritten with generic 'Pilot' ──────────────────────────────
# Both had existing F/O titles that the code didn't recognise as aviation
# (because 'F/O' is not in _AVIATION_RE) and replaced with generic 'Pilot'.

print("\n[13/14] Jaafar Alsarraf (139791)")
patch(139791, {
    "jobTitle": "F/O B737-800",
}, "job title → F/O B737-800 (was changed to generic 'Pilot' from existing 'F/O B737-800/MAX')")

print("\n[14/14] SARMAD Beg (139788)")
patch(139788, {
    "jobTitle": "F/O B737-800",
}, "job title → F/O B737-800 (was changed to generic 'Pilot' from existing 'F/O B737-800')")

print("\n" + "=" * 65)
print("  Done.")
print("=" * 65)
