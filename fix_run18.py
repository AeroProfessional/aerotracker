"""
fix_run18.py — Corrections for the sixth 50-candidate batch.
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
print("  FIX RUN 18 PROFILES")
print("=" * 65)

# 1. Joerg Schoenemann — employer "Airline and ACMI Charter" is a description
print("\n[1/12] Joerg Schoenemann (139869)")
patch(139869, {
    "currentClient": {"id": -1, "name": "Unknown"},
}, "employer → Unknown (was 'Airline and ACMI Charter' — service description, not a company)")

# 2. Gabor Szilagyi — employer "Manager Sales Aviation Products" is a job title
print("\n[2/12] Gabor Szilagyi (14167)")
patch(14167, {
    "currentClient": {"id": -1, "name": "RUAG"},
}, "employer → RUAG (was 'Manager Sales Aviation Products' — job title, not a company)")

# 3. Brian Kielberg — employer "26 years of airline experience" is a description
print("\n[3/12] Brian Kielberg (139844)")
patch(139844, {
    "currentClient": {"id": -1, "name": "Unknown"},
}, "employer → Unknown (was '26 years of airline experience' — description, not a company)")

# 4. Anastasija Frolova — employer is job title prefix; work type changed to Engineering incorrectly
print("\n[4/12] Anastasija Frolova (138494)")
patch(138494, {
    "currentClient": {"id": -1, "name": "Unknown"},
    "workTypes": [{"id": 474, "name": "Head Office"}],
}, "employer → Unknown; work type → Head Office (was Engineering)")

# 5. PIOTR SACILOWSKI — employer "Excellent knowledge of the aviation" is bio text
print("\n[5/12] PIOTR SACILOWSKI (139841)")
patch(139841, {
    "currentClient": {"id": -1, "name": "AES-GSE"},
}, "employer → AES-GSE (was 'Excellent knowledge of the aviation' — bio text, not a company)")

# 6. Mohammad Bhuiya — employer "AVIATION FLIGHT TIME" is flight hours text
print("\n[6/12] Mohammad Bhuiya (59759)")
patch(59759, {
    "currentClient": {"id": -1, "name": "Genesis Consultancy Sp. z o.o."},
}, "employer → Genesis Consultancy Sp. z o.o. (was 'AVIATION FLIGHT TIME' — flight hours text)")

# 7. Bokang Mojela — job title and skills completely wrong (company name used as job title)
print("\n[7/12] Bokang Mojela (139836)")
patch(139836, {
    "jobTitle": "Aircraft Technician",
    "workTypes": [{"id": 471, "name": "Engineering"}],
    "skills": [{"id": -1, "name": "South African"}],
}, "job title → Aircraft Technician; work type → Engineering; skills → ['South African']")

# 8. John Till — nationality skill wiped by bug (Operations profiles losing nationality)
print("\n[8/12] John Till (28614)")
patch(28614, {
    "skills": [{"id": -1, "name": "British"}],
}, "skills → ['British'] (nationality was wiped due to code bug)")

# 9. Abdullah Gasem — same nationality wipeout bug
print("\n[9/12] Abdullah Gasem (139834)")
patch(139834, {
    "skills": [{"id": -1, "name": "Saudi Arabian"}],
}, "skills → ['Saudi Arabian'] (nationality was wiped due to code bug)")

# 10. Neha Yadav — employer "Aviation Technical Services &" is truncated
print("\n[10/12] Neha Yadav (132516)")
patch(132516, {
    "currentClient": {"id": -1, "name": "Unknown"},
}, "employer → Unknown (was 'Aviation Technical Services &' — truncated text)")

# 11. Krishnamurthy Sathiyanarayanan — employer "TN India Completion date:" is a CV form field
print("\n[11/12] Krishnamurthy Sathiyanarayanan (128440)")
patch(128440, {
    "currentClient": {"id": -1, "name": "Unknown"},
}, "employer → Unknown (was 'TN India Completion date:' — CV form field, not a company)")

# 12. Bruno Terrazas — 'Main Crew' skill wrong for Flight Deck profile
print("\n[12/12] Bruno Terrazas (139854)")
patch(139854, {
    "skills": [{"id": -1, "name": "ICAO"}],
}, "skills → ['ICAO'] only (removed 'Main Crew' — cabin crew skill on flight deck profile; NOTE: nationality unknown, check CV manually)")

print("\n" + "=" * 65)
print("  Done.")
print("=" * 65)
