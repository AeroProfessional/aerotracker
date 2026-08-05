"""
fix_run15.py — Corrections for the third 50-candidate test batch.

Root causes fixed in update_tracker.py:
  1. Numbered list prefix not stripped → "1. Adria Airways" now → "Adria Airways"
  2. MM/YYYY date range not stripped → "Avion Express 08/2017 – 04/2019" now fixed
  3. Em-dash location suffix not stripped → "Kenya Airways – Nairobi" now fixed
  4. "–Description" suffix → "Shree Airlines –From Technical Services..." now fixed
  5. "Lead/Head" ending not rejected → "Engineering Team Lead" now rejected
  6. "Dfferences Course" / "further more" now rejected as job titles
  7. Multi-match crash when best=None → now handled gracefully
  8. "leasing"/"distribution"/"optimisation" added to action-word rejections
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
print("  FIX RUN 15 PROFILES")
print("=" * 65)

# ── 1. MUNDHER AL HARTHI (140187) ─────────────────────────────────────────────
# Job title set to "Dfferences Course" (OCR grabbed a course header)
# Employer set to "Oman Air RVSM" ("RVSM" is a qualification suffix, not part of company name)
print("\n[1/10] MUNDHER AL HARTHI (140187)")
patch(140187, {
    "jobTitle": "Unknown",
    "currentClient": {"id": -1, "name": "Oman Air"},
}, "job title → Unknown, employer → 'Oman Air' (was 'Dfferences Course' / 'Oman Air RVSM')")

# ── 2. Sibongiseni Moyo (99280) ────────────────────────────────────────────────
# Job title set to "further more" (text fragment from CV)
print("\n[2/10] Sibongiseni Moyo (99280)")
patch(99280, {
    "jobTitle": "Unknown",
}, "job title → Unknown (was 'further more' — CV text fragment)")

# ── 3. Philly Owuor (140185) ──────────────────────────────────────────────────
# Employer set to "Kenya Airways – Nairobi" → strip location suffix
print("\n[3/10] Philly Owuor (140185)")
patch(140185, {
    "currentClient": {"id": -1, "name": "Kenya Airways"},
}, "employer → 'Kenya Airways' (was 'Kenya Airways – Nairobi' — location suffix stripped)")

# ── 4. Ciacoi Luca (140043) ───────────────────────────────────────────────────
# Employer set to "Avion Express 08/2017 – 04/2019" → strip date range
print("\n[4/10] Ciacoi Luca (140043)")
patch(140043, {
    "currentClient": {"id": -1, "name": "Avion Express"},
}, "employer → 'Avion Express' (was 'Avion Express 08/2017 – 04/2019' — date range stripped)")

# ── 5. Parash Pokhrel (140042) ────────────────────────────────────────────────
# Employer set to "Shree Airlines –From Technical Services In-charge" → strip description
print("\n[5/10] Parash Pokhrel (140042)")
patch(140042, {
    "currentClient": {"id": -1, "name": "Shree Airlines"},
}, "employer → 'Shree Airlines' (was '...–From Technical Services In-charge' — suffix stripped)")

# ── 6. Akhtar Rasool (139977) ─────────────────────────────────────────────────
# Employer set to "Engineering Team Lead" (job title used as employer)
# Restore existing Tracker value "TELTRIUM"
print("\n[6/10] Akhtar Rasool (139977)")
patch(139977, {
    "currentClient": {"id": -1, "name": "TELTRIUM"},
}, "employer → 'TELTRIUM' (was 'Engineering Team Lead' — job title mistaken for employer)")

# ── 7. Marjan Kos (69430) ─────────────────────────────────────────────────────
# Employer set to "1. Adria Airways" (numbered list prefix)
print("\n[7/10] Marjan Kos (69430)")
patch(69430, {
    "currentClient": {"id": -1, "name": "Adria Airways"},
}, "employer → 'Adria Airways' (was '1. Adria Airways' — list number prefix stripped)")

# ── 8. ALIKEMAL SEVINCH (140074) ─────────────────────────────────────────────
# Employer replaced from "Grandstrap Industrial Packaging Corporation" to
# "Logistics & Distribution Optimisation" (skills description, not a company)
# Restore original
print("\n[8/10] ALIKEMAL SEVINCH (140074)")
patch(140074, {
    "currentClient": {"id": -1, "name": "Grandstrap Industrial Packaging Corporation"},
}, "employer → restored 'Grandstrap Industrial Packaging Corporation' (was skills description)")

# ── 9. Piter Cumarin (140045) ────────────────────────────────────────────────
# Employer replaced from "Copa Airlines" to "Airlines Leasing Airplanes" (description)
# Restore Copa Airlines; also work type was changed to Engineering — revert to Operations
print("\n[9/10] Piter Cumarin (140045)")
patch(140045, {
    "currentClient": {"id": -1, "name": "Copa Airlines"},
    "workTypes": [{"id": 473, "name": "Operations"}],
}, "employer → restored 'Copa Airlines', work type → Operations (was 'Airlines Leasing Airplanes' + Engineering)")

# ── 10. Abdul Kareem Mohamed Ibrahim (139966) ────────────────────────────────
# Employer set to "Gulf Air (GFA) – Bahrain" → strip location suffix → "Gulf Air"
print("\n[10/10] Abdul Kareem Mohamed Ibrahim (139966)")
patch(139966, {
    "currentClient": {"id": -1, "name": "Gulf Air"},
}, "employer → 'Gulf Air' (was 'Gulf Air (GFA) – Bahrain' — location suffix stripped)")

print("\n" + "=" * 65)
print("  Done.")
print("=" * 65)
