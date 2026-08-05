"""
fix_run14.py — Corrections for the second 50-candidate test batch.

9 profiles received incorrect employer or job title data due to parser bugs
now fixed in update_tracker.py.

Root causes fixed in code:
  1. _AIRLINE_SIGNAL_RE regex broken — "Air Haifa" type names weren't protected
  2. _jt_stripped stripping existing Tracker titles (removing "- Subtitle" parts)
  3. "Proficient in..." not in leading-word rejection list
  4. "manager", "director" missing from EMPLOYER_SINGULAR_ROLE_END_RE
  5. Comma check not stripping short numeric suffixes (e.g. ", 9")
  6. "Aviation Expérience" / "Aviation Education" not in section-header rejection
  7. Parenthetical abbreviations like (LAME) not detected
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
        print(f"      Response: {r.text[:200]}")
    return ok

print("=" * 65)
print("  FIX RUN 14 PROFILES")
print("=" * 65)

# ── 1. Anish Thomas Wilson (140178) ───────────────────────────────────────────
# Job title was truncated: "Aircraft Technician I - Maintenance Operations"→"Aircraft Technician I"
# Employer was replaced: "Global Aerospace Logistics (GAL)"→"Proficient in Aviation Software & Manuals"
print("\n[1/9] Anish Thomas Wilson (140178)")
patch(140178, {
    "jobTitle": "Aircraft Technician I - Maintenance Operations",
    "currentClient": {"id": -1, "name": "Global Aerospace Logistics (GAL)"},
}, "job title restored, employer restored → 'Global Aerospace Logistics (GAL)'")

# ── 2. Shailesh Subhash Pol (140176) ──────────────────────────────────────────
# Employer: "Squadron Leader (retd) IndIan aIr Force" → should be "Indian Air Force"
print("\n[2/9] Shailesh Subhash Pol (140176)")
patch(140176, {
    "currentClient": {"id": -1, "name": "Indian Air Force"},
}, "employer corrected → 'Indian Air Force' (was garbled with military rank prefix)")

# ── 3. Alvi Mara (77945) ──────────────────────────────────────────────────────
# Employer: "Turkish Airlines INC, 9" → strip trailing ", 9"
print("\n[3/9] Alvi Mara (77945)")
patch(77945, {
    "currentClient": {"id": -1, "name": "Turkish Airlines"},
}, "employer corrected → 'Turkish Airlines' (trailing ', 9' stripped)")

# ── 4. Mohammed Althnayan (139680) ────────────────────────────────────────────
# Employer was set to "LOGISTICS MANAGER" (job title used as employer)
# Restore to previous value "SCOPA MILITARY INDUSTRIES"
print("\n[4/9] Mohammed Althnayan (139680)")
patch(139680, {
    "currentClient": {"id": -1, "name": "SCOPA MILITARY INDUSTRIES"},
}, "employer restored → 'SCOPA MILITARY INDUSTRIES' (was overwritten with job title text)")

# ── 5. Jose Luis Lopez Cutillas (40677) ───────────────────────────────────────
# Employer: "Aviation Education and JAA Licenses" → restore existing "Air Haifa"
# Caused by _AIRLINE_SIGNAL_RE not matching "Air Haifa" (air\s+\w broken regex)
print("\n[5/9] Jose Luis Lopez Cutillas (40677)")
patch(40677, {
    "currentClient": {"id": -1, "name": "Air Haifa"},
}, "employer restored → 'Air Haifa' (airline protection was broken by regex bug)")

# ── 6. Ananthan Joshua (137816) ───────────────────────────────────────────────
# Employer: "AIRLINES (LAME) BOEING 787-" — garbage line from CV
# Previous value "LTD" was also wrong. Clear to blank.
print("\n[6/9] Ananthan Joshua (137816)")
patch(137816, {
    "currentClient": {"id": -1, "name": ""},
}, "employer cleared (was 'AIRLINES (LAME) BOEING 787-' — OCR garbage line)")

# ── 7. Léon Deivalassane-Lamartine (60701) ─────────────────────────────────────
# Employer: "Aviation Expérience" (French section header) → restore "Sky Express"
print("\n[7/9] Léon Deivalassane-Lamartine (60701)")
patch(60701, {
    "currentClient": {"id": -1, "name": "Sky Express"},
}, "employer restored → 'Sky Express' (was 'Aviation Expérience' — French CV section header)")

# ── 8. Roby Muharomansyah (139855) ────────────────────────────────────────────
# Job title: "Bussiness Administration" (non-aviation, typo) → Unknown
print("\n[8/9] Roby Muharomansyah (139855)")
patch(139855, {
    "jobTitle": "Unknown",
}, "job title → Unknown (was 'Bussiness Administration' — non-aviation, not cleared)")

# ── 9. Muhammad Gunardi (140098) ──────────────────────────────────────────────
# Employer: "Boeing 737-200/400(Mandala Airlines" → extract "Mandala Airlines"
print("\n[9/9] Muhammad Gunardi (140098)")
patch(140098, {
    "currentClient": {"id": -1, "name": "Mandala Airlines"},
}, "employer corrected → 'Mandala Airlines' (was garbled aircraft type + company on one line)")

print("\n" + "=" * 65)
print("  Done.")
print("=" * 65)
