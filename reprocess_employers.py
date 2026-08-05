"""
Re-parse CVs for candidates whose employer was incorrectly cleared.
Applies full plausibility checks — same logic as the main script.
"""
import sys, os, requests, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from update_tracker import (get_jwt, h, TRACKER_API, get_cv_text, parse_cv,
                             _is_plausible_employer)

jwt = get_jwt()

candidates = [
    (139982, "Meram Mohamed"),
    (140112, "Carlo Rizza"),
    (139769, "Giorgio Micoli"),
    (137816, "Ananthan Joshua"),
]

print("=" * 65)
print("  REPROCESSING EMPLOYERS FROM CVs")
print("=" * 65)

def patch_employer(resource_id, name, employer_name):
    r = requests.patch(f"{TRACKER_API}/api/v1/Resource/{resource_id}",
                       headers=h(jwt),
                       json={"currentClient": {"id": -1, "name": employer_name}},
                       timeout=15)
    ok = r.status_code in (200, 204)
    print(f"  {'✓' if ok else '✗'} Employer → '{employer_name}'  (HTTP {r.status_code})")

def clean_extracted(emp):
    """Strip trailing location suffixes and truncate at opening parenthesis."""
    if not emp:
        return emp
    # Strip anything after " – " or " - " (location qualifiers)
    emp = re.sub(r'\s+[–\-]\s+[A-Z][a-z].*$', '', emp).strip()
    # Strip trailing comma+location: "AELIA, Venice Lido Airport..." → "AELIA"
    emp = re.sub(r',\s+[A-Z][a-z].*$', '', emp).strip()
    # Strip unclosed parenthetical at end: "Company (ATO IT.ATO..." → "Company"
    emp = re.sub(r'\s*\([^)]*$', '', emp).strip()
    return emp.rstrip('.,;– ')

for resource_id, name in candidates:
    print(f"\n[{name} ({resource_id})]")

    cv_text = get_cv_text(jwt, resource_id)
    if not cv_text or len(cv_text) < 30:
        print(f"  No CV — setting 'Unknown'")
        patch_employer(resource_id, name, "Unknown")
        continue

    print(f"  CV: {len(cv_text)} characters")

    parsed = parse_cv(cv_text, name)
    employer = parsed.get("current_employer", "").strip()

    # Apply plausibility gate (bypassed when calling parse_cv directly)
    if employer and not _is_plausible_employer(employer):
        print(f"  ✗ Rejected implausible: '{employer}'")
        employer = ""

    # Clean trailing noise (locations, unclosed parentheses)
    if employer:
        employer = clean_extracted(employer)

    if employer:
        patch_employer(resource_id, name, employer)
    else:
        print(f"  CV did not yield a clear employer — setting 'Unknown'")
        patch_employer(resource_id, name, "Unknown")

print("\n" + "=" * 65)
print("  Done.")
print("=" * 65)
