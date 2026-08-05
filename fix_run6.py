"""
fix_run6.py — One-time fixes for run 6 profile issues.

  126160  Ricardo Molina     — restore Venezuelan + Venezuela, fix employer
  140017  Dominik Alke       — fix employer (Airline Transport Pilot → Unknown)
  140018  Ahmad Alhamad      — fix employer (❖ Air safety course → Unknown)
  140028  Kenneth Kamau      — was already fixed by fix_run5.py; included as check

Run:  py fix_run6.py
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

def set_employer(resource_id, employer_name, label):
    patch(resource_id, {"employer": employer_name}, f"{label} → employer '{employer_name}'")

def add_skills(resource_id, skill_names_and_ids, label):
    """Add skills by ID, keeping existing ones."""
    r = requests.get(f"{TRACKER_API}/api/v1/Resource/{resource_id}", headers=h(jwt), timeout=15)
    if r.status_code != 200:
        print(f"  ✗ Could not read {resource_id}")
        return
    existing = [{"id": s["id"]} for s in (r.json().get("quickSkills") or []) if s.get("id")]
    existing_ids = {s["id"] for s in existing}
    new_skills = existing + [{"id": sid} for name, sid in skill_names_and_ids if sid not in existing_ids]
    r2 = requests.patch(
        f"{TRACKER_API}/api/v1/Resource/{resource_id}",
        headers=h(jwt), json={"quickSkills": new_skills}, timeout=15
    )
    added = [name for name, sid in skill_names_and_ids if sid not in existing_ids]
    print(f"  {'✓' if r2.status_code in (200, 204) else '✗'} [{resource_id}] {label} → added skills {added}  (HTTP {r2.status_code})")

def replace_skills(resource_id, skill_ids_to_remove, skill_ids_to_add_names, label):
    """Remove specific skills by ID and add new ones."""
    r = requests.get(f"{TRACKER_API}/api/v1/Resource/{resource_id}", headers=h(jwt), timeout=15)
    if r.status_code != 200:
        print(f"  ✗ Could not read {resource_id}")
        return
    existing = [{"id": s["id"]} for s in (r.json().get("quickSkills") or []) if s.get("id")]
    kept = [s for s in existing if s["id"] not in skill_ids_to_remove]
    kept_ids = {s["id"] for s in kept}
    for name, sid in skill_ids_to_add_names:
        if sid not in kept_ids:
            kept.append({"id": sid})
    r2 = requests.patch(
        f"{TRACKER_API}/api/v1/Resource/{resource_id}",
        headers=h(jwt), json={"quickSkills": kept}, timeout=15
    )
    print(f"  {'✓' if r2.status_code in (200, 204) else '✗'} [{resource_id}] {label}  (HTTP {r2.status_code})")

print("=" * 60)
print("  FIX RUN 6 PROFILES")
print("=" * 60)

# 1. Ricardo Molina (126160)
# Employer was overwritten with "ICAO Airline TrasportPilot Licence" — revert to existing
# Nationality Venezuelan + Venezuela FCL were cleared — restore
# Also ICAO was added; original had EASA. Remove ICAO (id varies), restore Venezuelan + Venezuela
print("\n[1/3] Ricardo Molina (126160)")
# First read current state
r = requests.get(f"{TRACKER_API}/api/v1/Resource/126160", headers=h(jwt), timeout=15)
if r.status_code == 200:
    rec = r.json()
    print(f"  Current skills: {[s.get('name') for s in rec.get('quickSkills', [])]}")
    print(f"  Current employer: {(rec.get('currentClient') or {}).get('name') or rec.get('employer', '')}")

# Restore employer
set_employer(126160, "Línea Aérea Conviasa", "Ricardo Molina")

# Restore Venezuelan (nationality ID 876 area-43) + Venezuela (FCL ID 287 area-39)
# Remove any added ICAO (ID 185) if it wasn't there before
# Original skills: Venezuelan(876), EMB190(use free-text/existing), Venezuela(287), EASA(183), First Officer(1079)
replace_skills(
    126160,
    skill_ids_to_remove={185},          # remove ICAO (added wrongly)
    skill_ids_to_add_names=[            # restore the two country skills
        ("Venezuelan", 876),
        ("Venezuela",  287),
    ],
    label="Ricardo Molina — restore Venezuelan+Venezuela, remove ICAO"
)

# 2. Dominik Alke (140017) — employer "Airline Transport Pilot" → Unknown
print("\n[2/3] Dominik Alke (140017)")
set_employer(140017, "Unknown", "Dominik Alke")

# 3. Ahmad Alhamad (140018) — employer "❖ Air safety course" → Unknown
print("\n[3/3] Ahmad Alhamad (140018)")
set_employer(140018, "Unknown", "Ahmad Alhamad")

print("\n" + "=" * 60)
print("  Done.")
print("=" * 60)
