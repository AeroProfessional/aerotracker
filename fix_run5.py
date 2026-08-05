"""
fix_run5.py — One-time fixes for run 5 profile issues.

Profiles to fix:
  140031  Abdalrahman Hassan  — add 'Flight instructor' skill
  131473  Raimon David Arce   — revert employer to 'Puregold Duty Free (Subic) Inc.'
  140033  Maged Mourad        — revert employer to 'Air Cairo'
  140028  Kenneth Kamau       — fix employer to 'HELINT AVIATION LIMITED'
  140029  Maya Asenova        — fix employer to 'BG Agro Trading Company Ltd'

Run:  py fix_run5.py
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import requests
from update_tracker import get_jwt, h, TRACKER_API

jwt = get_jwt()

def patch_profile(resource_id, payload, label):
    r = requests.patch(
        f"{TRACKER_API}/api/v1/Resource/{resource_id}",
        headers=h(jwt), json=payload, timeout=15
    )
    print(f"  {'✓' if r.status_code in (200,204) else '✗'} [{resource_id}] {label}  (HTTP {r.status_code})")
    return r.status_code in (200, 204)

def patch_employer(resource_id, employer_name, resource_label):
    """Set employer via currentClient search."""
    # Search for employer
    r = requests.get(
        f"{TRACKER_API}/api/v1/Client",
        headers=h(jwt),
        params={"search": employer_name, "pageSize": 5},
        timeout=15
    )
    client_id = None
    if r.status_code == 200:
        items = r.json() if isinstance(r.json(), list) else (r.json().get("items") or r.json().get("results") or [])
        for item in items:
            name = (item.get("name") or item.get("clientName") or "").strip()
            if name.lower() == employer_name.lower():
                client_id = item.get("id") or item.get("clientId")
                break
        if not client_id and items:
            client_id = items[0].get("id") or items[0].get("clientId")

    if client_id:
        payload = {"currentClient": {"id": client_id, "name": employer_name}}
        patch_profile(resource_id, payload, f"employer → {employer_name} (via client ID {client_id})")
    else:
        # Fall back to free-text employer field
        payload = {"employer": employer_name}
        r2 = requests.patch(
            f"{TRACKER_API}/api/v1/Resource/{resource_id}",
            headers=h(jwt), json=payload, timeout=15
        )
        print(f"  {'✓' if r2.status_code in (200,204) else '✗'} [{resource_id}] {resource_label} employer → '{employer_name}' (free text, HTTP {r2.status_code})")

def patch_skills(resource_id, skill_ids, resource_label):
    """PATCH quickSkills — appends skill IDs to existing skills."""
    # Get current skills first
    r = requests.get(f"{TRACKER_API}/api/v1/Resource/{resource_id}", headers=h(jwt), timeout=15)
    if r.status_code != 200:
        print(f"  ✗ Could not read {resource_id}")
        return
    rec = r.json()
    existing = [{"id": s["id"]} for s in (rec.get("quickSkills") or []) if s.get("id")]
    existing_ids = {s["id"] for s in existing}

    new_skills = existing + [{"id": sid} for sid in skill_ids if sid not in existing_ids]
    r2 = requests.patch(
        f"{TRACKER_API}/api/v1/Resource/{resource_id}",
        headers=h(jwt),
        json={"quickSkills": new_skills},
        timeout=15
    )
    print(f"  {'✓' if r2.status_code in (200,204) else '✗'} [{resource_id}] {resource_label} skills patched (HTTP {r2.status_code})")

print("=" * 60)
print("  FIX RUN 5 PROFILES")
print("=" * 60)

# 1. Abdalrahman Hassan (140031) — add Flight instructor skill (ID 1077)
print("\n[1/5] Abdalrahman Hassan (140031) — add Flight instructor skill")
patch_skills(140031, [1077], "Abdalrahman Hassan")

# 2. Raimon David Arce (131473) — revert employer
print("\n[2/5] Raimon David Arce (131473) — revert employer to Puregold Duty Free (Subic) Inc.")
patch_employer(131473, "Puregold Duty Free (Subic) Inc.", "Raimon David Arce")

# 3. Maged Mourad (140033) — revert employer to Air Cairo
print("\n[3/5] Maged Mourad (140033) — revert employer to Air Cairo")
patch_employer(140033, "Air Cairo", "Maged Mourad")

# 4. Kenneth Kamau (140028) — fix employer (strip trailing date)
print("\n[4/5] Kenneth Kamau (140028) — fix employer to HELINT AVIATION LIMITED")
patch_employer(140028, "HELINT AVIATION LIMITED", "Kenneth Kamau")

# 5. Maya Asenova (140029) — fix employer (strip 'Economy' prefix)
print("\n[5/5] Maya Asenova (140029) — fix employer to BG Agro Trading Company Ltd")
patch_employer(140029, "BG Agro Trading Company Ltd", "Maya Asenova")

print("\n" + "=" * 60)
print("  Done.")
print("=" * 60)
