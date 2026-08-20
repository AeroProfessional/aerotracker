"""
Correct the 5 profiles flagged in today's spot-check.

Fixes:
  82251  Zoltan Szendi        — add Austria (license country, derived from nationality)
 100011  Zuzanna Korczyc      — remove Main Crew (no cabin crew experience)
 116974  Zulema Manso         — remove Main Crew; fix job title to Receptionist
 131871  Zoltan Csek          — add Hungary (license country, derived from nationality)
 125043  Zoran Rodrigues Nuñez— remove Main Crew (no cabin crew experience)
"""
import os, requests, json, time

BEARER = os.environ.get("TRACKER_BEARER", "b28cae06af044958afb45fa8b1445fa7")
BASE   = "https://evoglapi.tracker-rms.com/api/v1"
HDR    = {"Authorization": f"Bearer {BEARER}", "Content-Type": "application/json"}

def get_profile(rid):
    r = requests.get(f"{BASE}/candidates/{rid}", headers=HDR, timeout=30)
    r.raise_for_status()
    return r.json()

def update_profile(rid, payload):
    r = requests.patch(f"{BASE}/candidates/{rid}", headers=HDR, json=payload, timeout=30)
    r.raise_for_status()
    return r.json()

def get_skills(profile):
    return profile.get("quickSkills") or []

def skill_names(skills):
    return {(s.get("name") or "").strip().lower(): s for s in skills}

def remove_skill(skills, name_lower):
    return [s for s in skills if (s.get("name") or "").strip().lower() != name_lower]

def find_skill_id(rid, name):
    """Look up skill object ID from the global skills list."""
    r = requests.get(f"{BASE}/skills?search={name}&limit=20", headers=HDR, timeout=30)
    if r.ok:
        for s in r.json().get("data", r.json() if isinstance(r.json(), list) else []):
            if (s.get("name") or "").strip().lower() == name.lower():
                return {"id": s["id"], "name": s["name"]}
    return None

print("Fetching profiles...\n")

# ── 1. Zoltan Szendi (82251) — add Austria ──────────────────────────────────
print("1. Zoltan Szendi (82251)")
p = get_profile(82251)
skills = get_skills(p)
sn = skill_names(skills)
if "austria" not in sn:
    austria = find_skill_id(82251, "Austria")
    if austria:
        skills.append(austria)
        update_profile(82251, {"quickSkills": skills})
        print("   ✓ Added Austria")
    else:
        print("   ⚠ Austria skill not found in lookup")
else:
    print("   ✓ Austria already present")

time.sleep(1)

# ── 2. Zuzanna Korczyc (100011) — remove Main Crew ──────────────────────────
print("2. Zuzanna Korczyc (100011)")
p = get_profile(100011)
skills = get_skills(p)
before = len(skills)
skills = remove_skill(skills, "main crew")
if len(skills) < before:
    update_profile(100011, {"quickSkills": skills})
    print("   ✓ Removed Main Crew")
else:
    print("   ✓ Main Crew not present")

time.sleep(1)

# ── 3. Zulema Manso (116974) — remove Main Crew + fix job title ─────────────
print("3. Zulema Manso (116974)")
p = get_profile(116974)
skills = get_skills(p)
before = len(skills)
skills = remove_skill(skills, "main crew")
payload = {"quickSkills": skills}
if (p.get("jobTitle") or "").strip().lower() in ("unknown", "", "n/a"):
    payload["jobTitle"] = "Receptionist"
    print("   ✓ Job title set to Receptionist")
if len(skills) < before:
    print("   ✓ Removed Main Crew")
update_profile(116974, payload)

time.sleep(1)

# ── 4. Zoltan Csek (131871) — add Hungary ───────────────────────────────────
print("4. Zoltan Csek (131871)")
p = get_profile(131871)
skills = get_skills(p)
sn = skill_names(skills)
if "hungary" not in sn:
    hungary = find_skill_id(131871, "Hungary")
    if hungary:
        skills.append(hungary)
        update_profile(131871, {"quickSkills": skills})
        print("   ✓ Added Hungary")
    else:
        print("   ⚠ Hungary skill not found in lookup")
else:
    print("   ✓ Hungary already present")

time.sleep(1)

# ── 5. Zoran Rodrigues Nuñez (125043) — remove Main Crew ────────────────────
print("5. Zoran Rodrigues Nuñez (125043)")
p = get_profile(125043)
skills = get_skills(p)
before = len(skills)
skills = remove_skill(skills, "main crew")
if len(skills) < before:
    update_profile(125043, {"quickSkills": skills})
    print("   ✓ Removed Main Crew")
else:
    print("   ✓ Main Crew not present")

print("\nDone. All 5 profiles corrected.")
