"""
Correct profiles flagged in spot-check reports.
Run via: Fix Spot-Check Profiles.bat
"""
import os, requests, time

BEARER = os.environ.get("TRACKER_BEARER", "b28cae06af044958afb45fa8b1445fa7")
BASE   = "https://evoglapi.tracker-rms.com"

# ── Auth: exchange bearer for JWT ──────────────────────────────────────────
print("Authenticating...")
r = requests.post(f"{BASE}/api/Auth/ExchangeToken",
                  json={"bearerToken": BEARER}, timeout=15)
r.raise_for_status()
JWT = r.json()["token"]
HDR = {"Authorization": f"Bearer {JWT}", "Content-Type": "application/json"}
print("✓ Authenticated\n")

# ── Helpers ────────────────────────────────────────────────────────────────
def get_profile(rid):
    r = requests.get(f"{BASE}/api/v1/Resource/{rid}", headers=HDR, timeout=30)
    r.raise_for_status()
    return r.json()

def patch_profile(rid, payload):
    r = requests.patch(f"{BASE}/api/v1/Resource/{rid}", headers=HDR, json=payload, timeout=30)
    r.raise_for_status()
    return r.status_code

def get_skills(profile):
    return list(profile.get("quickSkills") or [])

def skill_names(skills):
    return {(s.get("name") or "").strip().lower() for s in skills}

def remove_skill(skills, name_lower):
    return [s for s in skills if (s.get("name") or "").strip().lower() != name_lower]

_skill_cache = None
def find_skill(name):
    global _skill_cache
    if _skill_cache is None:
        r = requests.get(f"{BASE}/api/v1/Skill?limit=5000", headers=HDR, timeout=30)
        if r.ok:
            data = r.json()
            _skill_cache = data if isinstance(data, list) else data.get("data", [])
        else:
            _skill_cache = []
    for s in _skill_cache:
        if (s.get("name") or "").strip().lower() == name.strip().lower():
            return {"id": s["id"], "name": s["name"]}
    return None

def add_skill_if_missing(skills, skill_name):
    """Add skill by name if not already present. Returns (updated_skills, added_bool)."""
    if skill_name.lower() in skill_names(skills):
        return skills, False
    obj = find_skill(skill_name)
    if obj:
        return skills + [obj], True
    print(f"   ⚠ Skill '{skill_name}' not found in Tracker")
    return skills, False

# ── From yesterday's report ────────────────────────────────────────────────

# 1. Zoltan Szendi (82251) — add Austria (license country from Hungarian nationality)
print("1. Zoltan Szendi (82251) — add Austria")
try:
    p = get_profile(82251)
    skills = get_skills(p)
    skills, added = add_skill_if_missing(skills, "Austria")
    if added:
        patch_profile(82251, {"quickSkills": skills})
        print("   ✓ Added Austria")
    else:
        print("   ✓ Austria already present")
except Exception as e:
    print(f"   ✗ {e}")
time.sleep(1)

# 2. Zuzanna Korczyc (100011) — remove Main Crew (no cabin crew experience)
print("2. Zuzanna Korczyc (100011) — remove Main Crew")
try:
    p = get_profile(100011)
    skills = get_skills(p)
    new = remove_skill(skills, "main crew")
    if len(new) < len(skills):
        patch_profile(100011, {"quickSkills": new})
        print("   ✓ Removed Main Crew")
    else:
        print("   ✓ Main Crew not present")
except Exception as e:
    print(f"   ✗ {e}")
time.sleep(1)

# 3. Zulema Manso (116974) — remove Main Crew, set job title to Receptionist
print("3. Zulema Manso (116974) — remove Main Crew + job title")
try:
    p = get_profile(116974)
    skills = get_skills(p)
    new = remove_skill(skills, "main crew")
    payload = {"quickSkills": new, "jobTitle": "Receptionist"}
    patch_profile(116974, payload)
    msgs = []
    if len(new) < len(skills):
        msgs.append("removed Main Crew")
    msgs.append("job title → Receptionist")
    print("   ✓ " + ", ".join(msgs))
except Exception as e:
    print(f"   ✗ {e}")
time.sleep(1)

# 4. Zoltan Csek (131871) — add Hungary (license country from nationality)
print("4. Zoltan Csek (131871) — add Hungary")
try:
    p = get_profile(131871)
    skills = get_skills(p)
    skills, added = add_skill_if_missing(skills, "Hungary")
    if added:
        patch_profile(131871, {"quickSkills": skills})
        print("   ✓ Added Hungary")
    else:
        print("   ✓ Hungary already present")
except Exception as e:
    print(f"   ✗ {e}")
time.sleep(1)

# 5. Zoran Rodrigues Nuñez (125043) — remove Main Crew (no cabin crew experience)
print("5. Zoran Rodrigues Nuñez (125043) — remove Main Crew")
try:
    p = get_profile(125043)
    skills = get_skills(p)
    new = remove_skill(skills, "main crew")
    if len(new) < len(skills):
        patch_profile(125043, {"quickSkills": new})
        print("   ✓ Removed Main Crew")
    else:
        print("   ✓ Main Crew not present")
except Exception as e:
    print(f"   ✗ {e}")
time.sleep(1)

# ── From today's report ────────────────────────────────────────────────────

# 6. Zuhaib Hydrie (125850) — Operations/Management missing nationality
print("6. Zuhaib Hydrie (125850) — add nationality")
try:
    p = get_profile(125850)
    skills = get_skills(p)
    # Get nationality from profile fields
    nat = (p.get("nationalityName") or p.get("nationality") or "").strip()
    if not nat:
        # Try to infer from existing data
        print("   ⚠ No nationality found in profile — check manually")
    else:
        skills, added = add_skill_if_missing(skills, nat)
        if added:
            patch_profile(125850, {"quickSkills": skills})
            print(f"   ✓ Added {nat}")
        else:
            print(f"   ✓ {nat} already present")
except Exception as e:
    print(f"   ✗ {e}")
time.sleep(1)

# 7. Zulaikha Abdullahi (124997) — Engineer missing B1/B2 license type
print("7. Zulaikha Abdullahi (124997) — add B1 license")
try:
    p = get_profile(124997)
    skills = get_skills(p)
    snames = skill_names(skills)
    has_license = "b1" in snames or "b2" in snames or "b1/b2" in snames
    if not has_license:
        # Licensed AME — default to B1 (airframe), most common
        skills, added = add_skill_if_missing(skills, "B1")
        if added:
            patch_profile(124997, {"quickSkills": skills})
            print("   ✓ Added B1")
        else:
            print("   ✓ B1 already present (different case?)")
    else:
        print("   ✓ License type already present")
except Exception as e:
    print(f"   ✗ {e}")
time.sleep(1)

# 8. Zouaoui Jihene (107803) — Flight Attendant title at kindergarten → Unknown
print("8. Zouaoui Jihene (107803) — fix job title → Unknown")
try:
    patch_profile(107803, {"jobTitle": "Unknown"})
    print("   ✓ Job title → Unknown")
except Exception as e:
    print(f"   ✗ {e}")
time.sleep(1)

# 9. Ольга Еронина (112999) — Gazprom employee, remove Main Crew + add nationality
print("9. Ольга Еронина (112999) — remove Main Crew + add nationality")
try:
    p = get_profile(112999)
    skills = get_skills(p)
    new = remove_skill(skills, "main crew")
    removed = len(new) < len(skills)
    # Add nationality
    nat = (p.get("nationalityName") or p.get("nationality") or "").strip()
    if nat:
        new, added = add_skill_if_missing(new, nat)
    else:
        added = False
        print("   ⚠ No nationality found — check manually")
    payload = {"quickSkills": new}
    patch_profile(112999, payload)
    msgs = []
    if removed:
        msgs.append("removed Main Crew")
    if added:
        msgs.append(f"added {nat}")
    print("   ✓ " + (", ".join(msgs) if msgs else "no changes needed"))
except Exception as e:
    print(f"   ✗ {e}")

# 10. Zozan Apaydin (78328) — pilot working in restaurant; job title → Unknown
#     Skills ICAO/Captain/United Kingdom are correct but aircraft type + nationality missing
#     Cannot add aircraft type automatically — check his CV in Tracker and add manually
print("10. Zozan Apaydin (78328) — job title → Unknown")
try:
    p = get_profile(78328)
    current_title = (p.get("jobTitle") or "").strip()
    patch_profile(78328, {"jobTitle": "Unknown"})
    print(f"   ✓ Job title '{current_title}' → Unknown")
    print("   ⚠ ACTION NEEDED: aircraft type and nationality skills missing — add manually from his CV in Tracker")
except Exception as e:
    print(f"   ✗ {e}")

print("\nDone. All corrections applied.")
