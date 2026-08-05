"""
correct_profiles.py — Directly corrects two profiles that were mis-updated in the test run.

Alisha Pardiwalla  (ID 139900): wrong employer, wrong skills (FAA/US instead of ICAO/India)
Aminu Abdullahi    (ID 139901): wrong work type (Head Office), wrong job title

Run:  py correct_profiles.py
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pythoncom
pythoncom.CoInitialize()

from update_tracker import get_jwt, load_all_skills, update_resource, h
import requests

TRACKER_API = "https://evoglapi.tracker-rms.com"

def patch(jwt, resource_id, payload):
    r = requests.patch(f"{TRACKER_API}/api/v1/Resource/{resource_id}",
                       json=payload, headers=h(jwt), timeout=15)
    return r.status_code, r.text

def main():
    jwt = get_jwt()
    if not jwt:
        print("✗ Could not obtain JWT")
        return

    skills_lookup, country_skills_set, nationality_ids, licence_country_ids, licence_country_lookup = load_all_skills(jwt)

    def skill(name):
        key = name.strip().lower()
        if key in skills_lookup:
            return skills_lookup[key]
        # Try startswith
        for k, v in skills_lookup.items():
            if k.startswith(key) or key.startswith(k):
                if len(k) >= 4 and len(key) >= 4:
                    return v
        print(f"  ⚠  Skill '{name}' not found in lookup — adding as free text")
        return {"id": 0, "name": name}

    # ── Alisha Pardiwalla (ID 139900) ─────────────────────────────────────────
    # Correct: A320 First Officer | Interglobe Aviation LTD (INDIGO) | Flight Deck
    # Skills:  India (nationality), India (FCL country), ICAO, A320, First Officer
    print("\n" + "="*60)
    print("  Alisha Pardiwalla (ID 139900)")
    print("="*60)

    nat_india  = next(({"id": v["id"], "name": v["name"]} for k, v in skills_lookup.items()
                        if k in ("india","indian") and v.get("id") in nationality_ids), None)
    lic_india  = next(({"id": v["id"], "name": v["name"]} for k, v in licence_country_lookup.items()
                        if "india" in k), None)
    icao_obj   = skill("ICAO")
    a320_obj   = skill("A320")
    fo_obj     = skill("First Officer")

    alisha_skills = [s for s in [nat_india, lic_india, icao_obj, a320_obj, fo_obj] if s]
    print(f"  Skills resolved: {[s['name'] for s in alisha_skills]}")

    status, resp = patch(jwt, 139900, {
        "jobTitle":      "A320 First Officer",
        "currentClient": {"id": -1, "name": "Interglobe Aviation LTD (INDIGO)"},
        "workTypes":     [{"id": 472, "name": "Flight Deck"}],
        "quickSkills":   alisha_skills,
    })
    print(f"  {'✓' if 200 <= status < 300 else '✗'} HTTP {status}")

    # ── Aminu Abdullahi (ID 139901) ───────────────────────────────────────────
    # Correct: Unknown | Federal Airport Authority of Nigeria | Cabin Crew
    # Skills:  Nigeria (nationality)
    print("\n" + "="*60)
    print("  Aminu Abdullahi (ID 139901)")
    print("="*60)

    nat_nigeria = next(({"id": v["id"], "name": v["name"]} for k, v in skills_lookup.items()
                         if "nigeria" in k and v.get("id") in nationality_ids), None)
    if not nat_nigeria:
        # Try existing skill from Tracker record
        nat_nigeria = next(({"id": v["id"], "name": v["name"]} for k, v in skills_lookup.items()
                             if "nigeria" in k), None)
    if not nat_nigeria:
        nat_nigeria = {"id": 0, "name": "Nigeria"}

    aminu_skills = [nat_nigeria]
    print(f"  Skills resolved: {[s['name'] for s in aminu_skills]}")

    status, resp = patch(jwt, 139901, {
        "jobTitle":      "Unknown",
        "currentClient": {"id": -1, "name": "Federal Airport Authority of Nigeria"},
        "workTypes":     [{"id": 469, "name": "Cabin Crew"}],
        "quickSkills":   aminu_skills,
    })
    print(f"  {'✓' if 200 <= status < 300 else '✗'} HTTP {status}")

    print("\n  Done.")

if __name__ == "__main__":
    main()
