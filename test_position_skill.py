"""
test_position_skill.py — Find quickSkills IDs for all position skills in area-37,
                          and scan other areas for 'Flight Instructor'.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import requests
from update_tracker import get_jwt, h, TRACKER_API

# Skill area IDs to scan for "instructor"
SCAN_AREAS = {
    37: "position",
    38: "issuing_authority",
    39: "fcl_country",
    40: "aircraft",
    43: "nationality",
    44: "tri_tre",
    45: "unknown_45",
    46: "eng_licence",
    47: "cabin_service",
    48: "cabin_seniority",
}

def main():
    jwt = get_jwt()

    print("[1] All area-37 (position) skills from MetaData:")
    r = requests.get(
        f"{TRACKER_API}/api/v1/MetaData/Skills/Areas/37/Skills",
        params={"pageSize": 200, "pageNumber": 1},
        headers=h(jwt), timeout=15
    )
    print(f"  HTTP {r.status_code}")
    if r.status_code == 200:
        data = r.json()
        skills = data.get("skills", data) if isinstance(data, dict) else data
        for s in skills:
            print(f"  skillId={s.get('skillId'):>6}  name={s.get('skillName')}")

    print()
    print("[2] Scanning all skill areas for 'instructor' or 'Flight'...")
    for area_id, area_name in SCAN_AREAS.items():
        r = requests.get(
            f"{TRACKER_API}/api/v1/MetaData/Skills/Areas/{area_id}/Skills",
            params={"pageSize": 200, "pageNumber": 1},
            headers=h(jwt), timeout=15
        )
        if r.status_code != 200:
            print(f"  Area {area_id} ({area_name}): HTTP {r.status_code}")
            continue
        data = r.json()
        skills = data.get("skills", data) if isinstance(data, dict) else data
        matches = [s for s in skills
                   if "instructor" in (s.get("skillName") or "").lower()
                   or (area_id == 37)]  # always show all area-37 skills
        if matches:
            for s in matches:
                print(f"  Area {area_id} ({area_name}): skillId={s.get('skillId'):>6}  name={s.get('skillName')}")
        else:
            print(f"  Area {area_id} ({area_name}): no 'instructor' match (total: {len(skills)})")

    print()
    print("[3] Searching existing Flight Deck profiles for 'Flight Instructor' in quickSkills...")
    # Fetch first page of Flight Deck candidates
    r = requests.get(
        f"{TRACKER_API}/api/v1/Resource",
        params={"workTypeId": 472, "pageSize": 50, "pageNumber": 1},
        headers=h(jwt), timeout=30
    )
    if r.status_code == 200:
        data = r.json()
        resources = (data.get("resources") or data.get("items") or
                     data.get("value") or data.get("$values") or
                     (data if isinstance(data, list) else []))
        print(f"  Found {len(resources)} Flight Deck candidates in page 1")
        found = False
        for res in resources:
            rid = res.get("id") or res.get("resourceId")
            qs = res.get("quickSkills") or []
            for s in qs:
                if "instructor" in (s.get("name") or "").lower():
                    print(f"  ** FOUND: resourceId={rid}  skillId={s.get('id'):>6}  name={s.get('name')}")
                    found = True
        if not found:
            print("  No 'instructor' skills found in first 50 Flight Deck profiles")
            print("  → 'Flight Instructor' may not exist in Tracker yet — needs adding via admin")
    else:
        print(f"  HTTP {r.status_code}: {r.text[:200]}")

if __name__ == "__main__":
    main()
