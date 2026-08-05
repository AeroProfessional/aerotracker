"""
reprocess_batch.py — Re-run the update logic for profiles that were mis-updated.
                     Fetches CVs directly from Tracker by ID — no emails needed.

Run:  py reprocess_batch.py
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    import pythoncom
    pythoncom.CoInitialize()
except ImportError:
    pass

import requests
from update_tracker import get_jwt, load_all_skills, process_one, h, TRACKER_API

# ── Candidates to reprocess ────────────────────────────────────────────────────
# Format: (tracker_resource_id, exact_name_as_stored_in_tracker)
REPROCESS = [
    (74070,  "Mikko Vainio"),
    (106282, "kawinraj Yuganandam"),
    (139906, "TAIF HAMEED"),
    (139907, "Ibrahim Zubair"),
    (106766, "Josè Fiori"),
    (139953, "Dipanshu Khatri"),
    (89019,  "MOHAMMAD AGHDAIE"),
    (129932, "Rebekah Gaston"),
]

# ── Direct employer overrides ──────────────────────────────────────────────────
# Use when the CV text doesn't expose the employer in a parseable format.
EMPLOYER_OVERRIDES = {
    74070:  "Braathens Regional Airways AB",  # Mikko Vainio — CV not extractable
    139906: "Iraqi Airways",                  # TAIF HAMEED — CV has address, not employer name
    139953: "FSTC",                           # Dipanshu Khatri — OCR merged employer name
    89019:  "Iran Air",                       # Mohammad Aghdaie — corrupted by wrong match
}

# ── Direct skills overrides ────────────────────────────────────────────────────
# Use when the script can't produce the correct skill set automatically.
# Structure: {"nationality": [...], "fcl_country": [...], "other": [...]}
# nationality  → resolved via area-43 (nationality) lookup
# fcl_country  → resolved via area-39 (FCL country/licence country) lookup
# other        → resolved via general lookup (B1, B2, EASA, aircraft types, positions)
SKILLS_OVERRIDES = {
    # MIKKO VAINIO: Finnish = nationality, B1 + B2 only.
    74070: {
        "nationality": ["Finnish"],
        "other":       ["B1", "B2"],
    },
    # KAWINRAJ YUGANANDAM: Bahraini = nationality (ops/management = nationality only).
    106282: {
        "nationality": ["Bahraini"],
    },
    # TAIF HAMEED: Iraqi = nationality, Greece = FCL country (area-39 stays as country names), EASA, First Officer.
    139906: {
        "nationality":  ["Iraqi"],
        "fcl_country":  ["Greece"],
        "other":        ["EASA", "First Officer"],
    },
    # IBRAHIM ZUBAIR: Iraqi = nationality, Greece = FCL country, EASA, B737-NG, First Officer.
    139907: {
        "nationality":  ["Iraqi"],
        "fcl_country":  ["Greece"],
        "other":        ["EASA", "B737-NG", "First Officer"],
    },
    # DIPANSHU KHATRI: Indian = nationality, India = FCL country (area-39 country name), ICAO, Flight Instructor.
    139953: {
        "nationality":  ["Indian"],
        "fcl_country":  ["India"],
        "other":        ["ICAO", "Flight Instructor"],
    },
    # MOHAMMAD AGHDAIE: Iranian = nationality, Iran = FCL country (area-39 country name), ICAO, A320, A330, First Officer.
    89019: {
        "nationality":  ["Iranian"],
        "fcl_country":  ["Iran"],
        "other":        ["ICAO", "A320", "A330", "First Officer"],
    },
    # REBEKAH GASTON: British = nationality, VIP = current service level only.
    129932: {
        "nationality":  ["British"],
        "other":        ["VIP"],
    },
    # JOSÈ FIORI: Italian = nationality (Management — nationality only, process_one stripped it).
    106766: {
        "nationality":  ["Italian"],
    },
}

# ── Direct job title + work type overrides ────────────────────────────────────
# Use when the script sets the wrong job title or work type and it can't be
# trusted to re-parse correctly from the CV.
# work_type_id: Flight Deck=472, Cabin Crew=469, Management=470, Engineering=471, Operations=473
PROFILE_OVERRIDES = {
    139953: {
        "jobTitle":  "Chief Flight Instructor",
        "workTypes": [{"id": 472, "name": "Flight Deck"}],
    },
    89019: {
        "jobTitle":  "First Officer",
        "workTypes": [{"id": 472, "name": "Flight Deck"}],
    },
}

def apply_profile_override(jwt, resource_id, overrides):
    """Directly patch job title and/or work type on a Tracker profile."""
    r = requests.patch(
        f"{TRACKER_API}/api/v1/Resource/{resource_id}",
        json=overrides,
        headers=h(jwt), timeout=15
    )
    return r.status_code


def apply_employer_override(jwt, resource_id, employer):
    """Directly patch just the employer (currentClient) on a Tracker profile."""
    r = requests.patch(
        f"{TRACKER_API}/api/v1/Resource/{resource_id}",
        json={"currentClient": {"id": -1, "name": employer}},
        headers=h(jwt), timeout=15
    )
    return r.status_code

def apply_skills_override(jwt, resource_id, skill_spec, skills_lookup, licence_country_lookup):
    """
    Directly set the full skills list for a candidate.
    skill_spec: dict with keys "nationality", "fcl_country", "other" (all optional lists).
    - nationality  → resolved via main skills_lookup (area 43 IDs)
    - fcl_country  → resolved via licence_country_lookup (area 39 IDs, correct FCL bucket)
    - other        → resolved via main skills_lookup (B1/B2, EASA, positions, aircraft)
    Returns (status_code, resolved_skills_sent).
    """
    resolved = []
    buckets = [
        (skill_spec.get("nationality", []),  skills_lookup),
        (skill_spec.get("fcl_country",  []),  licence_country_lookup),
        (skill_spec.get("other",        []),  skills_lookup),
    ]
    for names, lookup in buckets:
        for name in names:
            match = lookup.get(name.strip().lower())
            if match:
                resolved.append({"id": match["id"], "name": match["name"]})
                print(f"    → '{name}': ID {match['id']}")
            else:
                resolved.append({"id": 0, "name": name})
                print(f"  ⚠  Skills override: '{name}' not in lookup — adding as free text (id=0)")
    r = requests.patch(
        f"{TRACKER_API}/api/v1/Resource/{resource_id}",
        json={"quickSkills": resolved},
        headers=h(jwt), timeout=15
    )
    return r.status_code, resolved


def verify_profile(jwt, resource_id):
    """Fetch the profile from Tracker and return the current quickSkills names."""
    r = requests.get(
        f"{TRACKER_API}/api/v1/Resource/{resource_id}",
        headers=h(jwt), timeout=15
    )
    if r.status_code == 200:
        qs = r.json().get("quickSkills") or []
        return [s.get("name") for s in qs if s.get("name")]
    return None

def main():
    print("="*60)
    print("  REPROCESS BATCH — re-updating mis-processed profiles")
    print("="*60)

    jwt = get_jwt()
    if not jwt:
        print("✗ Could not obtain JWT")
        return

    (skills_lookup, country_skills_set,
     nationality_ids, licence_country_ids,
     licence_country_lookup) = load_all_skills(jwt)

    # Build a minimal name_index so find_candidate can resolve these IDs
    # without needing to fetch all ~140k Tracker records.
    name_index = {}
    for rid, name in REPROCESS:
        name_index[name.strip().lower()] = rid

    total = len(REPROCESS)
    ok = 0
    failed = 0

    for i, (rid, name) in enumerate(REPROCESS, 1):
        print(f"\n--- Candidate {i} of {total} ---")
        try:
            result = process_one(
                name,
                jwt,
                name_index,
                skills_lookup,
                email_cand=None,          # no email to move
                country_skills_set=country_skills_set,
                nationality_ids=nationality_ids,
                licence_country_ids=licence_country_ids,
                licence_country_lookup=licence_country_lookup,
                force_reprocess=True,     # bypass is_profile_complete check
            )
            if result:
                ok += 1
            else:
                failed += 1
        except Exception as exc:
            import traceback
            print(f"  ✗ Exception: {exc}")
            traceback.print_exc()
            failed += 1

    # ── Phase 2: profile overrides (job title / work type) ─────────────────────
    if PROFILE_OVERRIDES:
        print(f"\n{'='*60}")
        print("  Applying profile overrides (job title / work type)...")
        for rid, overrides in PROFILE_OVERRIDES.items():
            try:
                status = apply_profile_override(jwt, rid, overrides)
                if 200 <= status < 300:
                    print(f"  ✓ ID {rid}: profile patched (HTTP {status}) — {overrides}")
                else:
                    print(f"  ✗ ID {rid}: profile patch FAILED (HTTP {status})")
            except Exception as oe:
                print(f"  ✗ ID {rid}: profile patch exception: {oe}")

    # ── Phase 3: employer overrides ─────────────────────────────────────────────
    if EMPLOYER_OVERRIDES:
        print(f"\n{'='*60}")
        print("  Applying employer overrides...")
        for rid, emp in EMPLOYER_OVERRIDES.items():
            try:
                status = apply_employer_override(jwt, rid, emp)
                if 200 <= status < 300:
                    print(f"  ✓ ID {rid}: employer set to '{emp}'")
                else:
                    print(f"  ✗ ID {rid}: employer override FAILED (HTTP {status})")
            except Exception as oe:
                print(f"  ✗ ID {rid}: employer override exception: {oe}")

    # ── Phase 4: skills overrides ────────────────────────────────────────────────
    if SKILLS_OVERRIDES:
        print(f"\n{'='*60}")
        print("  Applying skills overrides...")
        for rid, skill_spec in SKILLS_OVERRIDES.items():
            try:
                status, sent = apply_skills_override(jwt, rid, skill_spec, skills_lookup, licence_country_lookup)
                if 200 <= status < 300:
                    all_skills = (skill_spec.get("nationality", [])
                                  + skill_spec.get("fcl_country", [])
                                  + skill_spec.get("other", []))
                    print(f"  ✓ ID {rid}: PATCH accepted (HTTP {status}) — sent {all_skills}")
                    # Verify what Tracker actually saved
                    import time as _time
                    _time.sleep(0.5)  # brief pause to let Tracker commit
                    saved = verify_profile(jwt, rid)
                    if saved is not None:
                        print(f"  ↩ Tracker now shows: {saved}")
                        saved_lower = [x.strip().lower() for x in saved]
                        missing = [s for s in all_skills if s.strip().lower() not in saved_lower]
                        if missing:
                            print(f"  ⚠  MISMATCH — these skills were NOT saved by Tracker: {missing}")
                        else:
                            print(f"  ✓ All skills confirmed in Tracker")
                    else:
                        print(f"  ↩ Could not verify (profile fetch failed)")
                else:
                    print(f"  ✗ ID {rid}: skills override FAILED (HTTP {status})")
            except Exception as oe:
                print(f"  ✗ ID {rid}: skills override exception: {oe}")

    print(f"\n{'='*60}")
    print(f"  Done. Updated: {ok} | Failed: {failed}")
    print("="*60)

if __name__ == "__main__":
    main()
