"""
fix_skipped_profiles.py — Re-processes specific candidates that were
incorrectly auto-skipped or incorrectly updated in the last run.

Candidates:
  - Ramon Aymami Perez   (ID 137977) — auto-skipped: "Cabin Crew Candidate" title, no nationality
  - Claudia Garcia       (ID 120975) — auto-skipped: "Unknown" title, no nationality
  - Alberto Martín       (ID 128737) — auto-skipped: "Main Crew" in flight deck skills
  - Chris Du Plessis     (ID 134908) — updated but wrong: FAA→United States missed, wrong employer

Run:  py fix_skipped_profiles.py
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pythoncom
pythoncom.CoInitialize()

from update_tracker import (
    get_jwt,
    build_candidate_index,
    load_all_skills,
    process_one,
    find_candidate,
)

CANDIDATES = [
    "Ramon Aymami Perez",
    "Claudia Garcia",
    "Alberto Martín",
    "Chris Du Plessis",
]

def main():
    print("\n" + "="*60)
    print("  FIX SKIPPED PROFILES")
    print("="*60)

    jwt = get_jwt()
    if not jwt:
        print("✗ Could not obtain JWT — check credentials")
        return

    print("\nBuilding candidate index (this takes a few minutes)...")
    name_index = {}
    build_candidate_index(jwt, name_index)
    print(f"  Index built: {len(name_index)} entries")

    skills_lookup = load_all_skills(jwt)
    print(f"  Skills loaded: {len(skills_lookup)} entries\n")

    succeeded = 0
    for name in CANDIDATES:
        print("\n" + "="*60)
        print(f"  {name}")
        print("="*60)
        ok = process_one(name, jwt, name_index, skills_lookup, email_cand=None)
        if ok:
            succeeded += 1

    print("\n" + "="*60)
    print(f"  Done. {succeeded}/{len(CANDIDATES)} profiles fixed.")
    print("="*60)

if __name__ == "__main__":
    main()
