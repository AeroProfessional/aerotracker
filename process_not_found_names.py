"""
process_not_found_names.py — Re-process candidates from tracker_not_found.txt.

Reads every name logged as unmatched or mismatched, re-runs the full find +
update pipeline (no email needed — CVs are downloaded directly from Tracker).
Candidates that succeed are removed from future runs; those that still fail
stay in the log.

Run:  py process_not_found_names.py
"""
import sys, os, re, json, time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pythoncom
pythoncom.CoInitialize()

from update_tracker import (
    get_jwt,
    build_candidate_index,
    load_all_skills,
    process_one,
)

HERE           = os.path.dirname(os.path.abspath(__file__))
NOT_FOUND_LOG  = os.path.join(HERE, "tracker_not_found.txt")
RESULTS_LOG    = os.path.join(HERE, "name_retry_results.txt")


def load_names():
    """
    Return a deduplicated, ordered list of candidate names to retry.
    - Plain lines  → truly not found last time
    - MISMATCH lines → found a close match but it was rejected
    Both are worth retrying with the current (improved) matching logic.
    Arabic-script names are skipped (can't match automatically).
    """
    names = []
    seen  = set()
    mismatch_re = re.compile(
        r"MISMATCH: looking for '(.+?)', found '(.+?)' \(ID (\d+)\)"
    )

    with open(NOT_FOUND_LOG, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue

            m = mismatch_re.match(line)
            if m:
                name = m.group(1).strip()
            else:
                name = line.strip()

            # Skip Arabic-script names — find_candidate already rejects these
            arabic_chars = sum(1 for c in name if '؀' <= c <= 'ۿ')
            if arabic_chars > len(name) * 0.4:
                continue

            key = name.lower().strip()
            if key not in seen:
                seen.add(key)
                names.append(name)

    return names


def main():
    print("=" * 60)
    print("  Process Not-Found Names — no email required")
    print("=" * 60)

    if not os.path.exists(NOT_FOUND_LOG):
        print("  tracker_not_found.txt not found — nothing to do.")
        return

    names = load_names()
    if not names:
        print("  No names to retry.")
        return

    print(f"\n  {len(names)} unique name(s) to retry")

    # ── Auth ──────────────────────────────────────────────────────────────────
    print("\nAuthenticating with Tracker...")
    jwt = get_jwt()
    print("  ✓ JWT obtained")

    # ── Index ─────────────────────────────────────────────────────────────────
    print("\nBuilding candidate index (takes ~60 sec)...")
    name_index, extra_skills = build_candidate_index(jwt)
    print(f"  ✓ {len(name_index)} entries indexed")

    # ── Skills ────────────────────────────────────────────────────────────────
    print("\nLoading skills lookup...")
    (skills_lookup, country_skills_set,
     nationality_ids, licence_country_ids,
     licence_country_lookup) = load_all_skills(jwt)
    print(f"  ✓ {len(skills_lookup)} skills loaded")

    # ── Process ───────────────────────────────────────────────────────────────
    total        = len(names)
    succeeded    = []
    still_failed = []
    errors       = []

    for i, name in enumerate(names, 1):
        print(f"\n--- {i} of {total} ---")
        try:
            result = process_one(
                name, jwt, name_index, skills_lookup,
                email_cand=None,          # no email — CV fetched from Tracker directly
                country_skills_set=country_skills_set,
                nationality_ids=nationality_ids,
                licence_country_ids=licence_country_ids,
                licence_country_lookup=licence_country_lookup,
            )
        except Exception as exc:
            import traceback
            print(f"  ✗ Crashed: {exc}")
            traceback.print_exc()
            result = False
            errors.append((name, str(exc)))

        if result is True:
            succeeded.append(name)
        else:
            still_failed.append(name)

        # Refresh JWT every 20 candidates
        if i % 20 == 0:
            try:
                jwt = get_jwt()
            except Exception:
                pass

    # ── Results ───────────────────────────────────────────────────────────────
    print(f"\n{'=' * 60}")
    print(f"  Total:        {total}")
    print(f"  Succeeded:    {len(succeeded)}")
    print(f"  Still failed: {len(still_failed)}")
    if errors:
        print(f"  Crashed:      {len(errors)}")

    # Write results log
    with open(RESULTS_LOG, "w", encoding="utf-8") as f:
        f.write(f"Run: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Total: {total}  Succeeded: {len(succeeded)}  Failed: {len(still_failed)}\n\n")
        if succeeded:
            f.write("=== SUCCEEDED ===\n")
            for n in succeeded:
                f.write(f"  {n}\n")
            f.write("\n")
        if still_failed:
            f.write("=== STILL FAILED ===\n")
            for n in still_failed:
                f.write(f"  {n}\n")
            f.write("\n")
        if errors:
            f.write("=== ERRORS ===\n")
            for n, e in errors:
                f.write(f"  {n}: {e}\n")

    print(f"\n  Full results saved to: name_retry_results.txt")

    # Rewrite tracker_not_found.txt to keep only the ones that still fail
    if still_failed or errors:
        remaining = set(n.lower() for n in still_failed + [e[0] for e in errors])
        with open(NOT_FOUND_LOG, "w", encoding="utf-8") as f:
            for n in still_failed:
                f.write(f"{n}\n")
            for n, _ in errors:
                f.write(f"{n}\n")
        print(f"  tracker_not_found.txt updated — {len(remaining)} name(s) remain")
    else:
        # All processed — clear the log
        open(NOT_FOUND_LOG, "w").close()
        print("  tracker_not_found.txt cleared — all candidates processed!")

    print("=" * 60)


if __name__ == "__main__":
    main()
