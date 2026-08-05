"""
retry_not_found.py — Retry all candidates sitting in the Not Found folder.

Reads every email from "New regs/Not Found", re-attempts Tracker lookup and
profile update with the current (improved) matching logic. Emails that succeed
are moved to the Done folder; emails that still can't be matched stay in
Not Found.

Run:  py retry_not_found.py
"""
import sys, os, json, time

# ── Make update_tracker importable ────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pythoncom
pythoncom.CoInitialize()

# Import shared logic from the main script
from update_tracker import (
    get_jwt,
    build_candidate_index,
    load_all_skills,
    process_one,
    move_to_done,
    parse_email_for_name,
    _get_source_folder,       # finds NEW REGS TO ACTION — already knows the support mailbox
    EMAIL_DONE_FOLDER,
    EMAIL_NOT_FOUND_FOLDER,
)

CACHE_FILE   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tracker_cache.json")
CACHE_MAX_MINS = 1440   # 24 hours


# ── Folder reader ──────────────────────────────────────────────────────────────

def _find_not_found_folder():
    """Find the Not Found subfolder inside NEW REGS TO ACTION.
    Uses _get_source_folder() which already knows how to reach the support mailbox.
    """
    try:
        src = _get_source_folder()   # returns the NEW REGS TO ACTION folder
        # "Not Found" is a direct subfolder
        for i in range(1, src.Folders.Count + 1):
            try:
                f = src.Folders.Item(i)
                if f.Name.strip().lower() == "not found":
                    return f
            except Exception:
                continue
    except Exception as e:
        print(f"  ⚠  Could not reach source folder: {e}")

    # Fallback: walk all Stores looking for any folder named "Not Found"
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns      = outlook.GetNamespace("MAPI")

        def _walk(folder, depth=0):
            if depth > 8:
                return None
            try:
                if folder.Name.strip().lower() == "not found":
                    return folder
            except Exception:
                return None
            try:
                for i in range(1, folder.Folders.Count + 1):
                    r = _walk(folder.Folders.Item(i), depth + 1)
                    if r is not None:
                        return r
            except Exception:
                pass
            return None

        for i in range(1, ns.Stores.Count + 1):
            try:
                root = ns.Stores.Item(i).GetRootFolder()
                r = _walk(root)
                if r is not None:
                    return r
            except Exception:
                continue
    except Exception:
        pass

    return None



def read_not_found_emails():
    """Return list of candidate dicts from the Not Found folder."""
    folder = _find_not_found_folder()
    if folder is None:
        print("  ✗ Could not locate 'Not Found' folder in Outlook.")
        return []

    print(f"  ✓ Found folder: {folder.FullFolderPath if hasattr(folder, 'FullFolderPath') else folder.Name}")

    candidates = []
    seen       = set()
    items      = folder.Items

    for item in items:
        try:
            name = parse_email_for_name(item.Subject or "", item.Body or "")
            if not name:
                continue
            key = name.strip().lower()
            if key in seen:
                continue
            seen.add(key)
            try:
                eid = item.EntryID
                sid = item.Parent.StoreID
            except Exception:
                eid, sid = None, None
            candidates.append({
                "name":        name.strip(),
                "item":        item,
                "entry_id":    eid,
                "store_id":    sid,
                "email_id":    None,
                "graph_token": None,
            })
        except Exception:
            continue

    return candidates


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  Retry Not Found — re-attempting unmatched candidates")
    print("=" * 60)

    # 1. Auth
    print("\nAuthenticating with Tracker...")
    jwt = get_jwt()
    print("  ✓ JWT obtained")

    # 2. Always rebuild the name index so we catch anyone added since last run
    # (78 retries are worth the extra 60 seconds)
    print("\nRebuilding candidate index from Tracker (takes ~60 sec)...")
    name_index, extra_skills = build_candidate_index(jwt)
    try:
        with open(CACHE_FILE, "w") as f:
            json.dump({"ts": time.time(), "index": name_index,
                       "extra_skills": extra_skills}, f)
        print(f"  ✓ Index saved ({len(name_index)} entries)")
    except Exception:
        pass

    # 3. Skills
    print("\nLoading skills lookup...")
    (skills_lookup, country_skills_set,
     nationality_ids, licence_country_ids,
     licence_country_lookup) = load_all_skills(jwt)
    print(f"  ✓ {len(skills_lookup)} skills loaded")

    # 4. Read emails
    print(f"\nReading emails from Not Found folder...")
    candidates = read_not_found_emails()
    if not candidates:
        print("  Nothing to retry.")
        return

    total = len(candidates)
    print(f"  ✓ {total} candidate(s) to retry")

    # 5. Process
    done         = 0
    still_failed = 0
    errors       = []

    for i, cand in enumerate(candidates, 1):
        name = cand["name"]
        print(f"\n--- Retry {i} of {total} ---")
        print(f"  {name}")
        print("=" * 60)

        try:
            result = process_one(
                name, jwt, name_index, skills_lookup,
                email_cand=cand,
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
            done += 1
            # Move out of Not Found → Done
            try:
                move_to_done(cand, dest_folder=EMAIL_DONE_FOLDER)
                print(f"  ↪ Moved to Done folder")
            except Exception as e:
                print(f"  ⚠  Could not move email: {e}")
        else:
            still_failed += 1
            # Leave in Not Found — no action needed

        # Refresh JWT every 20 candidates
        if i % 20 == 0:
            try:
                jwt = get_jwt()
            except Exception:
                pass

    # 6. Summary
    print(f"\n{'=' * 60}")
    print(f"  Retried:      {total}")
    print(f"  Succeeded:    {done}")
    print(f"  Still failed: {still_failed}")
    if errors:
        print(f"\n  Errors:")
        for n, e in errors:
            print(f"    {n}: {e}")
    print("=" * 60)


if __name__ == "__main__":
    main()
