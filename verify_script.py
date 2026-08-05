"""
verify_script.py — Run before update_tracker.py to confirm all critical fixes are intact.
If any check fails, it automatically restores from update_tracker.backup.py.

Run: py verify_script.py
"""
import os, sys

SCRIPT   = os.path.join(os.path.dirname(__file__), "update_tracker.py")
BACKUP   = os.path.join(os.path.dirname(__file__), "update_tracker.backup.py")

CHECKS = [
    # (description, string that MUST be present)
    ("Education filter in employer extraction",
     "EDUCATION_WORDS = re.compile"),

    ("ICAO/authority inserted into non_country not skills_objs",
     "non_country.insert(0, _auth_obj)"),

    ("Aircraft type inserted into non_country not skills_objs",
     "non_country.append(_ac_obj)"),

    ("Main Crew inserted into non_country not skills_objs",
     "non_country.insert(0, _mc_obj)"),

    ("Dedicated licence_country_lookup for area-39 IDs",
     "licence_country_lookup = {}  # area-39 only"),

    ("licence_country_lookup passed to process_one",
     "licence_country_lookup=licence_country_lookup"),

    ("FCL default uses licence_country_lookup not skills_lookup",
     "_fcl_obj  = licence_country_lookup.get(_nat_name.lower())"),

    ("OCR quality gate aviation keyword bypass",
     "_AVIATION_KWS = re.compile"),

    ("OCR threshold lowered to 0.45",
     "_ratio < 0.45"),

    ("Single __main__ block (no duplicate)",
     None),  # special: checked below

    ("Correct cache key is 'ts' not 'timestamp'",
     '"ts"] = __import__("time").time()'),

    ("Self-heal null bytes on startup",
     "_self_heal_nulls()"),

    ("Tesseract path auto-detection",
     "_set_tesseract_path()"),

    ("Not-found emails moved to Not Found folder",
     "move_to_done(_cand_email, dest_folder=EMAIL_NOT_FOUND_FOLDER)"),

    ("File does not end prematurely (tail intact)",
     "os.remove(LOCK_FILE)"),

    # ── New checks added 2026-07-20 ───────────────────────────────────────────
    ("Post-parse bad-title filter in process_one",
     "_FINAL_BAD_TITLE_RE"),

    ("Foreign-language header regex in is_valid_title",
     "FOREIGN_HEADER_RE"),

    ("Word count limit in is_valid_title (>7 words rejected)",
     "len(t.split()) > 7"),

    ("Invalid skill name filter before final dedup",
     "_INVALID_SKILL_NAME_RE"),

    ("None-word skill filter",
     "_NONE_WORD_RE"),

    ("Mojibake fix for rec_employer",
     'encode("latin-1").decode("utf-8")'),
]

def check():
    if not os.path.exists(SCRIPT):
        print(f"ERROR: {SCRIPT} not found!")
        return False

    with open(SCRIPT, "r", encoding="utf-8", errors="replace") as f:
        src = f.read()

    # Check for null bytes
    with open(SCRIPT, "rb") as f:
        raw = f.read()
    if b"\x00" in raw:
        print("FAIL: Null bytes found in script — file is corrupted")
        return False

    failures = []

    for desc, needle in CHECKS:
        if needle is None:
            # Special check: only one __main__ block
            count = src.count('if __name__ == "__main__"')
            if count != 1:
                failures.append(f"FAIL: '{desc}' — found {count} __main__ blocks (expected 1)")
            continue
        if needle not in src:
            failures.append(f"FAIL: '{desc}'")

    if failures:
        print("\n".join(failures))
        return False

    print(f"All {len(CHECKS)} checks passed — script is intact.")
    return True

if __name__ == "__main__":
    ok = check()
    if not ok:
        print("\n⚠  Script has lost critical fixes.")
        if os.path.exists(BACKUP):
            print("   Restoring from backup...")
            with open(BACKUP, "r", encoding="utf-8") as f:
                good = f.read()
            with open(SCRIPT, "w", encoding="utf-8") as f:
                f.write(good)
            print("   ✓ Restored. Re-running checks...")
            ok2 = check()
            if ok2:
                print("   ✓ Script restored successfully. Safe to run.")
            else:
                print("   ✗ Backup also failed checks — contact support.")
                sys.exit(1)
        else:
            print("   No backup found. Cannot auto-restore.")
            sys.exit(1)
    sys.exit(0 if ok else 1)
