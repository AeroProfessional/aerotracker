"""
audit_profiles.py — Validate previously processed candidate profiles in Tracker.

Checks every profile for:
  - Employer: does it look like a real company name?
  - Job title: does it look like a real professional title?
  - Work type: is it set?
  - Flight deck: does the profile have nationality + authority (ICAO/EASA/FAA)?

Uses tracker_processed.json["resource_ids"] for direct lookups when available
(populated from the next run onwards).  For older entries it falls back to a
name search in Tracker.

Usage:
  py audit_profiles.py                    # check up to 100 profiles
  py audit_profiles.py --limit 500        # check up to 500
  py audit_profiles.py --limit 0          # check ALL (slow)
  py audit_profiles.py --fix              # also clear implausible employers
"""
import sys, os, json, re, time, argparse
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import requests
from update_tracker import (
    get_jwt, h, TRACKER_API,
    _is_plausible_employer, _is_plausible_job_title,
    _clean_employer_simple, find_candidate,
)

parser = argparse.ArgumentParser()
parser.add_argument("--limit", type=int, default=100,
                    help="Max profiles to check (0 = all, default 100)")
parser.add_argument("--fix", action="store_true",
                    help="Auto-clear employers that fail the plausibility check")
args = parser.parse_args()

PROCESSED_FILE = os.path.join(os.path.dirname(__file__), "tracker_processed.json")

with open(PROCESSED_FILE) as f:
    data = json.load(f)

names         = data.get("names", [])
resource_ids  = data.get("resource_ids", {})   # name_lower → tracker_id

print(f"tracker_processed.json: {len(names)} names, {len(resource_ids)} with Tracker IDs stored")

# Determine which candidates to check
# Prefer those with stored resource IDs (direct lookup, no extra API call)
have_ids    = [(n, resource_ids[n]) for n in names if n in resource_ids]
need_search = [n for n in names if n not in resource_ids]

limit = args.limit if args.limit > 0 else len(names)
to_check_direct = have_ids[:limit]
remaining = limit - len(to_check_direct)
to_check_search = need_search[:remaining] if remaining > 0 else []

total_checking = len(to_check_direct) + len(to_check_search)
print(f"Checking {total_checking} profiles "
      f"({len(to_check_direct)} direct, {len(to_check_search)} via name search)...\n")

jwt = get_jwt()

# ── Known-bad patterns ─────────────────────────────────────────────────────────
SECTION_HDR_RE = re.compile(
    r"^(training\s+courses?|core\s+skills?|key\s+skills?|technical\s+skills?|"
    r"speciali[sz]ed?\s+skills?|specialist\s+skills?|"
    r"further\s+education|additional\s+skills?|professional\s+development|"
    r"diploma\s+in|bachelor|qualification|tax\s+return|income\s+tax|"
    r"in\s+which|areas?\s+of\s+(?:strength|expertise|interest)|"
    r"accomplishment|honour|honor|licensing|module\s+\d)\b",
    re.IGNORECASE
)

def check_profile(resource_id, name_hint=""):
    """Fetch one Tracker profile and return a list of issue strings (empty = clean)."""
    r = requests.get(f"{TRACKER_API}/api/v1/Resource/{resource_id}",
                     headers=h(jwt), timeout=15)
    if r.status_code == 401:
        raise RuntimeError("Bearer token expired")
    if r.status_code != 200:
        return None, []   # can't fetch — skip

    rec    = r.json()
    first  = rec.get("firstName") or rec.get("firstname") or ""
    sur    = rec.get("surname") or ""
    full   = f"{first} {sur}".strip() or name_hint
    emp    = (rec.get("currentClient") or {}).get("name") or rec.get("employer") or ""
    title  = rec.get("jobTitle") or ""
    wt     = rec.get("workTypes") or []
    skills = [s.get("name", "") for s in rec.get("quickSkills", [])]

    issues = []

    # Employer
    if emp:
        if SECTION_HDR_RE.match(emp.strip()):
            issues.append(f"employer is a CV section header: '{emp}'")
        elif not _is_plausible_employer(emp):
            issues.append(f"employer looks wrong: '{emp}'")

    # Job title
    if title and title.lower() != "unknown":
        if SECTION_HDR_RE.match(title.strip()):
            issues.append(f"job title is a CV section header: '{title}'")
        elif not _is_plausible_job_title(title):
            issues.append(f"job title looks wrong: '{title}'")

    # Work type
    if not wt:
        issues.append("work type not set")

    # Flight deck completeness
    wt_names = [w.get("name", "").lower() for w in wt] if wt else []
    if any("flight" in w for w in wt_names):
        skills_lower = [s.lower() for s in skills]
        has_auth = any(s in skills_lower for s in ("icao", "easa", "faa"))
        if not has_auth:
            issues.append("flight deck — no authority skill (ICAO/EASA/FAA)")
        # At least one non-authority, non-position skill = nationality
        non_auth = [s for s in skills_lower
                    if s not in {"icao","easa","faa","captain","first officer",
                                 "second officer","co-pilot","copilot","tri","tre",
                                 "main crew","vip","business","senior / cabin manager"}]
        if not non_auth:
            issues.append("flight deck — no nationality skill")

    return full, issues


issues_list = []
ok_count = error_count = 0

# ── Direct lookups ─────────────────────────────────────────────────────────────
for i, (name_lower, rid) in enumerate(to_check_direct, 1):
    try:
        full, issues = check_profile(rid, name_lower)
        if full is None:
            error_count += 1
            continue
        if issues:
            issues_list.append({"id": rid, "name": full, "issues": issues})
            print(f"  ⚠  [{rid}] {full}")
            for iss in issues:
                print(f"        → {iss}")
        else:
            ok_count += 1
        if i % 25 == 0:
            print(f"  ... checked {i}/{total_checking}")
            time.sleep(0.3)
    except RuntimeError:
        print("  ✗ Bearer token expired — stopping.")
        sys.exit(1)
    except Exception as e:
        error_count += 1

# ── Name-search fallback ───────────────────────────────────────────────────────
if to_check_search:
    print(f"\n  (Searching Tracker by name for {len(to_check_search)} older entries...)")
    # Build name index once
    try:
        ni_r = requests.get(f"{TRACKER_API}/api/v1/Resource?pageSize=5000&fields=id,firstName,surname",
                            headers=h(jwt), timeout=30)
        if ni_r.status_code == 200:
            _items = ni_r.json()
            if isinstance(_items, dict):
                _items = (_items.get("items") or _items.get("value") or
                          _items.get("results") or [])
            name_index = {
                f"{(it.get('firstName') or it.get('firstname') or '').strip()} "
                f"{(it.get('surname') or '').strip()}".strip().lower(): it.get("id")
                for it in _items if it.get("id")
            }
        else:
            name_index = {}
    except Exception:
        name_index = {}

    for i, name_lower in enumerate(to_check_search, 1):
        try:
            rid = find_candidate(name_lower, name_index, jwt)
            if rid is None or isinstance(rid, list):
                error_count += 1
                continue
            full, issues = check_profile(rid, name_lower)
            if full is None:
                error_count += 1
                continue
            if issues:
                issues_list.append({"id": rid, "name": full, "issues": issues})
                print(f"  ⚠  [{rid}] {full}")
                for iss in issues:
                    print(f"        → {iss}")
            else:
                ok_count += 1
            if i % 25 == 0:
                print(f"  ... searched {i}/{len(to_check_search)}")
                time.sleep(0.3)
        except Exception:
            error_count += 1

# ── Auto-fix ───────────────────────────────────────────────────────────────────
if args.fix and issues_list:
    print(f"\n{'='*60}")
    print("  AUTO-FIX: clearing implausible employers")
    print(f"{'='*60}")
    fixed = 0
    for issue in issues_list:
        emp_issues = [x for x in issue["issues"] if "employer" in x]
        if emp_issues and issue.get("id"):
            r = requests.patch(
                f"{TRACKER_API}/api/v1/Resource/{issue['id']}",
                headers=h(jwt), json={"employer": ""}, timeout=15
            )
            status = "✓" if r.status_code in (200, 204) else "✗"
            print(f"  {status} [{issue['id']}] {issue['name']}  (HTTP {r.status_code})")
            fixed += 1
    print(f"\n  Cleared {fixed} employer(s).")

# ── Report ─────────────────────────────────────────────────────────────────────
print(f"\n{'='*60}")
print(f"  AUDIT COMPLETE — checked {total_checking} of {len(names)} processed profiles")
print(f"{'='*60}")
print(f"  Clean:        {ok_count}")
print(f"  Issues found: {len(issues_list)}")
print(f"  Errors/skip:  {error_count}")

if issues_list:
    report_path = os.path.join(os.path.dirname(__file__), "audit_report.json")
    with open(report_path, "w") as f:
        json.dump(issues_list, f, indent=2)
    print(f"\n  Report saved to audit_report.json")
    if not args.fix:
        print(f"  Run with --fix to auto-clear implausible employers.")
