"""
diagnose_notfound.py  —  Run on Emily's machine (not the sandbox)
Tests EVERY possible Tracker API search endpoint for known not-found candidates.
Output goes to diagnose_results.txt AND screen.

Run:  py diagnose_notfound.py
"""
import requests, json, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

TRACKER_API    = "https://evoglapi.tracker-rms.com"
TRACKER_BEARER = "b28cae06af044958afb45fa8b1445fa7"

# ── Auth ──────────────────────────────────────────────────────────────────────
def get_jwt():
    r = requests.post(f"{TRACKER_API}/api/v1/Authentication/Login",
                      json={"apiKey": TRACKER_BEARER}, timeout=15)
    if r.status_code == 200:
        return r.json().get("token")
    raise Exception(f"Auth failed: {r.status_code} — {r.text[:200]}")

jwt  = get_jwt()
hdrs = {"Authorization": f"Bearer {jwt}", "Content-Type": "application/json"}
print("Connected to Tracker\n")

# ── Candidates to test ────────────────────────────────────────────────────────
NOT_FOUND = [
    "Philip Lipesa",
    "ZEED Alotibi",
    "Yazdaan Memon",
    "Omar Abdalla",
    "Mazen Lamfon",
    "Fadi Bakkar",
    "Haziq Ummer",
    "Nishan Jayasinghe",
    "Alisha Rana",
    "Sadiq Salah",
]

# ── Helpers ───────────────────────────────────────────────────────────────────
def post(endpoint, body):
    try:
        r = requests.post(f"{TRACKER_API}{endpoint}", json=body,
                          headers=hdrs, timeout=15)
        return r.status_code, r.text
    except Exception as e:
        return 0, str(e)

def parse_items(raw_text, status):
    """Try to extract a list of candidate dicts from the raw response text."""
    if status not in (200, 201):
        return None, raw_text[:120]
    try:
        data = json.loads(raw_text)
    except Exception:
        return None, raw_text[:120]
    if isinstance(data, list):
        return data, None
    if isinstance(data, dict):
        for key in ("items", "results", "value", "data", "resources", "contacts"):
            if key in data and isinstance(data[key], list):
                return data[key], None
        # No list found — show all keys so we know the shape
        return None, f"dict keys: {list(data.keys())}"
    return None, str(data)[:120]

def fmt_item(x):
    fn  = x.get("firstname")    or x.get("firstName")    or ""
    sn  = x.get("surname")      or x.get("lastName")      or ""
    nm  = x.get("name")         or x.get("resourceName")  or x.get("fullName") or ""
    rid = x.get("resourceId")   or x.get("contactId")     or x.get("id")       or "?"
    return f"    ID {rid} | first='{fn}' | surname='{sn}' | name='{nm}'"

# ── First: show page 1 unfiltered so we can see the field names ───────────────
print("=== Unfiltered page 1 (to check field names) ===")
status, raw = post("/api/v1/Resource/Search", {"pageSize": 3, "pageNumber": 1})
items, err = parse_items(raw, status)
if items:
    for x in items:
        print(" Sample record keys:", list(x.keys()))
        print(" ", fmt_item(x))
        break
else:
    print(" Error:", err)
print()

# ── Per-candidate tests ───────────────────────────────────────────────────────
all_output = []

for full_name in NOT_FOUND:
    parts = full_name.strip().split()
    first = parts[0]
    last  = parts[-1]
    mid   = " ".join(parts[1:-1]) if len(parts) > 2 else ""

    block = [f"\n{'='*60}", f"  {full_name}", f"{'='*60}"]
    print(f"\n{'='*60}\n  {full_name}\n{'='*60}")

    searches = [
        # ── Resource/Search ───────────────────────────────────
        ("Resource/Search  keyword=full",
         "/api/v1/Resource/Search",
         {"pageSize": 20, "pageNumber": 1, "keyword": full_name}),

        ("Resource/Search  searchTerm=full",
         "/api/v1/Resource/Search",
         {"pageSize": 20, "pageNumber": 1, "searchTerm": full_name}),

        ("Resource/Search  firstName+surname",
         "/api/v1/Resource/Search",
         {"pageSize": 20, "pageNumber": 1, "firstName": first, "surname": last}),

        ("Resource/Search  surname-only",
         "/api/v1/Resource/Search",
         {"pageSize": 20, "pageNumber": 1, "surname": last}),

        ("Resource/Search  firstName-only",
         "/api/v1/Resource/Search",
         {"pageSize": 20, "pageNumber": 1, "firstName": first}),

        ("Resource/Search  keyword=surname",
         "/api/v1/Resource/Search",
         {"pageSize": 20, "pageNumber": 1, "keyword": last}),

        # ── Resource/PagedSearch ──────────────────────────────
        ("Resource/PagedSearch  keyword=full",
         "/api/v1/Resource/PagedSearch",
         {"pageSize": 20, "pageNumber": 1, "keyword": full_name}),

        ("Resource/PagedSearch  searchTerm=full",
         "/api/v1/Resource/PagedSearch",
         {"pageSize": 20, "pageNumber": 1, "searchTerm": full_name}),

        # ── Contact/Search ────────────────────────────────────
        ("Contact/Search  firstName+lastName",
         "/api/v1/Contact/Search",
         {"pageSize": 20, "pageNumber": 1, "firstName": first, "lastName": last}),

        ("Contact/Search  searchTerm=full",
         "/api/v1/Contact/Search",
         {"pageSize": 20, "pageNumber": 1, "searchTerm": full_name}),

        # ── Contact/PagedSearch ───────────────────────────────
        ("Contact/PagedSearch  searchTerm=full",
         "/api/v1/Contact/PagedSearch",
         {"pageSize": 20, "pageNumber": 1, "searchTerm": full_name}),
    ]

    found_with = []

    for label, endpoint, body in searches:
        status, raw = post(endpoint, body)
        items, err  = parse_items(raw, status)

        if items and len(items) > 0:
            # Check if any item actually matches the name
            matched = []
            for x in items:
                fn = (x.get("firstname") or x.get("firstName") or "").lower()
                sn = (x.get("surname")   or x.get("lastName")  or "").lower()
                nm = (x.get("name") or x.get("resourceName") or "").lower()
                if (first.lower() in fn or first.lower() in nm or
                    last.lower()  in sn or last.lower()  in nm):
                    matched.append(x)

            line = (f"  [{label}] → {len(items)} record(s)"
                    + (f"  *** {len(matched)} NAME MATCH ***" if matched else ""))
            print(line)
            block.append(line)
            for x in items[:5]:
                row = fmt_item(x)
                print(row)
                block.append(row)
            if matched:
                found_with.append(label)
        elif err:
            line = f"  [{label}] → status={status}  {err}"
            print(line)
            block.append(line)
        else:
            line = f"  [{label}] → 0 results  (status={status})"
            print(line)
            block.append(line)

    if found_with:
        summary = f"\n  *** FOUND via: {', '.join(found_with)} ***"
    else:
        summary = "\n  *** NOT FOUND by any method ***"
    print(summary)
    block.append(summary)
    all_output.extend(block)

# ── Write results file ────────────────────────────────────────────────────────
with open("diagnose_results.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(all_output))

print("\n\nDone — full results saved to C:\\AeroTracker\\diagnose_results.txt")
print("Please paste the contents of that file back here.")
