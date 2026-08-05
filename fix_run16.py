"""
fix_run16.py — Corrections for the fourth 50-candidate batch.
All errors caused by regex parser (Groq was not active locally).
"""
import sys, os, requests
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from update_tracker import get_jwt, h, TRACKER_API

jwt = get_jwt()

def patch(resource_id, payload, label):
    r = requests.patch(
        f"{TRACKER_API}/api/v1/Resource/{resource_id}",
        headers=h(jwt), json=payload, timeout=15
    )
    ok = r.status_code in (200, 204)
    print(f"  {'✓' if ok else '✗'} [{resource_id}] {label}  (HTTP {r.status_code})")
    if not ok:
        print(f"      {r.text[:200]}")
    return ok

print("=" * 65)
print("  FIX RUN 16 PROFILES")
print("=" * 65)

# 1. Mohannad Alfawair — "Jordan Aviation Airline on B737-" → "Jordan Aviation"
print("\n[1/10] Mohannad Alfawair (139955)")
patch(139955, {
    "currentClient": {"id": -1, "name": "Jordan Aviation"},
}, "employer → 'Jordan Aviation' (was 'Jordan Aviation Airline on B737-')")

# 2. Qeist Mohd Ali — "Beechcraft King Air 300 (PT6A" → Unknown (aircraft type, not a company)
print("\n[2/10] Qeist Mohd Ali (106704)")
patch(106704, {
    "currentClient": {"id": -1, "name": "HM Aerospace Sdn Bhd"},
}, "employer → restored 'HM Aerospace Sdn Bhd' (was aircraft type 'Beechcraft King Air 300 (PT6A')")

# 3. Ranjeet Singh — "Air Arabia Dammam, Saudi Arabia" → "Air Arabia"
print("\n[3/10] Ranjeet Singh (51895)")
patch(51895, {
    "currentClient": {"id": -1, "name": "Air Arabia"},
}, "employer → 'Air Arabia' (was 'Air Arabia Dammam, Saudi Arabia' — location suffix stripped)")

# 4. Khalid Hamid — "May'23 to Date British Airways Ltd" → "British Airways"
print("\n[4/10] Khalid Hamid (74631)")
patch(74631, {
    "currentClient": {"id": -1, "name": "British Airways"},
}, "employer → 'British Airways' (was 'May'23 to Date British Airways Ltd' — date prefix stripped)")

# 5. Ahmed Shaker — "Aeronautical Engineer — Memphis Airlines" → "Memphis Airlines"
print("\n[5/10] Ahmed Shaker (139328)")
patch(139328, {
    "currentClient": {"id": -1, "name": "Memphis Airlines"},
}, "employer → 'Memphis Airlines' (was 'Aeronautical Engineer — Memphis Airlines' — job title prefix stripped)")

# 6. Camilo Trujillo — "Aviation Manager, Glencore Coal" → "Glencore Coal"
print("\n[6/10] Camilo Trujillo (139921)")
patch(139921, {
    "currentClient": {"id": -1, "name": "Glencore Coal"},
}, "employer → 'Glencore Coal' (was 'Aviation Manager, Glencore Coal' — job title prefix stripped)")

# 7. Abu Faisal — "Akij Aviation Ltd (Dhaka" → "Akij Aviation Ltd"
print("\n[7/10] Abu Faisal (139930)")
patch(139930, {
    "currentClient": {"id": -1, "name": "Akij Aviation Ltd"},
}, "employer → 'Akij Aviation Ltd' (was 'Akij Aviation Ltd (Dhaka' — location suffix stripped)")

# 8. Noor Halima Jahan — "AirAsia Airline, Kuala Lumpur, Malaysia" → "AirAsia"
print("\n[8/10] Noor Halima Jahan (139932)")
patch(139932, {
    "currentClient": {"id": -1, "name": "AirAsia"},
}, "employer → 'AirAsia' (was 'AirAsia Airline, Kuala Lumpur, Malaysia' — location suffix stripped)")

# 9. Huzaifa Balbale — job title "Resume" → Unknown
print("\n[9/10] Huzaifa Balbale (139934)")
patch(139934, {
    "jobTitle": "Unknown",
}, "job title → Unknown (was 'Resume' — CV header word picked up)")

# 10. Ayman Nafade — "7863799 CANADA INC.\"Ontario & QC" → "7863799 CANADA INC."
print("\n[10/10] Ayman Nafade (138780)")
patch(138780, {
    "currentClient": {"id": -1, "name": "7863799 CANADA INC."},
}, "employer → '7863799 CANADA INC.' (was with location suffix 'Ontario & QC')")

print("\n" + "=" * 65)
print("  Done.")
print("=" * 65)
