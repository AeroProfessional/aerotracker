"""
fix_run13.py — One-time fixes for profiles incorrectly updated in the 50-candidate test run.

Errors were caused by bugs now fixed in update_tracker.py:
  - Aircraft type alone used as job title (hyphen-strip regex was too greedy)
  - Flight Deck → Cabin Crew downgrade (no work-type protection)
  - Employer section headers / garbage phrases not caught by plausibility check
  - Generic single-word terms (e.g. "AIRLINES") passing plausibility

Candidates fixed here:
  140126  Peerapat Akawittayasakul  — employer cleared (was "Experience of Aviation")
  13270   Peter Kelly               — employer restored to "Killick Aerospace"
  140125  Ahmad Badran              — job title restored to "Head Of Technical Outstations"
  139982  Meram Mohamed             — employer cleared (was "Aviation & Aerospace Engineer")
  103288  Dmytro Koshytskyi         — work type → Flight Deck, remove Main Crew skill
  110022  Julio Guillen             — work type → Flight Deck, job title → Unknown, remove Main Crew
  106930  Yazan Khalayleh           — employer → "Jordan Civil Aviation Regulatory Commission"
  23150   Christophe Gestraud       — job title restored, employer restored to "Turkish Airlines"
  140012  Metin Tekin               — employer restored to "Turkish Airlines"
  140112  Carlo Rizza               — employer cleared (was "AVIATION EMPLOYMENT HISTORY:")
  140106  Daniel Lazcano            — employer → "IBEROJET" (stripped "EMPLOYER. " prefix)
  139769  Giorgio Micoli            — employer cleared (was "PROFESSIONAL AVIATION EXPERIENCE")
  140104  Rabah Lassad              — job title → Unknown, employer cleared
  55722   Florent Arsac             — employer → "Fly Amelia by Regourd Aviation"
  128376  Raphaël Leveillé          — employer restored to "Pelita Air Service"
  140114  Elimane Soussoko          — job title → "B737-500 First Officer"

Run:  py fix_run13.py
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
        print(f"      Response: {r.text[:200]}")
    return ok

def get_skills(resource_id):
    """Fetch current quickSkills list from Tracker."""
    r = requests.get(f"{TRACKER_API}/api/v1/Resource/{resource_id}",
                     headers=h(jwt), timeout=15)
    if r.status_code != 200:
        return None, None
    rec = r.json()
    return rec.get("quickSkills") or [], rec.get("workTypes") or []

print("=" * 65)
print("  FIX RUN 13 PROFILES")
print("=" * 65)

# ── 1. Peerapat Akawittayasakul (140126) ───────────────────────────────────────
print("\n[1/16] Peerapat Akawittayasakul (140126)")
patch(140126, {
    "currentClient": {"id": -1, "name": ""},
}, "employer cleared (was 'Experience of Aviation')")

# ── 2. Peter Kelly (13270) ─────────────────────────────────────────────────────
print("\n[2/16] Peter Kelly (13270)")
patch(13270, {
    "currentClient": {"id": -1, "name": "Killick Aerospace"},
}, "employer restored → 'Killick Aerospace'")

# ── 3. Ahmad Badran (140125) ───────────────────────────────────────────────────
print("\n[3/16] Ahmad Badran (140125)")
patch(140125, {
    "jobTitle": "Head Of Technical Outstations",
}, "job title restored → 'Head Of Technical Outstations'")

# ── 4. Meram Mohamed (139982) ──────────────────────────────────────────────────
print("\n[4/16] Meram Mohamed (139982)")
patch(139982, {
    "currentClient": {"id": -1, "name": ""},
}, "employer cleared (was 'Aviation & Aerospace Engineer')")

# ── 5. Dmytro Koshytskyi (103288) — restore Flight Deck, remove Main Crew ──────
print("\n[5/16] Dmytro Koshytskyi (103288)")
skills_103288, _ = get_skills(103288)
if skills_103288 is not None:
    cleaned = [s for s in skills_103288 if (s.get("name") or "").strip().lower() != "main crew"]
    patch(103288, {
        "workTypes": [{"id": 472, "name": "Flight Deck"}],
        "quickSkills": cleaned,
    }, "work type → Flight Deck, Main Crew skill removed")
else:
    print("  ✗ Could not fetch current record — skipping")

# ── 6. Julio Guillen (110022) — restore Flight Deck, job title → Unknown ───────
print("\n[6/16] Julio Guillen (110022)")
skills_110022, _ = get_skills(110022)
if skills_110022 is not None:
    cleaned = [s for s in skills_110022 if (s.get("name") or "").strip().lower() != "main crew"]
    patch(110022, {
        "jobTitle": "Unknown",
        "workTypes": [{"id": 472, "name": "Flight Deck"}],
        "quickSkills": cleaned,
    }, "work type → Flight Deck, job title → Unknown, Main Crew removed")
else:
    print("  ✗ Could not fetch current record — skipping")

# ── 7. Yazan Khalayleh (106930) ────────────────────────────────────────────────
print("\n[7/16] Yazan Khalayleh (106930)")
patch(106930, {
    "currentClient": {"id": -1, "name": "Jordan Civil Aviation Regulatory Commission"},
}, "employer → 'Jordan Civil Aviation Regulatory Commission' (trailing '[' stripped)")

# ── 8. Christophe Gestraud (23150) — restore title and employer ────────────────
print("\n[8/16] Christophe Gestraud (23150)")
patch(23150, {
    "jobTitle": "Captain Airbus 320 NEO.",
    "currentClient": {"id": -1, "name": "Turkish Airlines"},
}, "job title restored → 'Captain Airbus 320 NEO.', employer → 'Turkish Airlines'")

# ── 9. Metin Tekin (140012) ────────────────────────────────────────────────────
print("\n[9/16] Metin Tekin (140012)")
patch(140012, {
    "currentClient": {"id": -1, "name": "Turkish Airlines"},
}, "employer restored → 'Turkish Airlines' (was 'AIRLINES')")

# ── 10. Carlo Rizza (140112) ───────────────────────────────────────────────────
print("\n[10/16] Carlo Rizza (140112)")
patch(140112, {
    "currentClient": {"id": -1, "name": ""},
}, "employer cleared (was 'AVIATION EMPLOYMENT HISTORY:')")

# ── 11. Daniel Lazcano (140106) ────────────────────────────────────────────────
print("\n[11/16] Daniel Lazcano (140106)")
patch(140106, {
    "currentClient": {"id": -1, "name": "IBEROJET"},
}, "employer → 'IBEROJET' (stripped 'EMPLOYER. ' prefix)")

# ── 12. Giorgio Micoli (139769) ────────────────────────────────────────────────
print("\n[12/16] Giorgio Micoli (139769)")
patch(139769, {
    "currentClient": {"id": -1, "name": ""},
}, "employer cleared (was 'PROFESSIONAL AVIATION EXPERIENCE')")

# ── 13. Rabah Lassad (140104) ──────────────────────────────────────────────────
print("\n[13/16] Rabah Lassad (140104)")
patch(140104, {
    "jobTitle": "Unknown",
    "currentClient": {"id": -1, "name": ""},
}, "job title → Unknown (was 'Mastered languages'), employer cleared")

# ── 14. Florent Arsac (55722) ──────────────────────────────────────────────────
print("\n[14/16] Florent Arsac (55722)")
patch(55722, {
    "currentClient": {"id": -1, "name": "Fly Amelia by Regourd Aviation"},
}, "employer → 'Fly Amelia by Regourd Aviation' (stripped 'Current ' prefix)")

# ── 15. Raphaël Leveillé (128376) ─────────────────────────────────────────────
print("\n[15/16] Raphaël Leveillé (128376)")
patch(128376, {
    "currentClient": {"id": -1, "name": "Pelita Air Service"},
}, "employer restored → 'Pelita Air Service' (was 'AIRLINE CAPTAIN')")

# ── 16. Elimane Soussoko (140114) ─────────────────────────────────────────────
print("\n[16/16] Elimane Soussoko (140114)")
patch(140114, {
    "jobTitle": "B737-500 First Officer",
}, "job title → 'B737-500 First Officer' (was 'B737' due to hyphen-strip bug)")

print("\n" + "=" * 65)
print("  Done.")
print("=" * 65)
