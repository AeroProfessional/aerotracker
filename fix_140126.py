"""Quick fix for profile 140126 Peerapat Akawittayasakul.
CV clearly shows: Marketing manager of AT PHONE co., ltd.
"""
import sys, os, requests
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from update_tracker import get_jwt, h, TRACKER_API

jwt = get_jwt()
r = requests.patch(
    f"{TRACKER_API}/api/v1/Resource/140126",
    headers=h(jwt),
    json={
        "jobTitle": "Marketing Manager",
        "currentClient": {"id": -1, "name": "AT PHONE co., ltd."},
    },
    timeout=15
)
ok = r.status_code in (200, 204)
print(f"{'✓' if ok else '✗'} [140126] Peerapat Akawittayasakul  (HTTP {r.status_code})")
if not ok:
    print(r.text[:300])
