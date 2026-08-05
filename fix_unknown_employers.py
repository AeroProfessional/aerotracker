"""Set employer to 'Unknown' for profiles where it was cleared to blank."""
import sys, os, requests
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from update_tracker import get_jwt, h, TRACKER_API

jwt = get_jwt()

profiles = [
    (139982, "Meram Mohamed"),
    (140112, "Carlo Rizza"),
    (139769, "Giorgio Micoli"),
    (137816, "Ananthan Joshua"),
]

for rid, name in profiles:
    r = requests.patch(
        f"{TRACKER_API}/api/v1/Resource/{rid}",
        headers=h(jwt),
        json={"currentClient": {"id": -1, "name": "Unknown"}},
        timeout=15
    )
    ok = r.status_code in (200, 204)
    print(f"  {'✓' if ok else '✗'} [{rid}] {name} → employer set to 'Unknown'  (HTTP {r.status_code})")
