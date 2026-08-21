import requests, json

BASE   = "https://evoglapi.tracker-rms.com"
BEARER = "b28cae06af044958afb45fa8b1445fa7"

JWT = requests.post(f"{BASE}/api/Auth/ExchangeToken",
                    json={"bearerToken": BEARER}, timeout=15).json()["token"]
HDR = {"Authorization": f"Bearer {JWT}"}

# Try fetching a page of results and inspect the fields
r = requests.post(f"{BASE}/api/v1/Resource/Search",
                  json={"pageSize": 5, "pageNumber": 1},
                  headers=HDR, timeout=30).json()

items = r if isinstance(r, list) else (r.get("items") or r.get("data") or r.get("value") or [])

if not items:
    print("No results returned. Raw response:")
    print(json.dumps(r, indent=2)[:500])
else:
    print("Available date fields on a candidate record:")
    sample = items[0]
    for k, v in sample.items():
        if v and ("date" in k.lower() or "created" in k.lower() or "registered" in k.lower() or "added" in k.lower()):
            print(f"  {k}: {v}")
    print("\nSample candidates:")
    for c in items:
        name = f"{c.get('firstname') or c.get('firstName','')} {c.get('surname','')}"
        print(f"  {name.strip()}")
