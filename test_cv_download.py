"""
test_cv_download.py — Verify REST API CV download works (no web cookie needed).
Uses TAIF HAMEED (139906) who definitely has a CV in Tracker.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import requests, base64, json
from update_tracker import get_jwt, h, TRACKER_API

TAIF_ID = 139906

def main():
    jwt = get_jwt()

    # 1. List documents
    print(f"[1] GET /api/v1/Resource/{TAIF_ID}/Documents")
    r = requests.get(f"{TRACKER_API}/api/v1/Resource/{TAIF_ID}/Documents",
                     headers=h(jwt), timeout=15)
    print(f"  HTTP {r.status_code}  Content-Type: {r.headers.get('Content-Type','')}")
    try:
        docs_data = r.json()
        print(f"  Response type: {type(docs_data).__name__}")
        if isinstance(docs_data, list):
            docs = docs_data
        elif isinstance(docs_data, dict):
            docs = (docs_data.get("items") or docs_data.get("documents") or
                    docs_data.get("value") or docs_data.get("$values") or
                    docs_data.get("results") or docs_data.get("data") or [])
            print(f"  Dict keys: {list(docs_data.keys())}")
        else:
            docs = []
        print(f"  Documents found: {len(docs)}")
        for i, d in enumerate(docs[:3]):
            print(f"    [{i}] keys={list(d.keys())}  id={d.get('id') or d.get('documentId') or d.get('resourceDocumentId')}  name={d.get('filename') or d.get('name') or d.get('fileName')}")
    except Exception as e:
        print(f"  JSON parse error: {e}")
        print(f"  Raw: {r.text[:200]}")
        return

    if not docs:
        print("  No documents found — stopping")
        return

    # Pick first document and extract its ID
    doc = docs[0]
    doc_id = (doc.get("documentId") or doc.get("id") or
              doc.get("docId") or doc.get("resourceDocumentId") or
              doc.get("DocumentId") or doc.get("document_id"))
    print(f"\n  Using doc_id={doc_id}, keys={list(doc.keys())}")

    if not doc_id:
        print("  Cannot find document ID in response — check keys above")
        return

    # 2. Test each candidate download URL
    urls = [
        f"{TRACKER_API}/api/v1/Resource/{TAIF_ID}/Document/{doc_id}",        # swagger-correct (singular)
        f"{TRACKER_API}/api/v1/Resource/{TAIF_ID}/Documents/{doc_id}",       # plural variant
        f"{TRACKER_API}/api/v1/Resource/{TAIF_ID}/Documents/{doc_id}/Download",
        f"{TRACKER_API}/api/v1/Documents/{doc_id}/Download",
    ]

    print(f"\n[2] Testing download URLs for doc_id={doc_id}...")
    for url in urls:
        r = requests.get(url, headers=h(jwt), timeout=15)
        ct = r.headers.get("Content-Type", "")
        print(f"\n  URL: {url}")
        print(f"  HTTP {r.status_code}  Content-Type: {ct}  Size: {len(r.content)} bytes")

        if r.status_code == 200 and len(r.content) > 100:
            # Check if it's JSON (base64-wrapped)
            if "json" in ct.lower():
                try:
                    data = r.json()
                    print(f"  JSON keys: {list(data.keys()) if isinstance(data, dict) else type(data).__name__}")
                    # Look for base64 content
                    for k, v in (data.items() if isinstance(data, dict) else []):
                        if isinstance(v, str) and len(v) > 100:
                            try:
                                decoded = base64.b64decode(v)
                                print(f"  Key '{k}': base64-decoded to {len(decoded)} bytes, magic={decoded[:4]}")
                            except Exception:
                                print(f"  Key '{k}': string len={len(v)}, preview={v[:60]}")
                except Exception as e:
                    print(f"  JSON parse error: {e}")
            else:
                # Raw bytes
                magic = r.content[:4]
                print(f"  Raw bytes — magic={magic}  preview={r.content[:40]}")
                if magic[:3] == b'%PD' or magic == b'%PDF':
                    print("  -> PDF confirmed")
                elif magic[:2] == b'PK':
                    print("  -> DOCX/ZIP confirmed")
                print(f"  *** THIS URL WORKS ***")
            break
        elif r.status_code != 200:
            print(f"  Body: {r.text[:100]}")

if __name__ == "__main__":
    main()
