"""Test FastAPI app using TestClient (no server needed)."""
import sys
import os
sys.path.insert(0, r"C:\Users\holylai\Documents\n8n\Mannings")
os.chdir(r"C:\Users\holylai\Documents\n8n\Mannings")

from fastapi.testclient import TestClient
from app.main import app

client = TestClient(app)

tests = [
    ("Home page", "GET", "/?year=2026&month=5", None),
    ("FB Posts page", "GET", "/fb-posts?year=2026&month=5", None),
    ("Instagram page", "GET", "/instagram?year=2026&month=5", None),
    ("KPIs", "GET", "/api/data/kpis?year=2026&month=5", None),
    ("FB Page data", "GET", "/api/data/fb_page?year=2026&month=5", None),
    ("FB Posts data", "GET", "/api/data/fb_posts?year=2026&month=5", None),
    ("Instagram data", "GET", "/api/data/instagram?year=2026&month=5", None),
    ("Periods", "GET", "/api/periods", None),
    ("Images", "GET", "/api/images?year=2026&month=5", None),
]

print("=== FastAPI TestClient Results ===\n")
for name, method, path, body in tests:
    try:
        resp = client.get(path) if method == "GET" else client.post(path, json=body)
        content_length = len(resp.content)
        if resp.status_code == 200:
            preview = resp.text[:100].replace("\n", " ")
            print(f"  [OK]   {name:20s} {resp.status_code} ({content_length} bytes)")
            if "api" in path:
                print(f"         Preview: {preview}...")
        else:
            print(f"  [FAIL] {name:20s} {resp.status_code}")
            print(f"         Error: {resp.text[:300]}")
    except Exception as e:
        print(f"  [ERR]  {name:20s} {type(e).__name__}: {e}")

print("\n=== Done ===")
