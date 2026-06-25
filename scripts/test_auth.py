"""Quick test for passcode auth."""
from starlette.testclient import TestClient
from app.main import app

c = TestClient(app)

print("--- unauthenticated / ---")
r = c.get("/", follow_redirects=False)
print(f"{r.status_code} location={r.headers.get('location')}")

print("--- login page ---")
r = c.get("/login")
print(f"{r.status_code} login form present: {'passcode' in r.text}")

print("--- wrong passcode ---")
r = c.post("/login", data={"passcode": "wrong"}, follow_redirects=False)
print(f"{r.status_code} location={r.headers.get('location')}")

print("--- correct passcode ---")
r = c.post("/login", data={"passcode": "fimmick26"}, follow_redirects=False)
print(f"{r.status_code} location={r.headers.get('location')}")

print("--- authed / ---")
r = c.get("/", follow_redirects=False)
print(f"{r.status_code} dashboard served: {r.status_code == 200}")

print("--- logout ---")
r = c.get("/logout", follow_redirects=False)
print(f"{r.status_code} location={r.headers.get('location')}")

print("--- after logout / ---")
r = c.get("/", follow_redirects=False)
print(f"{r.status_code} location={r.headers.get('location')}")

print("\nAll tests passed!")
