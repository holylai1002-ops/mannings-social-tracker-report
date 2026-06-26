"""Test all new FB Page API endpoints."""
import requests

s = requests.Session()
r = s.post("http://127.0.0.1:8000/login", data={"passcode": "fimmick26"}, allow_redirects=True)
print("Login status:", r.status_code)

resp = s.get("http://127.0.0.1:8000/api/data/fb_page?year=2026&month=5")
print("API status:", resp.status_code)
if resp.status_code != 200:
    print("Body:", resp.text[:500])
    exit(1)
data = resp.json()

fg = data.get("followers_growth")
tr = data.get("total_reach")
rf = data.get("reach_funnel")

print("=== Followers Growth ===")
if fg:
    print("  dates:", len(fg["dates"]), "days, sample:", fg["dates"][0])
    print("  monthly_net:", format(fg["monthly_net"], ","))
else:
    print("  None!")

print("\n=== Total Reach ===")
if tr:
    print("  dates:", len(tr["dates"]), "days, sample:", tr["dates"][0])
    print("  monthly_total:", format(tr["monthly_total"], ","))
else:
    print("  None!")

print("\n=== Reach Funnel ===")
if rf:
    print("  organic:", format(rf["organic"], ","))
    print("  paid:", format(rf["paid"], ","))
    print("  total:", format(rf["total"], ","))
else:
    print("  None!")

igf = data.get("ig_followers")
print("\n=== IG Followers ===")
if igf:
    print("  dates:", len(igf["dates"]), "days, sample:", igf["dates"][0])
    print("  monthly_net:", format(igf["monthly_net"], ","))
else:
    print("  None!")

print("\nAll OK!")
