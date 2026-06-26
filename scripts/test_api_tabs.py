"""Test new API endpoints."""
import app.db.reader as r
r.settings.excel_path = 'Mannings_FB_IG_Dashboard_Feed_new.xlsx'
r._excel_cache.clear()
r._PERIOD_CACHE.clear()

from starlette.testclient import TestClient
from app.main import app
c = TestClient(app)
c.post('/login', data={'passcode': 'fimmick26'})

resp = c.get('/api/data/fb_page?year=2026&month=5')
data = resp.json()
fg = data.get('followers_growth')
tr = data.get('total_reach')

print("Followers Growth:")
if fg:
    print("  dates:", len(fg['dates']), 'days')
    print("  monthly_net:", fg['monthly_net'])
    print("  sample gain/loss/net:", fg['gain'][0], fg['loss'][0], fg['net'][0])
else:
    print("  None!")

print()
print("Total Reach:")
if tr:
    print("  dates:", len(tr['dates']), 'days')
    print("  monthly_total:", tr['monthly_total'])
    print("  monthly_total (M): {:.2f} M".format(tr['monthly_total'] / 1e6))
else:
    print("  None!")

print()
print("All OK!")
