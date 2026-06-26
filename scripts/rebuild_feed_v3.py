"""Rebuild feed Excel with API Log tabs (FB Followers, IG Followers, FB Reach Funnel)."""
import gc
import shutil
import pandas as pd
from utils import get_gspread_client

FEED = "Mannings_FB_IG_Dashboard_Feed.xlsx"
DST = "Mannings_FB_IG_Dashboard_Feed_v3.xlsx"

client = get_gspread_client()
sh = client.open_by_key("1f9HLS0HXs2B-a_fxvvcUKuRKz0iMvmEDVivoLFl2vr8")

# Read existing feed
xls = pd.ExcelFile(FEED)
all_sheets = {n: pd.read_excel(xls, n) for n in xls.sheet_names}
sheet_order = list(xls.sheet_names)
xls.close()
del xls
gc.collect()

# Remove old tabs that are being replaced
for old in ["Followers", "Reach Funnel", "Unique Page View", "FB Followers", "IG Followers", "FB Reach Funnel"]:
    if old in all_sheets:
        del all_sheets[old]
        sheet_order = [s for s in sheet_order if s != old]

# ── FB API Log ──
ws_fb = sh.worksheet("FB API Log")
df_fb = pd.DataFrame(ws_fb.get_all_records())
print("FB API Log:", len(df_fb), "rows")

# Unique Page View tab
df_view = df_fb[["Date", "Unique Page View"]].copy()
df_view["Unique Page View"] = pd.to_numeric(df_view["Unique Page View"], errors="coerce").fillna(0).astype(int)
all_sheets["Unique Page View"] = df_view
sheet_order.append("Unique Page View")
print("  Unique Page View:", len(df_view), "rows")

# FB Followers tab
fb_cols = ["Date", "Followers Gain", "Followers Loss", "Followers Net"]
fb_avail = [c for c in fb_cols if c in df_fb.columns]
df_fbf = df_fb[fb_avail].copy()
for c in ["Followers Gain", "Followers Loss", "Followers Net"]:
    if c in df_fbf.columns:
        df_fbf[c] = pd.to_numeric(df_fbf[c], errors="coerce").fillna(0).astype(int)
all_sheets["FB Followers"] = df_fbf
sheet_order.append("FB Followers")
print("  FB Followers:", len(df_fbf), "rows")

# ── IG API Log with forward-fill ──
ws_ig = sh.worksheet("IG API Log")
df_ig = pd.DataFrame(ws_ig.get_all_records())
print("IG API Log:", len(df_ig), "rows")

df_ig["Date"] = df_ig["Date"].astype(str).str.strip()
df_ig["Followers Net"] = pd.to_numeric(df_ig.get("Followers Net", 0), errors="coerce").fillna(0).astype(int)
dates = pd.to_datetime(df_ig["Date"], format="%d/%m/%Y", errors="coerce")
df_ig["_dt"] = dates

min_dt = df_ig["_dt"].min()
max_dt = df_ig["_dt"].max()
full_range = pd.date_range(min_dt, max_dt, freq="D")
df_daily = pd.DataFrame({"_dt": full_range})
df_daily = df_daily.merge(df_ig, on="_dt", how="left")

if "Total Followers" in df_daily.columns:
    tf = pd.to_numeric(df_daily["Total Followers"], errors="coerce")
    df_daily["Total Followers"] = tf.ffill().bfill().astype(int)

df_daily["Followers Net"] = df_daily["Followers Net"].fillna(0).astype(int)
mask_empty = df_daily["Followers Net"] == 0
if mask_empty.any() and "Total Followers" in df_daily.columns:
    diffs = df_daily["Total Followers"].diff().fillna(0).astype(int)
    df_daily.loc[mask_empty, "Followers Net"] = diffs[mask_empty]

df_ig_out = pd.DataFrame({
    "Date": df_daily["_dt"].dt.strftime("%d/%m/%Y"),
    "Followers Net": df_daily["Followers Net"].astype(int),
})
all_sheets["IG Followers"] = df_ig_out
sheet_order.append("IG Followers")
print("  IG Followers (forward-filled):", len(df_ig_out), "rows")

# ── FB Reach Funnel (from All FB Posts_New Cat) ──
ws_posts = sh.worksheet("All FB Posts_New Cat")
df_posts = pd.DataFrame(ws_posts.get_all_records())
col_map = {}
for c in df_posts.columns:
    s = c.strip()
    if s in ("Publish time", "Month No.", "Day", "Organic reach", "Paid reach"):
        col_map[s] = c  # stripped → actual

mn_actual = col_map.get("Month No.")
day_actual = col_map.get("Day")
month_no = pd.to_numeric(df_posts[mn_actual], errors="coerce") if mn_actual else pd.Series(dtype=float)
day = pd.to_numeric(df_posts[day_actual], errors="coerce") if day_actual else pd.Series(dtype=float)
year_int = pd.Series([2026] * len(df_posts))
valid = month_no.notna() & day.notna() & (month_no >= 1) & (month_no <= 12) & (day >= 1) & (day <= 31)

dates_rf = pd.to_datetime(
    year_int.astype(int).astype(str).where(valid, other="NaT") + "-" +
    month_no.astype("Int64").astype(str).str.zfill(2) + "-" +
    day.astype("Int64").astype(str).str.zfill(2),
    errors="coerce"
)

df_rf = pd.DataFrame({
    "Publish time": dates_rf.dt.strftime("%d/%m/%Y"),
    "Year": year_int.astype(int).where(valid, other=pd.NA),
    "Month No.": month_no.astype("Int64"),
    "Day": day.astype("Int64"),
})
for label in ("Organic reach", "Paid reach"):
    src = col_map.get(label)
    if src:
        df_rf[label] = pd.to_numeric(df_posts[src].astype(str).str.replace(",", ""), errors="coerce").fillna(0).astype(int)
    else:
        df_rf[label] = 0

df_rf = df_rf.dropna(subset=["Year", "Month No."]).reset_index(drop=True)
all_sheets["FB Reach Funnel"] = df_rf
sheet_order.append("FB Reach Funnel")
print("  FB Reach Funnel:", len(df_rf), "rows")

# Write
with pd.ExcelWriter(DST, engine="openpyxl") as w:
    for n in sheet_order:
        all_sheets[n].to_excel(w, sheet_name=n[:31], index=False)

print("\nWritten to", DST)
print("Tabs:", sheet_order)

# Verify
m = df_fbf[pd.to_datetime(df_fbf["Date"], format="%d/%m/%Y").dt.month == 5]
print("\nMay 2026 FB Followers net:", format(int(m["Followers Net"].sum()), ","))

ig_may = df_ig_out[pd.to_datetime(df_ig_out["Date"], format="%d/%m/%Y").dt.month == 5]
print("May 2026 IG Followers net:", format(int(ig_may["Followers Net"].sum()), ","))

rf_may = df_rf[(df_rf["Year"] == 2026) & (df_rf["Month No."] == 5)]
print("May 2026 FB Reach Funnel: organic={}, paid={}".format(
    format(int(rf_may["Organic reach"].sum()), ","),
    format(int(rf_may["Paid reach"].sum()), ",")
))
