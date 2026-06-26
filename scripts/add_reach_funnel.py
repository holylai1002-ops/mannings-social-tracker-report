"""Add Reach Funnel tab to feed file."""
import os
import shutil
import pandas as pd
from utils import get_gspread_client

client = get_gspread_client()
sh = client.open_by_key("1f9HLS0HXs2B-a_fxvvcUKuRKz0iMvmEDVivoLFl2vr8")
ws = sh.worksheet("All FB Posts_New Cat")
records = ws.get_all_records()
df = pd.DataFrame(records)
print("Source rows:", len(df))

# Find actual column names (some have trailing spaces)
pt_col = next((c for c in df.columns if c.strip() == "Publish time"), None)
mn_col = next((c for c in df.columns if c.strip() == "Month No."), None)
day_col = next((c for c in df.columns if c.strip() == "Day"), None)
org_col = next((c for c in df.columns if c.strip() == "Organic reach"), None)
paid_col = next((c for c in df.columns if c.strip() == "Paid reach"), None)

# Construct date from Month No. and Day (Year = 2026)
month_no = pd.to_numeric(df[mn_col], errors="coerce") if mn_col else pd.Series(dtype=float)
day = pd.to_numeric(df[day_col], errors="coerce") if day_col else pd.Series(dtype=float)
year = pd.Series([2026] * len(df))

# Build proper date strings
df_out = pd.DataFrame()
valid = month_no.notna() & day.notna() & (month_no >= 1) & (month_no <= 12) & (day >= 1) & (day <= 31)
year_int = year.astype(int)
df_out = df_out.assign(
    **{
        "Year": year_int.where(valid, other=pd.NA),
        "Month No.": month_no.astype("Int64"),
        "Day": day.astype("Int64"),
    }
)
# Construct Publish time as d/m/yyyy
dates = pd.to_datetime(
    df_out["Year"].astype("Int64").astype(str) + "-" +
    df_out["Month No."].astype(str).str.zfill(2) + "-" +
    df_out["Day"].astype(str).str.zfill(2),
    errors="coerce"
)
df_out["Publish time"] = dates.dt.strftime("%d/%m/%Y")

# Organic/Paid reach
for label, src in [("Organic reach", org_col), ("Paid reach", paid_col)]:
    if src:
        df_out[label] = pd.to_numeric(df[src].astype(str).str.replace(",", ""), errors="coerce").fillna(0).astype(int)
    else:
        df_out[label] = 0

df_out = df_out.dropna(subset=["Year", "Month No."]).reset_index(drop=True)
# Reorder columns
df_out = df_out[["Publish time", "Year", "Month No.", "Day", "Organic reach", "Paid reach"]]

print("Reach Funnel rows:", len(df_out))
print(df_out.head(5).to_string())
print()
months = df_out.groupby(["Year", "Month No."]).size()
print("Rows per month:")
print(months.to_string())

m = df_out[(df_out["Year"] == 2026) & (df_out["Month No."] == 5)]
organic = int(m["Organic reach"].sum())
paid = int(m["Paid reach"].sum())
print("\nMay 2026: Organic={}, Paid={}, Total={}".format(
    format(organic, ","), format(paid, ","), format(organic + paid, ",")))

# Write to feed (use backup approach to avoid lock issues)
import gc
src = "Mannings_FB_IG_Dashboard_Feed.xlsx"
dst = "Mannings_FB_IG_Dashboard_Feed_v2.xlsx"
xls = pd.ExcelFile(src)
all_sheets = {n: pd.read_excel(xls, n) for n in xls.sheet_names}
xls.close()
del xls
gc.collect()

if "Reach Funnel" in all_sheets:
    sheet_order = [n for n in all_sheets.keys() if n != "Reach Funnel"]
else:
    sheet_order = list(all_sheets.keys())
all_sheets["Reach Funnel"] = df_out
sheet_order.append("Reach Funnel")

with pd.ExcelWriter(dst, engine="openpyxl") as w:
    for n in sheet_order:
        all_sheets[n].to_excel(w, sheet_name=n[:31], index=False)

print("Written to", dst)
