"""
3c_meta_api.py - Pull raw insights from Meta Graph API (FB Page)

Queries 3 daily metrics for the Mannings FB Page:
  - page_total_media_view_unique
  - page_daily_follows_unique
  - page_daily_unfollows_unique

Output: Mannings_API_raw.xlsx with 3 tabs (one per metric).
Each tab: end_time (col 1) | value (col 2).

Run after 3a/3b or independently. 3b reads this file to add
'Unique Page View' and 'Followers' tabs to the dashboard feed.
"""
import logging
import requests
import pandas as pd
from datetime import date
from dateutil.relativedelta import relativedelta

from config import BASE_DIR, META_FB_PAGE_ID, META_PERMANENT_PAGE_TOKEN, META_API_BASE

OUTPUT_FILE = BASE_DIR / "Mannings_API_raw.xlsx"

METRICS = [
    "page_total_media_view_unique",
    "page_daily_follows_unique",
    "page_daily_unfollows_unique",
]

START_DATE = "2026-01-01"


def _month_windows(start, end):
    """Yield (since, until) pairs in ~1-month chunks (max 90 days)."""
    chunks = []
    cur = pd.Timestamp(start)
    end_ts = pd.Timestamp(end)
    while cur < end_ts:
        nxt = min(cur + relativedelta(months=1), end_ts)
        chunks.append((cur.strftime("%Y-%m-%d"), nxt.strftime("%Y-%m-%d")))
        cur = nxt
    return chunks


def fetch_metric(metric, since=START_DATE, until=None):
    """Fetch a daily metric from the Graph API, paginating by month."""
    if until is None:
        until = date.today().isoformat()

    windows = _month_windows(since, until)
    all_rows = []

    for w_since, w_until in windows:
        url = f"{META_API_BASE}/{META_FB_PAGE_ID}/insights"
        params = {
            "metric": metric,
            "period": "day",
            "since": w_since,
            "until": w_until,
            "access_token": META_PERMANENT_PAGE_TOKEN,
        }

        logging.info(f"  {metric}: {w_since} → {w_until} ...")
        resp = requests.get(url, params=params, timeout=30)
        resp.raise_for_status()
        payload = resp.json()

        data_list = payload.get("data", [])
        if data_list:
            for entry in data_list:
                for dv in entry.get("values", []):
                    all_rows.append({
                        "end_time": dv.get("end_time", ""),
                        "value": dv.get("value", 0),
                    })

    df = pd.DataFrame(all_rows)
    if not df.empty:
        df["end_time"] = pd.to_datetime(df["end_time"]).dt.strftime("%Y-%m-%d")
        df["value"] = pd.to_numeric(df["value"], errors="coerce").fillna(0).astype(int)
        df = df.drop_duplicates(subset=["end_time"], keep="last").sort_values("end_time").reset_index(drop=True)

    logging.info(f"    {metric}: {len(df)} rows total")
    return df


def main():
    logging.info("=== 3c_meta_api: Pull Meta Graph API insights ===")

    if not META_PERMANENT_PAGE_TOKEN:
        logging.error("META_PERMANENT_PAGE_TOKEN not set in .env")
        return

    sheets = {}
    for metric in METRICS:
        try:
            sheets[metric] = fetch_metric(metric)
        except Exception as e:
            logging.error(f"  Failed to fetch {metric}: {e}")
            sheets[metric] = pd.DataFrame(columns=["end_time", "value"])

    logging.info(f"Writing {OUTPUT_FILE} ...")
    with pd.ExcelWriter(str(OUTPUT_FILE), engine="openpyxl") as writer:
        for metric, df in sheets.items():
            df.to_excel(writer, sheet_name=metric, index=False)

    logging.info(f"DONE! {OUTPUT_FILE}")


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
    main()
