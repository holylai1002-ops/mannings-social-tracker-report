"""
1b_period_v2.py - Instagram Post & Story Processing (Target Period, Excel Output)
Based on 1b_new_period.py but:
  - Outputs .xlsx instead of .csv
  - deep_clean_df for Excel-safe characters
  - No text wrap on cells
"""
import warnings
import pandas as pd
warnings.filterwarnings('ignore', category=pd.errors.PerformanceWarning)
import re
import logging
from config import RAW_IG_POST_DIR, RAW_IG_STORY_DIR, BASE_DIR
from utils import (
    merge_csv_files, dedupe_by_id, convert_tz_la_to_hk,
    filter_valid_years, add_time_columns, normalize_columns
)

TARGET_YEAR = 2026
TARGET_MONTH = 6

OUTPUT_IG_POST = BASE_DIR / f"All_IG_Posts_{TARGET_YEAR}_{TARGET_MONTH:02d}.xlsx"
OUTPUT_IG_STORY = BASE_DIR / f"All_IGS_{TARGET_YEAR}_{TARGET_MONTH:02d}.xlsx"

_OPENPYXL_ILLEGAL_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f\ufffe\uffff]')

def strip_illegal_for_excel(val):
    if val is None:
        return ""
    if isinstance(val, (int, float, bool)):
        return val
    s = str(val)
    s = _OPENPYXL_ILLEGAL_RE.sub('', s)
    s = re.sub(r'[\x7f-\x9f\u200b-\u200f\u2028-\u202f\ufeff]', '', s)
    return s

def deep_clean_df(df):
    df = df.copy()
    for col in df.columns:
        df[col] = df[col].apply(strip_illegal_for_excel)
    return df

def calc_post_share_rate(row):
    try:
        shares = pd.to_numeric(row['Shares'], errors='coerce')
        reach = pd.to_numeric(row['Total Post Reach'], errors='coerce')
        if pd.isna(shares) or pd.isna(reach) or reach == 0:
            return "0.00%"
        return f"{shares / reach:.2%}"
    except Exception:
        return "0.00%"

def calc_story_share_rate(row):
    try:
        shares = pd.to_numeric(row['Shares'], errors='coerce')
        reach = pd.to_numeric(row['Total Reach'], errors='coerce')
        if pd.isna(shares) or pd.isna(reach) or reach == 0:
            return "0.00%"
        return f"{shares / reach:.2%}"
    except Exception:
        return "0.00%"

def filter_to_target_period(df, target_year, target_month):
    if df.empty:
        return df
    if 'HK Time DT' in df.columns:
        pt = df['HK Time DT']
    elif 'Publish time' in df.columns:
        pt = pd.to_datetime(df['Publish time'], errors='coerce')
    else:
        return df
    mask = (pt.dt.year == target_year) & (pt.dt.month == target_month)
    filtered = df[mask].copy()
    logging.info(f"  Filtered to {target_year}-{target_month:02d}: {len(filtered)} rows (from {len(df)})")
    return filtered

def write_xlsx(df, output_path, sheet_name):
    df = deep_clean_df(df)
    logging.info(f"  Output -> {output_path}")
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        from openpyxl.styles import Alignment
        worksheet = writer.book[sheet_name]
        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=False)

def process_ig_post():
    logging.info(f"=== 1b_period_v2: IG Post Processing ({TARGET_YEAR}-{TARGET_MONTH:02d}) ===")

    logging.info("Step 1: Merging all raw CSVs from IG raw/post/ ...")
    df = merge_csv_files(RAW_IG_POST_DIR)
    if df.empty:
        logging.error("No IG Post raw data found.")
        return None

    logging.info("Step 2: Deduplicating by Post ID ...")
    df = dedupe_by_id(df, 'Post ID')

    logging.info("Step 3: Timezone conversion (LA -> HK) ...")
    df = convert_tz_la_to_hk(df, 'Publish time')

    logging.info(f"Step 4: Filtering to {TARGET_YEAR} ...")
    df = filter_valid_years(df, 'HK Time DT', [TARGET_YEAR])

    logging.info(f"Step 5: Filtering to target period {TARGET_YEAR}-{TARGET_MONTH:02d} ...")
    df = filter_to_target_period(df, TARGET_YEAR, TARGET_MONTH)
    if df.empty:
        logging.warning(f"No IG Posts found for {TARGET_YEAR}-{TARGET_MONTH:02d}.")
        return None

    logging.info("Step 6: Renaming and computing metrics ...")
    df.rename(columns={
        'Description': 'Post Message',
        'Duration (sec)': 'Video Length',
        'Permalink': 'Post Link',
        'Post type': 'Type',
        'Reach': 'Total Post Reach',
    }, inplace=True)

    for col in ['Likes', 'Comments', 'Shares', 'Saves', 'Views']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        else:
            df[col] = 0

    df['Total Interactions'] = df['Likes'] + df['Comments'] + df['Shares'] + df['Saves']
    df['Share Rate'] = df.apply(calc_post_share_rate, axis=1)

    df = add_time_columns(df, 'HK Time DT')
    df = normalize_columns(df)

    final_cols = [
        'Month', 'Post Link', 'Post Message', 'Type', 'Pillar', 'Category',
        'Sub-Category', 'Campaign Name', 'Publish time', 'Likes', 'Comments',
        'Shares', 'Saves', 'Total Interactions', 'Views', 'Total Post Reach',
        'Share Rate', 'Video Length', 'Month No.', 'Day', 'Hour'
    ]
    for col in final_cols:
        if col not in df.columns:
            df[col] = ""
    df = df[final_cols]

    write_xlsx(df, OUTPUT_IG_POST, "IG_Posts")
    logging.info(f"DONE! {len(df)} IG Posts processed for {TARGET_YEAR}-{TARGET_MONTH:02d}.")
    return df

def process_ig_story():
    logging.info(f"=== 1b_period_v2: IG Story Processing ({TARGET_YEAR}-{TARGET_MONTH:02d}) ===")

    logging.info("Step 1: Merging all raw CSVs from IG raw/story/ ...")
    df = merge_csv_files(RAW_IG_STORY_DIR)
    if df.empty:
        logging.error("No IG Story raw data found.")
        return None

    logging.info("Step 2: Deduplicating by Post ID ...")
    df = dedupe_by_id(df, 'Post ID')

    logging.info("Step 3: Timezone conversion (LA -> HK) ...")
    df = convert_tz_la_to_hk(df, 'Publish time')

    logging.info(f"Step 4: Filtering to {TARGET_YEAR} ...")
    df = filter_valid_years(df, 'HK Time DT', [TARGET_YEAR])

    logging.info(f"Step 5: Filtering to target period {TARGET_YEAR}-{TARGET_MONTH:02d} ...")
    df = filter_to_target_period(df, TARGET_YEAR, TARGET_MONTH)
    if df.empty:
        logging.warning(f"No IG Stories found for {TARGET_YEAR}-{TARGET_MONTH:02d}.")
        return None

    logging.info("Step 6: Renaming and computing metrics ...")
    df.rename(columns={
        'Reach': 'Total Reach',
        'Permalink': 'Post Link',
    }, inplace=True)

    if 'Shares' not in df.columns:
        df['Shares'] = 0
    if 'Total Reach' not in df.columns:
        df['Total Reach'] = 0

    df['Share Rate'] = df.apply(calc_story_share_rate, axis=1)

    for col in ['Total Reach', 'Shares', 'Link clicks', 'Taps forward', 'Taps back', 'Exits', 'Replies']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

    df = add_time_columns(df, 'HK Time DT')
    df = normalize_columns(df)

    final_cols = [
        'Month', 'Post Link', 'Description', 'Pillar', 'Category',
        'Sub-Category', 'Campaign Name', 'Publish time', 'Total Reach',
        'Shares', 'Share Rate', 'Link clicks', 'Month No.', 'Day', 'Hour'
    ]
    for col in final_cols:
        if col not in df.columns:
            df[col] = ""
    df = df[final_cols]

    write_xlsx(df, OUTPUT_IG_STORY, "IG_Stories")
    logging.info(f"DONE! {len(df)} IG Stories processed for {TARGET_YEAR}-{TARGET_MONTH:02d}.")
    return df

if __name__ == "__main__":
    process_ig_post()
    process_ig_story()
