"""
1b_v2_year.py - Instagram Post & Story Processing (ALL data, Excel Output)
Based on 1b_new.py with:
  - Outputs .xlsx instead of .csv
  - deep_clean_df for Excel-safe characters
  - No text wrap on cells
  - filter to target year
"""
import warnings
import pandas as pd
import logging
import re
warnings.filterwarnings('ignore', category=pd.errors.PerformanceWarning)

from config import RAW_IG_POST_DIR, RAW_IG_STORY_DIR, BASE_DIR
from utils import (
    merge_csv_files, dedupe_by_id, convert_tz_la_to_hk,
    filter_valid_years, add_time_columns, normalize_columns
)

TARGET_YEAR = 2026

OUTPUT_IG_POST = BASE_DIR / f"All_IG_Posts_{TARGET_YEAR}.xlsx"
OUTPUT_IG_STORY = BASE_DIR / f"All_IGS_{TARGET_YEAR}.xlsx"

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
    logging.info("=== 1b_new_v2: IG Post Processing (ALL periods) ===")
    df = merge_csv_files(RAW_IG_POST_DIR)
    if df.empty:
        logging.error("No IG Post raw data found.")
        return None
    df = dedupe_by_id(df, 'Post ID')
    df = convert_tz_la_to_hk(df, 'Publish time')
    df = filter_valid_years(df, 'HK Time DT', [TARGET_YEAR])
    df.rename(columns={
        'Description': 'Post Message', 'Duration (sec)': 'Video Length',
        'Permalink': 'Post Link', 'Post type': 'Type', 'Reach': 'Total Post Reach',
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
    logging.info(f"DONE! {len(df)} IG Posts processed.")
    return df

def process_ig_story():
    logging.info("=== 1b_new_v2: IG Story Processing (ALL periods) ===")
    df = merge_csv_files(RAW_IG_STORY_DIR)
    if df.empty:
        logging.error("No IG Story raw data found.")
        return None
    df = dedupe_by_id(df, 'Post ID')
    df = convert_tz_la_to_hk(df, 'Publish time')
    df = filter_valid_years(df, 'HK Time DT', [TARGET_YEAR])
    df.rename(columns={'Reach': 'Total Reach', 'Permalink': 'Post Link'}, inplace=True)
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
    logging.info(f"DONE! {len(df)} IG Stories processed.")
    return df

if __name__ == "__main__":
    process_ig_post()
    process_ig_story()
