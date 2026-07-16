"""
2a_v2_year.py - Community Comment Extraction (ALL data, Excel Output)
Based on 2a_new.py with:
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

from config import BASE_DIR, GSHEET_COMMENTS_ID
from utils import get_gspread_client

TARGET_YEAR = 2026

OUTPUT_XLSX = BASE_DIR / f"Mannings_Comments_RAW_{TARGET_YEAR}.xlsx"

OUTPUT_COLUMNS = [
    "Post Date", "Post", "Comment Date", "Name", "Mannings Teams",
    "Message", "With attachment?", "Fimmick Suggested Reply/Action",
    "Mannings Team's Comment", "Digital team Final Check", "Fimmick Latest Action",
    "Link", "Category (comment)", "Type",
    "Sentiment\n- Positive\n- Neutral\n- Negative",
    "Month Note"
]

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

def filter_to_target_period(df, target_year):
    if df.empty:
        return df

    target_year_str = str(target_year)
    mask = pd.Series([True] * len(df), index=df.index)

    if 'Comment Date' in df.columns:
        cd_str = df['Comment Date'].fillna('').astype(str)
        mask &= cd_str.str.contains(target_year_str)

    filtered = df[mask].copy()
    logging.info(f"  Filtered to {target_year}: {len(filtered)} rows (from {len(df)})")
    return filtered

def extract_all_comments():
    logging.info("=== 2a_v2_year: Comment Extraction (ALL periods) ===")

    try:
        client = get_gspread_client()
        spreadsheet = client.open_by_key(GSHEET_COMMENTS_ID)
        logging.info("Connected to Community Comment Webbook Google Sheet.")
    except Exception as e:
        logging.error(f"Failed to connect: {e}")
        return None

    all_data = []
    all_worksheets = spreadsheet.worksheets()

    for worksheet in all_worksheets:
        tab_name = worksheet.title
        tab_name_lower = tab_name.lower()

        if "indicator" in tab_name_lower or "log" in tab_name_lower:
            logging.info(f"  Skipping tab (filtered): {tab_name}")
            continue

        try:
            raw_data = worksheet.get_all_values()
            if not raw_data or len(raw_data) < 2:
                continue

            headers = raw_data[0]
            data = raw_data[1:]
            df = pd.DataFrame(data, columns=headers)

            rename_map = {}
            for col in df.columns:
                col_lower = str(col).lower()
                if 'catagory' in col_lower or 'category' in col_lower:
                    rename_map[col] = 'Category (comment)'
                elif 'fimmick final check' in col_lower or 'digital team final check' in col_lower:
                    rename_map[col] = 'Digital team Final Check'
                elif 'monthly note' in col_lower or 'month note' in col_lower:
                    rename_map[col] = 'Month Note'

            df.rename(columns=rename_map, inplace=True)

            for target_col in OUTPUT_COLUMNS:
                if target_col not in df.columns:
                    matched = False
                    for existing_col in df.columns:
                        if str(target_col).replace('\n', '').replace(' ', '') == str(existing_col).replace('\n', '').replace(' ', ''):
                            df.rename(columns={existing_col: target_col}, inplace=True)
                            matched = True
                            break
                    if not matched:
                        df[target_col] = ''

            filtered_df = df[OUTPUT_COLUMNS].copy()
            filtered_df['Source_Tab'] = tab_name
            all_data.append(filtered_df)
            logging.info(f"  Tab '{tab_name}': {len(filtered_df)} rows")

        except Exception as e:
            logging.error(f"  Error reading tab '{tab_name}': {e}")

    if not all_data:
        logging.warning("No comments found in any tab.")
        return None

    final_df = pd.concat(all_data, ignore_index=True)
    final_df = filter_to_target_period(final_df, TARGET_YEAR)
    if final_df.empty:
        logging.warning(f"No comments found for {TARGET_YEAR}.")
        return None
    final_df = deep_clean_df(final_df)

    logging.info(f"  Output -> {OUTPUT_XLSX}")
    with pd.ExcelWriter(OUTPUT_XLSX, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name="Comments")
        from openpyxl.styles import Alignment
        worksheet = writer.book["Comments"]
        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=False)

    logging.info(f"DONE! {len(final_df)} total comments extracted.")
    return final_df

if __name__ == "__main__":
    extract_all_comments()
