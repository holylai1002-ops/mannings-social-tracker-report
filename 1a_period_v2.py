"""
1a_period.py - Facebook Post Processing (Target Period Version)
Based on 1a_new.py but filtered to a specific TARGET_YEAR / TARGET_MONTH.
  - Merges ALL CSVs in FB raw/ folder (no single file path)
  - Removes duplicates by Post ID (overlapping export date ranges)
  - Filters to TARGET_YEAR-TARGET_MONTH only
  - Updated classification rules (Ecommerce, GNC, Newness, expanded price KW)
  - Dark Post merge from Judy's Sheet for the TARGET period only
  - Outputs processed CSV for upload to Social Tracker Google Sheet
"""
import warnings
import pandas as pd
warnings.filterwarnings('ignore', category=pd.errors.PerformanceWarning)

import gspread
import re
import logging
from datetime import datetime
from config import RAW_FB_DIR, BASE_DIR, GSHEET_JUDY_ID
from utils import (
    merge_csv_files, dedupe_by_id, convert_tz_la_to_hk,
    filter_valid_years, add_time_columns, normalize_columns,
    get_gspread_client
)


# ==============================================================================
# CONFIGURATION BLOCK (Target Period Filter)
# ==============================================================================
TARGET_YEAR = 2026
TARGET_MONTH = 6  # e.g., 4 for April

OUTPUT_FILE = BASE_DIR / f"All FB Posts_New Cat_{TARGET_YEAR}_{TARGET_MONTH:02d}.xlsx"


# This regex targets the standard illegal characters defined by XML/OpenPyXL
_OPENPYXL_ILLEGAL_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f\ufffe\uffff]')

def strip_illegal_for_excel(val):
    """Remove every character openpyxl will reject, preserving valid text/emoji."""
    if val is None:
        return ""
    if isinstance(val, (int, float, bool)):
        return val
    s = str(val)
    s = _OPENPYXL_ILLEGAL_RE.sub('', s)
    # Extra layer to target invisible formatting characters often found in social media
    s = re.sub(r'[\x7f-\x9f\u200b-\u200f\u2028-\u202f\ufeff]', '', s)
    return s

def deep_clean_df(df):
    """Apply strip_illegal_for_excel to every cell in the DataFrame."""
    df = df.copy()
    for col in df.columns:
        df[col] = df[col].apply(strip_illegal_for_excel)
    return df

def fetch_gsheet_dark_posts(target_year, target_month):
    """Fetches dark posts from Judy's Sheet — only the target month tab."""
    try:
        client = get_gspread_client()
        sh = client.open_by_key(GSHEET_JUDY_ID)
    except Exception as e:
        logging.warning(f"Cannot connect to Judy's Sheet: {e}")
        return pd.DataFrame()

    target_dt = datetime(target_year, target_month, 1)
    tab_name = target_dt.strftime('%b %Y')

    try:
        logging.info(f"  Reading Judy Sheet tab '{tab_name}' ...")
        ws = sh.worksheet(tab_name)
        data = ws.get_all_values()
    except Exception as e:
        logging.warning(f"Cannot read tab '{tab_name}': {e}")
        return pd.DataFrame()

    if not data or len(data) < 2:
        logging.info(f"  Tab '{tab_name}' is empty.")
        return pd.DataFrame()

    headers = [re.sub(r'\s+', ' ', str(h)).strip() for h in data[0]]
    rows = data[1:]
    logging.info(f"  Loaded {len(rows)} rows from '{tab_name}'")

    max_cols = max(len(headers), 12, max(len(r) for r in rows))
    padded_rows = [r + [''] * (max_cols - len(r)) for r in rows]
    padded_headers = headers + [f"Unnamed_Col_{i}" for i in range(len(headers), max_cols)]
    df_gs = pd.DataFrame(padded_rows, columns=padded_headers)

    dark_post_col = next((c for c in df_gs.columns if 'dark post' in c.lower()), None)
    post_date_col = next((c for c in df_gs.columns if 'post date' in c.lower()), None)
    time_col = next((c for c in df_gs.columns if 'time' == c.lower()), None)
    id_col = next((c for c in df_gs.columns if 'id#' in c.lower() or 'id' == c.lower()), None)
    format_col = next((c for c in df_gs.columns if 'post format' in c.lower()), None)

    if not dark_post_col:
        logging.warning("Cannot find 'Dark Post' column in Judy's sheet")
        return pd.DataFrame()

    df_gs[dark_post_col] = df_gs[dark_post_col].fillna('').astype(str).str.strip().str.lower()
    df_gs = df_gs[df_gs[dark_post_col] == 'dark post']

    new_data = []
    for idx, row in df_gs.iterrows():
        post_date_str = str(row[post_date_col]).strip() if post_date_col else ""
        time_str = str(row[time_col]).strip() if time_col else ""
        post_id = str(row[id_col]).strip() if id_col else ""
        desc = str(row.iloc[11]).strip() if len(row) > 11 else ""

        raw_format = str(row[format_col]).strip().lower() if format_col else ""
        if raw_format in ['single', 'multi', 'singel']:
            mapped_post_type = 'Photo'
        elif raw_format == 'carousel':
            mapped_post_type = 'Link'
        elif raw_format == 'video':
            mapped_post_type = 'Video'
        else:
            mapped_post_type = 'Dark Post'

        if post_date_str:
            match = re.search(r'([A-Za-z]+)\s+(\d+)', post_date_str)
            if match:
                m_str, d_str = match.groups()
                parsed_dt = None
                try:
                    if time_str:
                        parsed_dt = pd.to_datetime(f"{m_str} {d_str} {target_year} {time_str}")
                    else:
                        parsed_dt = pd.to_datetime(f"{m_str} {d_str} {target_year}")
                    if not (parsed_dt.year == target_year and parsed_dt.month == target_month):
                        parsed_dt = None
                except:
                    parsed_dt = None
            else:
                parsed_dt = None
        else:
            parsed_dt = None

        if parsed_dt is None:
            continue

        if time_str:
            pub_time = f"{parsed_dt.month}/{parsed_dt.day}/{parsed_dt.year} {parsed_dt.hour}:{parsed_dt.minute:02d}"
            hour_val = parsed_dt.hour
        else:
            pub_time = f"{parsed_dt.month}/{parsed_dt.day}/{parsed_dt.year}"
            hour_val = ""

        permalink = f"https://www.facebook.com/manningshongkong/posts/{post_id}" if post_id else ""

        new_data.append({
            'Publish time': pub_time,
            'Month': parsed_dt.strftime('%b'),
            'Month No.': parsed_dt.month,
            'Day': parsed_dt.isoweekday(),
            'Hour': hour_val,
            'Dark Post': 'Y',
            'Permalink': permalink,
            'Description': desc,
            'Paid': 'Paid',
            'Organic': '',
            'Post type': mapped_post_type
        })

    logging.info(f"  Extracted {len(new_data)} Dark Posts from '{tab_name}'")
    return pd.DataFrame(new_data)
    


def categorize_promotions(row):
    """
    Updated classification rules (priority order, first match wins).
    """
    desc = str(row.get('Description', ''))
    desc_lower = desc.lower()
    pub_time_str = str(row.get('Publish time', ''))

    try:
        dt_obj = pd.to_datetime(pub_time_str)
        weekday = dt_obj.isoweekday()
        day_of_month = dt_obj.day
    except:
        weekday = -1
        day_of_month = -1

    pillar = ""
    category = ""
    sub_category = ""

    has_488_60 = ("488" in desc and "60" in desc)
    has_388_50 = ("388" in desc and "50" in desc)
    has_10_pct = ("10%" in desc)
    has_150_20 = ("150" in desc and "20" in desc)
    has_300_50 = ("300" in desc and "50" in desc)
    has_500_50 = ("500" in desc and "50" in desc)
    has_150_50 = ("150" in desc and "50" in desc)

    has_price_kw = any(kw in desc for kw in [
        "出位價", "優惠", "換購", "買滿", "即減", "優惠券", "折扣",
        "488", "388", "10%", "150", "300", "500"
    ])
    has_own_brand = any(kw in desc_lower for kw in ["mannings guardian", "萬寧 guardian", "萬寧自家"])
    has_exchange = ("換購優惠" in desc)
    has_keep_giving = ("萬寧多款優惠keep住送比你" in desc)
    has_wednesday = ("星期三" in desc)
    has_enjoy_card = ("enJoy卡客戶專享" in desc or "enjoy卡客戶專享" in desc_lower)
    has_yuu = ("yuu會員限定" in desc)
    has_yuu_dcv = ("兌換萬寧電子現金券慳最多12%" in desc)

    has_ecommerce = ("官網限定" in desc)
    has_newness = ("新登場" in desc)
    has_gnc = ("#GNC" in desc or "#gnc" in desc or "GNC" in desc)

    if has_ecommerce:
        pillar, category, sub_category = "Ecommerce", "Ecommerce", "Ecommerce"
    elif has_newness:
        pillar, category, sub_category = "Category", "Newness", "Supplier Innovation"
    elif has_gnc:
        pillar, category, sub_category = "GNC", "GNC", "GNC"
    elif has_yuu_dcv:
        pillar, category, sub_category = "CRM", "yuu DCV", ""
    elif has_yuu and day_of_month == 10:
        pillar, category, sub_category = "CRM", "yuu Day", ""
    elif (has_price_kw or has_488_60 or has_388_50 or has_10_pct or has_150_20 or
          has_300_50 or has_500_50 or has_150_50 or has_exchange or has_keep_giving) and (weekday == 5):
        pillar, category, sub_category = "Sales", "BAU Promotion", "Friday PNP"
    elif has_wednesday and (weekday in [2, 3]):
        pillar, category, sub_category = "Sales", "BAU Promotion", "Happy Wednesday"
    elif has_own_brand:
        pillar, category, sub_category = "Sales", "Own Brand", "Own Brand"
    elif has_enjoy_card and (day_of_month in [1, 20]):
        pillar, category, sub_category = "Sales", "Payment", "Enjoy Card Day"

    return pd.Series([pillar, category, sub_category])


def filter_to_target_period(df, target_year, target_month):
    """Filter DataFrame to the target year-month using HK Time DT (preferred) or Publish time."""
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


def process_facebook_data():
    logging.info(f"=== 1a_period: Facebook Post Processing ({TARGET_YEAR}-{TARGET_MONTH:02d}) ===")

    logging.info("Step 1: Merging all raw CSVs from FB raw/ ...")
    df = merge_csv_files(RAW_FB_DIR)
    if df.empty:
        logging.error("No FB raw data found.")
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
        logging.warning(f"No FB posts found for {TARGET_YEAR}-{TARGET_MONTH:02d}.")
        return None

    logging.info("Step 6: Processing text and metrics ...")
    if 'Title' not in df.columns:
        df['Title'] = ""
    if 'Description' not in df.columns:
        df['Description'] = ""

    def get_longer_text(row):
        t = str(row['Title']).strip() if pd.notna(row['Title']) else ""
        d = str(row['Description']).strip() if pd.notna(row['Description']) else ""
        if t.lower() == 'nan': t = ""
        if d.lower() == 'nan': d = ""
        if len(t) > len(d): return t
        elif len(d) > len(t): return d
        else: return t

    df['Description'] = df.apply(get_longer_text, axis=1)
    df.drop(columns=['Title'], inplace=True, errors='ignore')

    df.rename(columns={
        'Reach from Organic posts': 'Organic reach',
        'Reach from Boosted posts': 'Paid reach',
        'Matched Audience Targeting Consumption (Photo Click)': 'Photo Click'
    }, inplace=True)

    for col in ['Reactions', 'Comments', 'Shares']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        else:
            df[col] = 0
    df['Interactions'] = df['Reactions'] + df['Comments'] + df['Shares']

    if 'Paid reach' in df.columns:
        df['Paid reach'] = pd.to_numeric(df['Paid reach'], errors='coerce').fillna(0)
        df['Organic'] = df['Paid reach'].apply(lambda x: 'Organic' if x == 0 else '')
        df['Paid'] = df['Paid reach'].apply(lambda x: 'Paid' if x != 0 else '')
    else:
        df['Paid reach'] = 0
        df['Organic'], df['Paid'] = 'Organic', ''

    df = add_time_columns(df, 'HK Time DT')

    if 'Average Seconds viewed' in df.columns:
        df['Average Watch Time'] = pd.to_numeric(df['Average Seconds viewed'], errors='coerce').round(2).fillna("")
    else:
        df['Average Watch Time'] = ""

    if 'Duration (sec)' in df.columns:
        df['Video Length'] = pd.to_numeric(df['Duration (sec)'], errors='coerce').fillna("")
    else:
        df['Video Length'] = ""

    df.sort_values(by='HK Time DT', ascending=True, na_position='last', inplace=True)
    df.drop(columns=['HK Time DT'], inplace=True, errors='ignore')

    logging.info(f"Step 7: Merging Dark Posts from Judy's Sheet for {TARGET_YEAR}-{TARGET_MONTH:02d} ...")
    df_gsheet_dark_posts = fetch_gsheet_dark_posts(TARGET_YEAR, TARGET_MONTH)
    if not df_gsheet_dark_posts.empty:
        df = pd.concat([df, df_gsheet_dark_posts], ignore_index=True)
        logging.info(f"  Merged dark posts: total {len(df)} rows")

    logging.info("Step 8: Applying updated classification rules ...")
    df[['Pillar', 'Category', 'Sub-Category']] = df.apply(categorize_promotions, axis=1)
    df = normalize_columns(df)

    final_columns = [
        'Month', 'Permalink', 'Description', 'Post type', 'Pillar', 'Category', 'Sub-Category', 'Campaign Name',
        'Publish time', 'Reactions', 'Comments', 'Shares', 'Interactions', 'Reach',
        'Views', 'Link Clicks', 'Photo Click', '3-second video views', 'Video Length',
        'Organic reach', 'Paid reach', 'Dark Post', 'Organic', 'Paid', 'Month No.',
        'Day', 'Hour', 'Average Watch Time'
    ]
    for col in final_columns:
        if col not in df.columns:
            df[col] = ""
            
    df = df[final_columns]
    # CRITICAL: Clean the data to prevent IllegalCharacterError
    df = deep_clean_df(df)
    
    # Proceed to export
    logging.info(f"Step 9: Output -> {OUTPUT_FILE}")
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="FB_Posts")
        
        from openpyxl.styles import Alignment
        # Fix: Fetching directly from the workbook ensures it uses openpyxl syntax
        worksheet = writer.book["FB_Posts"]
        
        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=False)

if __name__ == "__main__":
    process_facebook_data()
