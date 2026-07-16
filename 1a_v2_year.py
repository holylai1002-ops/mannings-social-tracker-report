"""
1a_v2_year.py - Facebook Post Processing (ALL data, Excel Output)
Based on 1a_new.py with:
  - Outputs .xlsx instead of .csv
  - deep_clean_df for Excel-safe characters
  - No text wrap on cells
  - filter to target year only
"""
import warnings
import pandas as pd
import re
import logging
warnings.filterwarnings('ignore', category=pd.errors.PerformanceWarning)

from config import RAW_FB_DIR, BASE_DIR, GSHEET_JUDY_ID
from utils import (
    merge_csv_files, dedupe_by_id, convert_tz_la_to_hk,
    filter_valid_years, add_time_columns, normalize_columns,
    get_gspread_client
)

TARGET_YEAR = 2026

OUTPUT_FILE = BASE_DIR / f"All FB Posts_New Cat_{TARGET_YEAR}.xlsx"

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

def fetch_gsheet_dark_posts():
    try:
        client = get_gspread_client()
        sh = client.open_by_key(GSHEET_JUDY_ID)
    except Exception as e:
        logging.warning(f"Cannot connect to Judy's Google Sheet for Dark Posts: {e}")
        return pd.DataFrame()

    all_rows = []
    headers = []

    for ws in sh.worksheets():
        tab_name = ws.title
        if str(TARGET_YEAR) not in tab_name:
            continue
        try:
            data = ws.get_all_values()
            if not data or len(data) < 2:
                continue

            if not headers:
                headers = [re.sub(r'\s+', ' ', str(h)).strip() for h in data[0]]

            all_rows.extend(data[1:])
            logging.info(f"  Judy Sheet tab '{tab_name}': {len(data)-1} rows loaded")
        except Exception as e:
            logging.warning(f"  Failed reading tab '{tab_name}': {e}")

    if not all_rows:
        return pd.DataFrame()

    actual_max_row_len = max([len(row) for row in all_rows]) if all_rows else 0
    max_cols = max(len(headers), 12, actual_max_row_len)

    padded_rows = [row + [''] * (max_cols - len(row)) for row in all_rows]
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

        match = re.search(r'([A-Za-z]+)\s+(\d+)', post_date_str)
        if not match:
            continue

        m_str, d_str = match.groups()
        year_guess = TARGET_YEAR
        try:
            if time_str:
                dt_obj = pd.to_datetime(f"{m_str} {d_str} {year_guess} {time_str}")
            else:
                dt_obj = pd.to_datetime(f"{m_str} {d_str} {year_guess}")
        except:
            try:
                dt_obj = pd.to_datetime(f"{m_str} {d_str} {year_guess-1} {time_str}")
            except:
                continue

        if time_str:
            pub_time = f"{dt_obj.month}/{dt_obj.day}/{dt_obj.year} {dt_obj.hour}:{dt_obj.minute:02d}"
            hour_val = dt_obj.hour
        else:
            pub_time = f"{dt_obj.month}/{dt_obj.day}/{dt_obj.year}"
            hour_val = ""

        permalink = f"https://www.facebook.com/manningshongkong/posts/{post_id}" if post_id else ""

        new_data.append({
            'Publish time': pub_time,
            'Month': dt_obj.strftime('%b'),
            'Month No.': dt_obj.month,
            'Day': dt_obj.isoweekday(),
            'Hour': hour_val,
            'Dark Post': 'Y',
            'Permalink': permalink,
            'Description': desc,
            'Paid': 'Paid',
            'Organic': '',
            'Post type': mapped_post_type
        })

    logging.info(f"  Extracted {len(new_data)} Dark Posts from Judy's Sheet")
    return pd.DataFrame(new_data)

def categorize_promotions(row):
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
    has_price_kw = any(kw in desc for kw in ["出位價", "優惠", "換購", "買滿", "即減", "優惠券", "折扣", "488", "388", "10%", "150", "300", "500"])
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
    elif (has_price_kw or has_488_60 or has_388_50 or has_10_pct or has_150_20 or has_300_50 or has_500_50 or has_150_50 or has_exchange or has_keep_giving) and (weekday == 5):
        pillar, category, sub_category = "Sales", "BAU Promotion", "Friday PNP"
    elif has_wednesday and (weekday in [2, 3]):
        pillar, category, sub_category = "Sales", "BAU Promotion", "Happy Wednesday"
    elif has_own_brand:
        pillar, category, sub_category = "Sales", "Own Brand", "Own Brand"
    elif has_enjoy_card and (day_of_month in [1, 20]):
        pillar, category, sub_category = "Sales", "Payment", "Enjoy Card Day"
    return pd.Series([pillar, category, sub_category])

def process_facebook_data():
    logging.info("=== 1a_v2_year: Facebook Post Processing (ALL periods) ===")

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

    logging.info("Step 5: Processing text and metrics ...")
    if 'Title' not in df.columns:
        df['Title'] = ""
    if 'Description' not in df.columns:
        df['Description'] = ""
    def get_longer_text(row):
        t = str(row['Title']).strip() if pd.notna(row['Title']) else ""
        d = str(row['Description']).strip() if pd.notna(row['Description']) else ""
        if t.lower() == 'nan': t = ""
        if d.lower() == 'nan': d = ""
        return t if len(t) > len(d) else d if len(d) > len(t) else t
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

    logging.info("Step 6: Merging Dark Posts from Judy's Google Sheet ...")
    df_gsheet_dark_posts = fetch_gsheet_dark_posts()
    if not df_gsheet_dark_posts.empty:
        df = pd.concat([df, df_gsheet_dark_posts], ignore_index=True)

    logging.info("Step 7: Applying updated classification rules ...")
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

    df = deep_clean_df(df)

    logging.info(f"Step 8: Output -> {OUTPUT_FILE}")
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="FB_Posts")
        from openpyxl.styles import Alignment
        worksheet = writer.book["FB_Posts"]
        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=False)
    logging.info(f"DONE! {len(df)} rows processed.")
    return df

if __name__ == "__main__":
    process_facebook_data()
