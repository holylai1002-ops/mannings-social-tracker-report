"""
2b_new_period_v2.py - Comment Data Cleaning (Target Period, Excel Output)
Based on 2b_chart_period_v2.py with:
  - Reads/Writes .xlsx instead of .csv
  - deep_clean_df for Excel-safe characters
  - Adds category fuzzy matching from 2b_chart_period_v2.py
  - No chart generation
"""
import warnings
import pandas as pd
import logging
import re
import os
import difflib
warnings.filterwarnings('ignore', category=pd.errors.PerformanceWarning)

from config import BASE_DIR
from utils import clean_sentiment_data, clean_category_data

TARGET_YEAR = 2026
TARGET_MONTH = 6

INPUT_XLSX = BASE_DIR / f"Mannings_Comments_RAW_{TARGET_YEAR}_{TARGET_MONTH:02d}.xlsx"
OUTPUT_XLSX = BASE_DIR / f"FB Comments - {TARGET_YEAR}_{TARGET_MONTH:02d}.xlsx"

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

PROTECTED_CATEGORIES = {
    'Compliment', 'Complaint', 'Enquiry', 'Enquriy', 'Others',
    'Website/App Experience', 'Delivery Issues', 'Customer Service',
    'Product Condition', 'Spam'
}

def fuzzy_clean_category(df):
    """Extra category fuzzy matching from reference script.
    Matches rare categories (count <= 3) against common ones (count > 3).
    PROTECTED_CATEGORIES are never fuzzy-matched (prevents Compliment→Complaint etc)."""
    if 'Category' not in df.columns:
        return df
    cat_counts = df['Category'].astype(str).str.strip().value_counts()
    std_cats = cat_counts[cat_counts > 3].index.tolist()

    def fuzzy_category(val):
        val = str(val).strip()
        if val in std_cats or val in PROTECTED_CATEGORIES or val in ['Nan', '', 'None', 'Null']:
            return val
        matches = difflib.get_close_matches(val, std_cats, n=1, cutoff=0.8)
        return matches[0] if matches else val

    df['Category'] = df['Category'].apply(fuzzy_category)
    return df

def clean_comment_data(df):
    logging.info(f"=== 2b_new_period_v2: Comment Cleaning ({TARGET_YEAR}-{TARGET_MONTH:02d}) ===")

    cols_rename = {}
    sentiment_cols = [col for col in df.columns if 'sentiment' in col.lower()]
    if sentiment_cols:
        cols_rename[sentiment_cols[0]] = 'Sentiment'
    category_cols = [col for col in df.columns if 'catagory' in col.lower() or 'category' in col.lower()]
    if category_cols:
        cols_rename[category_cols[0]] = 'Category'
    type_cols = [col for col in df.columns if col.lower() == 'type']
    if type_cols:
        cols_rename[type_cols[0]] = 'Type'
    df = df.rename(columns=cols_rename)

    df = clean_sentiment_data(df)
    df = clean_category_data(df)
    df = fuzzy_clean_category(df)

    if 'Type' in df.columns:
        df['Type'] = df['Type'].astype(str).str.strip().str.title()
        type_mapping = {'Promtion': 'Promotion', 'Enqury': 'Enquiry'}
        df['Type'] = df['Type'].replace(type_mapping)

    missing_report = []
    for field in ['Sentiment', 'Category', 'Type']:
        if field in df.columns:
            missing_mask = df[field].isna() | (df[field].astype(str).str.strip() == '') | (df[field].astype(str).str.lower() == 'nan')
            missing_count = missing_mask.sum()
            if missing_count > 0:
                missing_report.append(f"  {field}: {missing_count} blank values")

    if missing_report:
        logging.warning("Data audit - blank fields found:")
        for msg in missing_report:
            logging.warning(msg)
    else:
        logging.info("Data audit passed: all critical fields populated.")

    return df

if __name__ == "__main__":
    if not os.path.exists(INPUT_XLSX):
        logging.error(f"Input not found: {INPUT_XLSX}. Run 2a_new_period_v2.py first.")
    else:
        df = pd.read_excel(INPUT_XLSX)
        df = clean_comment_data(df)
        df = deep_clean_df(df)

        logging.info(f"  Output -> {OUTPUT_XLSX}")
        with pd.ExcelWriter(OUTPUT_XLSX, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Comments_Cleaned")
            from openpyxl.styles import Alignment
            worksheet = writer.book["Comments_Cleaned"]
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=False)

        logging.info(f"DONE! {len(df)} comments cleaned for {TARGET_YEAR}-{TARGET_MONTH:02d}.")
