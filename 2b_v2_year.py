"""
2b_v2_year.py - Comment Data Cleaning + Chart Generation (ALL data, Excel Output)
Based on 2b_chart_period_v2.py with:
  - Reads/Writes .xlsx (ALL data, no period filter)
  - deep_clean_df for Excel-safe characters
  - Adds category fuzzy matching from reference script
  - Generates 3 chart PNGs (Overall Sentiment, Category Dashboard, Type Dashboard)

Input:  Mannings_Comments_RAW_ALL.xlsx
Output: FB Comments - ALL.xlsx
        Mannings_00_Overall_Sentiment_ALL.png
        Mannings_01_Category_Dashboard_ALL.png
        Mannings_02_Type_Dashboard_ALL.png
"""
import warnings
import pandas as pd
import logging
import re
import os
import sys
import difflib
import textwrap
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import seaborn as sns

warnings.filterwarnings('ignore', category=pd.errors.PerformanceWarning)

if sys.platform == 'win32':
    import ctypes
    kernel32 = ctypes.windll.kernel32
    kernel32.SetConsoleOutputCP(65001)
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')

from config import BASE_DIR
from utils import clean_sentiment_data, clean_category_data

# ==============================================================================
# CONFIGURATION
# ==============================================================================
TARGET_YEAR = 2026

INPUT_XLSX = BASE_DIR / f"Mannings_Comments_RAW_{TARGET_YEAR}.xlsx"
OUTPUT_XLSX = BASE_DIR / f"FB Comments - {TARGET_YEAR}.xlsx"

"""
OUTPUT_OVERALL_IMG = BASE_DIR / f"Mannings_00_Overall_Sentiment_{TARGET_YEAR}.png"
OUTPUT_CATEGORY_IMG = BASE_DIR / f"Mannings_01_Category_Dashboard_{TARGET_YEAR}.png"
OUTPUT_TYPE_IMG = BASE_DIR / f"Mannings_02_Type_Dashboard_{TARGET_YEAR}.png"

COLOR_MAP = {
    'Positive': '#43A047',
    'Neutral': '#674eaa',
    'Negative': '#E53935'
}
"""

# ==============================================================================
# EXCEL-SAFE CLEANING
# ==============================================================================
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

# ==============================================================================
# DATA CLEANING
# ==============================================================================
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
    logging.info(f"=== 2b_v2_year: Comment Cleaning + Charts ({TARGET_YEAR}) ===")

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

"""
# ==============================================================================
# CHART GENERATION
# ==============================================================================
def set_font_and_style():
    plt.rcParams['font.sans-serif'] = ['Arial', 'Helvetica', 'sans-serif']
    plt.rcParams['axes.unicode_minus'] = False
    plt.rcParams['axes.facecolor'] = '#F4F6F9'
    plt.rcParams['figure.facecolor'] = '#F4F6F9'
    plt.rcParams['axes.edgecolor'] = '#E0E0E0'

def draw_overall_dashboard(df, output_path):
    logging.info("Generating Overall Sentiment Chart ...")
    df_plot = df.dropna(subset=['Sentiment']).copy()
    df_plot = df_plot[df_plot['Sentiment'].isin(['Positive', 'Neutral', 'Negative'])]
    if df_plot.empty:
        return

    fig, ax_pie = plt.subplots(figsize=(10, 8))
    sent_counts = df_plot['Sentiment'].value_counts()
    colors = [COLOR_MAP.get(s, '#999999') for s in sent_counts.index]

    def make_autopct(values):
        def my_autopct(pct):
            total = sum(values)
            val = int(round(pct * total / 100.0))
            return f'{val} ({pct:.1f}%)'
        return my_autopct

    wedges, texts, autotexts = ax_pie.pie(
        sent_counts, labels=None, autopct=make_autopct(sent_counts), pctdistance=0.70,
        startangle=140, colors=colors,
        wedgeprops=dict(width=0.4, edgecolor='#F4F6F9', linewidth=3)
    )

    for autotext in autotexts:
        autotext.set_fontsize(10)
        autotext.set_fontweight('bold')
        autotext.set_color('#1C2833')

    ax_pie.text(0, 0.08, f"{len(df_plot)}", ha='center', va='center', fontsize=26, fontweight='900', color='#1C2833')
    ax_pie.text(0, -0.12, "Total Comments", ha='center', va='center', fontsize=10, color='#7F8C8D', fontweight='bold')

    ax_pie.legend(wedges, sent_counts.index, loc="center left", bbox_to_anchor=(1.0, 0.5), frameon=False, fontsize=11)

    plt.savefig(output_path, dpi=300, bbox_inches='tight', facecolor='#F4F6F9')
    plt.close()

def draw_category_dashboard(df, output_path):
    logging.info("Generating Category Dashboard ...")

    df_plot = df.dropna(subset=['Category', 'Sentiment']).copy()
    df_plot = df_plot[~df_plot['Category'].isin(['Nan', '', 'None', 'Null'])]
    df_plot = df_plot[df_plot['Sentiment'].isin(['Positive', 'Neutral', 'Negative'])]
    if df_plot.empty:
        return

    cat_order = df_plot['Category'].value_counts().sort_values(ascending=False).index.tolist()
    N = len(cat_order)

    ct = pd.crosstab(df_plot['Sentiment'], df_plot['Category'])
    for idx in ['Positive', 'Neutral', 'Negative']:
        if idx not in ct.index:
            ct.loc[idx] = 0
    ct = ct.reindex(['Positive', 'Neutral', 'Negative'])[cat_order]

    ct['Grand Total'] = ct.sum(axis=1)
    ct.loc['Grand Total'] = ct.sum(axis=0)

    fig = plt.figure(figsize=(24, 11))

    w_sentiment = 3.2
    w_category = 2.2
    w_grand_total = 2.5
    total_units = w_sentiment + (N * w_category) + w_grand_total

    left_base = 0.03
    right_base = 0.97
    total_span = right_base - left_base

    col_widths = [w_sentiment / total_units] + [w_category / total_units] * N + [w_grand_total / total_units]

    chart_left = left_base + (w_sentiment / total_units) * total_span
    chart_right = right_base - (w_grand_total / total_units) * total_span
    chart_width = chart_right - chart_left

    ax_chart = fig.add_axes([chart_left, 0.40, chart_width, 0.50])
    chart_data = pd.crosstab(df_plot['Category'], df_plot['Sentiment']).reindex(cat_order, fill_value=0)
    bar_cols = [s for s in ['Negative', 'Neutral', 'Positive'] if s in chart_data.columns]

    chart_data[bar_cols].plot(kind='bar', stacked=True, ax=ax_chart,
                              color=[COLOR_MAP.get(c) for c in bar_cols], width=0.35, edgecolor='none')

    ax_chart.set_ylabel('Number of Comments (Count)', fontsize=11, fontweight='bold', color='#7F8C8D')
    ax_chart.set_xlabel('')

    ax_chart.set_xticklabels([])
    ax_chart.set_xticks([])
    ax_chart.set_xlim(-0.5, N - 0.5)
    ax_chart.set_xmargin(0)
    sns.despine(ax=ax_chart, top=True, right=True, bottom=True, left=False)
    ax_chart.grid(axis='y', linestyle=':', alpha=0.5, color='#BDC3C7')

    handles, labels = ax_chart.get_legend_handles_labels()
    ax_chart.legend(handles[::-1], labels[::-1], loc='upper right', frameon=False, fontsize=11)

    ax_table = fig.add_axes([left_base, 0.06, total_span, 0.32])
    ax_table.axis('off')

    col_labels = ['Sentiment'] + list(ct.columns)
    cells = []
    for idx, row in ct.iterrows():
        cells.append([idx] + [f'{int(v)}' for v in row])

    m_table = ax_table.table(cellText=cells, colLabels=col_labels, colWidths=col_widths, loc='center', cellLoc='center')
    m_table.auto_set_font_size(False)
    m_table.set_fontsize(11)
    m_table.scale(1, 2.5)

    for (row, col), cell in m_table.get_celld().items():
        cell.set_linewidth(0.5)
        cell.set_edgecolor('#D3D3D3')

        if row == 0:
            cell.set_text_props(weight='bold', color='#000000', size=11)
            cell.set_facecolor('#D6E4F0')
        elif row == len(cells):
            cell.set_text_props(weight='bold', color='#000000', size=11)
            cell.set_facecolor('#B9D1EA')
        else:
            cell.set_facecolor('#F2F7FC')
            if col == 0:
                cell.set_text_props(weight='bold', color='#1C2833')

    plt.savefig(output_path, dpi=300, bbox_inches='tight', facecolor='#F4F6F9')
    plt.close()

def draw_type_dashboard(df, output_path):
    logging.info("Generating Type Dashboard ...")

    df_plot = df.dropna(subset=['Type', 'Sentiment']).copy()
    df_plot = df_plot[~df_plot['Type'].isin(['Nan', '', 'None', 'Null'])]
    df_plot = df_plot[df_plot['Sentiment'].isin(['Positive', 'Neutral', 'Negative'])]
    if df_plot.empty:
        return

    type_order = df_plot['Type'].value_counts().sort_values(ascending=False).index.tolist()
    N = len(type_order)

    ct = pd.crosstab(df_plot['Sentiment'], df_plot['Type'])
    for idx in ['Positive', 'Neutral', 'Negative']:
        if idx not in ct.index:
            ct.loc[idx] = 0
    ct = ct.reindex(['Positive', 'Neutral', 'Negative'])[type_order]

    ct['Grand Total'] = ct.sum(axis=1)
    ct.loc['Grand Total'] = ct.sum(axis=0)

    dynamic_width = max(24, N * 1.6)
    fig = plt.figure(figsize=(dynamic_width, 12))

    w_sentiment = 3.2
    w_type = 2.6
    w_grand_total = 3.0
    total_units = w_sentiment + (N * w_type) + w_grand_total

    left_base = 0.03
    right_base = 0.97
    total_span = right_base - left_base

    col_widths = [w_sentiment / total_units] + [w_type / total_units] * N + [w_grand_total / total_units]

    chart_left = left_base + (w_sentiment / total_units) * total_span
    chart_right = right_base - (w_grand_total / total_units) * total_span
    chart_width = chart_right - chart_left

    ax_chart = fig.add_axes([chart_left, 0.44, chart_width, 0.48])
    chart_data = pd.crosstab(df_plot['Type'], df_plot['Sentiment']).reindex(type_order, fill_value=0)
    bar_cols = [s for s in ['Negative', 'Neutral', 'Positive'] if s in chart_data.columns]

    chart_data[bar_cols].plot(kind='bar', stacked=True, ax=ax_chart,
                              color=[COLOR_MAP.get(c) for c in bar_cols], width=0.35, edgecolor='none')

    ax_chart.set_ylabel('Number of Comments (Count)', fontsize=11, fontweight='bold', color='#7F8C8D')
    ax_chart.set_xlabel('')

    ax_chart.set_xticklabels([])
    ax_chart.set_xticks([])
    ax_chart.set_xlim(-0.5, N - 0.5)
    ax_chart.set_xmargin(0)
    sns.despine(ax=ax_chart, top=True, right=True, bottom=True, left=False)
    ax_chart.grid(axis='y', linestyle=':', alpha=0.5, color='#BDC3C7')

    handles, labels = ax_chart.get_legend_handles_labels()
    ax_chart.legend(handles[::-1], labels[::-1], loc='upper right', frameon=False, fontsize=11)

    ax_table = fig.add_axes([left_base, 0.05, total_span, 0.35])
    ax_table.axis('off')

    wrapped_headers = [textwrap.fill(str(col), width=16) for col in ct.columns]
    col_labels = ['Sentiment'] + wrapped_headers

    cells = []
    for idx, row in ct.iterrows():
        cells.append([idx] + [f'{int(v)}' for v in row])

    m_table = ax_table.table(cellText=cells, colLabels=col_labels, colWidths=col_widths, loc='center', cellLoc='center')
    m_table.auto_set_font_size(False)
    m_table.set_fontsize(9.5)
    m_table.scale(1, 3.2)

    for (row, col), cell in m_table.get_celld().items():
        cell.set_linewidth(0.5)
        cell.set_edgecolor('#D3D3D3')

        if row == 0:
            cell.set_text_props(weight='bold', color='#000000', size=9.5)
            cell.set_facecolor('#D6E4F0')
        elif row == len(cells):
            cell.set_text_props(weight='bold', color='#000000', size=10)
            cell.set_facecolor('#B9D1EA')
        else:
            cell.set_facecolor('#F2F7FC')
            if col == 0:
                cell.set_text_props(weight='bold', color='#1C2833')

    plt.savefig(output_path, dpi=300, bbox_inches='tight', facecolor='#F4F6F9')
    plt.close()
"""

# ==============================================================================
# MAIN
# ==============================================================================
if __name__ == "__main__":
    
    ### set_font_and_style() ###

    if not os.path.exists(INPUT_XLSX):
        logging.error(f"Input not found: {INPUT_XLSX}. Run 2a_new_v2.py first.")
    else:
        df = pd.read_excel(INPUT_XLSX)
        df = clean_comment_data(df)
        df = deep_clean_df(df)

        # --- Write cleaned xlsx ---
        logging.info(f"  Output -> {OUTPUT_XLSX}")
        with pd.ExcelWriter(OUTPUT_XLSX, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Comments_Cleaned")
            from openpyxl.styles import Alignment
            worksheet = writer.book["Comments_Cleaned"]
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=False)

        logging.info(f"Cleaned xlsx: {len(df)} comments (ALL periods).")

        """
        # --- Generate charts ---
        draw_overall_dashboard(df, OUTPUT_OVERALL_IMG)
        logging.info(f"  Chart -> {OUTPUT_OVERALL_IMG}")

        if 'Category' in df.columns:
            draw_category_dashboard(df, OUTPUT_CATEGORY_IMG)
            logging.info(f"  Chart -> {OUTPUT_CATEGORY_IMG}")

        if 'Type' in df.columns:
            draw_type_dashboard(df, OUTPUT_TYPE_IMG)
            logging.info(f"  Chart -> {OUTPUT_TYPE_IMG}")
        """

        logging.info("DONE!")
