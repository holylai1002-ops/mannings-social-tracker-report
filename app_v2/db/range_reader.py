"""
Range-based Excel reader for app_v2.

Loads Mannings_FB_IG_Dashboard_Feed.xlsx and filters by arbitrary date ranges
(start_date, end_date) instead of fixed year/month periods.
"""

import time
import logging
from datetime import date, timedelta
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from app_v2.config import settings

logger = logging.getLogger(__name__)

MONTH_NAMES = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December",
}

MONTH_MAP = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
    "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
}

# Sheets that have individual post rows with a date column
POST_LEVEL_SHEETS = {
    "FB Wall Post Performance": "Publish time",
    "Category Performance - BAU": "Publish time",
    "Sub-Category Performance - PNP": "Publish time",
    "Pillar Performance - CRM": "Publish time",
    "Pillar Performance - Ecommerce": "Publish time",
    "Pillar Performance - GNC": "Publish time",
    "Pillar Performance - Branding": "Publish time",
    "Pillar Performance - Category": "Publish time",
    "Pillar Performance - Sales": "Publish time",
    "Pillar Performance - Others": "Publish time",
    "IG Wall Post Performance": "Post Date",
    "IG Story Performance": "Publish time",
}

DAILY_SHEETS = {"Unique Page View", "FB Followers", "IG Followers"}

PIVOT_SHEETS = {
    "FB Pivot (Category)": ("FB Wall Post Performance", "Category"),
    "FB Pivot (Post Type)": ("FB Wall Post Performance", "Post type"),
    "FB Pivot (Pillar)": ("FB Wall Post Performance", "Pillar"),
    "IG Pivot (Pillar)": ("IG Wall Post Performance", "Pillar"),
    "IG Story Pivot": ("IG Story Performance", "Pillar"),
}


# ==============================================================================
# RangeData
# ==============================================================================
@dataclass
class RangeData:
    start: str
    end: str
    label: str
    period_str: str
    sheets: dict = field(default_factory=dict)
    _kpis_cache: dict = field(default_factory=dict)

    def get(self, name: str) -> pd.DataFrame:
        return self.sheets.get(name, pd.DataFrame())

    def kpis(self) -> dict:
        if self._kpis_cache:
            return self._kpis_cache

        fb_followers_df = self.get("FB Followers")
        ig_followers_df = self.get("IG Followers")
        fb_wall = self.get("FB Wall Post Performance")
        ig_wall = self.get("IG Wall Post Performance")

        def _safe_sum(df, col):
            if df.empty or col not in df.columns:
                return 0
            return int(pd.to_numeric(df[col], errors="coerce").fillna(0).sum())

        fb_followers = 0
        fb_growth = 0
        if not fb_followers_df.empty and "Total Followers" in fb_followers_df.columns:
            fb_followers = int(pd.to_numeric(fb_followers_df["Total Followers"], errors="coerce").fillna(0).iloc[0])
        if not fb_followers_df.empty and "Followers Net" in fb_followers_df.columns:
            fb_growth = int(pd.to_numeric(fb_followers_df["Followers Net"], errors="coerce").fillna(0).sum())

        ig_followers = 0
        ig_growth = 0
        if not ig_followers_df.empty and "Total Followers" in ig_followers_df.columns:
            ig_followers = int(pd.to_numeric(ig_followers_df["Total Followers"], errors="coerce").fillna(0).iloc[0])
        if not ig_followers_df.empty and "Followers Net" in ig_followers_df.columns:
            ig_growth = int(pd.to_numeric(ig_followers_df["Followers Net"], errors="coerce").fillna(0).sum())

        result = {
            "fb_followers": fb_followers,
            "fb_growth": fb_growth,
            "fb_wall_posts": len(fb_wall),
            "fb_interactions": _safe_sum(fb_wall, "Interactions"),
            "ig_followers": ig_followers,
            "ig_growth": ig_growth,
            "ig_reach": _safe_sum(ig_wall, "Total Post Reach"),
            "ig_interactions": _safe_sum(ig_wall, "Total Interactions"),
        }
        self._kpis_cache = result
        return result


# ==============================================================================
# Excel cache
# ==============================================================================
_excel_cache: dict[str, tuple[float, pd.ExcelFile]] = {}


def _get_excel_file(path: str) -> pd.ExcelFile:
    mtime = Path(path).stat().st_mtime
    cached = _excel_cache.get(path)
    if cached and cached[0] == mtime:
        return cached[1]
    logger.info(f"Loading Excel: {path}")
    xl = pd.ExcelFile(path)
    _excel_cache[path] = (mtime, xl)
    return xl


# ==============================================================================
# Date helpers
# ==============================================================================
def _parse_date_col(series: pd.Series) -> pd.Series:
    """Parse a date column that may be DD/MM/YYYY or DD/MM/YYYY HH:MM."""
    s = series.astype(str).str.strip()
    dt = pd.to_datetime(s, format="%d/%m/%Y %H:%M", errors="coerce")
    dt = dt.fillna(pd.to_datetime(s, format="%d/%m/%Y", errors="coerce"))
    return dt


def _filter_by_date_range(df: pd.DataFrame, date_col: str,
                          start: pd.Timestamp, end: pd.Timestamp) -> pd.DataFrame:
    if df.empty or date_col not in df.columns:
        return pd.DataFrame()

    # Prefer integer _Year/_Month/_Day columns (no date parsing ambiguity)
    if all(c in df.columns for c in ('_Year', '_Month', '_Day')):
        y = pd.to_numeric(df['_Year'], errors='coerce').fillna(0).astype(int)
        m = pd.to_numeric(df['_Month'], errors='coerce').fillna(0).astype(int)
        d = pd.to_numeric(df['_Day'], errors='coerce').fillna(0).astype(int)
        ymd = y * 10000 + m * 100 + d
        start_ymd = start.year * 10000 + start.month * 100 + start.day
        end_ymd = end.year * 10000 + end.month * 100 + end.day
        mask = (ymd >= start_ymd) & (ymd <= end_ymd) & (ymd > 0)
        return df[mask].copy()

    # Fallback: parse date strings
    dt = _parse_date_col(df[date_col])
    mask = (dt >= start) & (dt <= end)
    return df[mask].copy()


def _filter_daily(df: pd.DataFrame, start: pd.Timestamp, end: pd.Timestamp) -> pd.DataFrame:
    if df.empty or "Date" not in df.columns:
        return pd.DataFrame()
    return _filter_by_date_range(df, "Date", start, end)


def _filter_reach_funnel(df: pd.DataFrame, start: pd.Timestamp, end: pd.Timestamp) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    months_in_range = set()
    cur = start.replace(day=1)
    while cur <= end:
        months_in_range.add((cur.year, cur.month))
        if cur.month == 12:
            cur = cur.replace(year=cur.year + 1, month=1)
        else:
            cur = cur.replace(month=cur.month + 1)
    if "Year" in df.columns and "Month No." in df.columns:
        yr = pd.to_numeric(df["Year"], errors="coerce")
        mo = pd.to_numeric(df["Month No."], errors="coerce")
        mask = df.apply(lambda r: (int(r["Year"]), int(r["Month No."])) in months_in_range
                        if pd.notna(r["Year"]) and pd.notna(r["Month No."]) else False, axis=1)
        return df[mask].copy()
    return df


def _reaggregate_pivot(df_posts: pd.DataFrame, group_col: str) -> pd.DataFrame:
    """Re-aggregate a pivot table from filtered post-level data."""
    if df_posts.empty or group_col not in df_posts.columns:
        return pd.DataFrame(columns=["Period", group_col, "No. of Posts", "Total Interactions"])

    df_posts = df_posts.copy()

    # Detect interaction column (different sheets use different names)
    int_col = None
    for candidate in ["Interactions", "Total Interactions"]:
        if candidate in df_posts.columns:
            int_col = candidate
            break

    # Count posts
    result = df_posts.groupby(group_col).agg(
        **{"No. of Posts": (group_col, "count")},
    ).reset_index()

    # Aggregate interaction column
    if int_col:
        df_posts["_int_num"] = pd.to_numeric(df_posts[int_col].astype(str).str.replace(",", ""), errors="coerce").fillna(0)
        agg = df_posts.groupby(group_col)["_int_num"].sum()
        result["Total Interactions"] = result[group_col].map(agg).fillna(0).astype(int)
    else:
        result["Total Interactions"] = 0

    # Aggregate extra numeric columns (e.g. Link clicks, Total Reach)
    for extra_col in ["Link clicks", "Total Reach", "Total Post Reach"]:
        if extra_col in df_posts.columns:
            df_posts["_extra_num"] = pd.to_numeric(df_posts[extra_col].astype(str).str.replace(",", ""), errors="coerce").fillna(0)
            agg = df_posts.groupby(group_col)["_extra_num"].sum()
            out_name = f"SUM of {extra_col}" if extra_col == "Link clicks" else extra_col
            result[out_name] = result[group_col].map(agg).fillna(0).astype(int)

    result.insert(0, "Period", "Range")
    return result


def _filter_comments(df: pd.DataFrame, start: pd.Timestamp, end: pd.Timestamp) -> pd.DataFrame:
    if df.empty:
        return df

    # Prefer integer _Year/_Month/_Day columns (no date parsing ambiguity)
    if all(c in df.columns for c in ('_Year', '_Month', '_Day')):
        y = pd.to_numeric(df['_Year'], errors='coerce').fillna(0).astype(int)
        m = pd.to_numeric(df['_Month'], errors='coerce').fillna(0).astype(int)
        d = pd.to_numeric(df['_Day'], errors='coerce').fillna(0).astype(int)
        ymd = y * 10000 + m * 100 + d
        start_ymd = start.year * 10000 + start.month * 100 + start.day
        end_ymd = end.year * 10000 + end.month * 100 + end.day
        mask = (ymd >= start_ymd) & (ymd <= end_ymd) & (ymd > 0)
        return df[mask].copy()

    # Fallback: use _Period for full-month ranges
    if '_Period' in df.columns:
        is_full_month = (
            start.day == 1
            and end.day == end.days_in_month
            and start.month == end.month
            and start.year == end.year
        )
        if is_full_month:
            target = f"{start.year}-{start.month:02d}"
            return df[df["_Period"].astype(str) == target].copy()

        periods = set()
        cur = start.to_period("M")
        end_period = end.to_period("M")
        while cur <= end_period:
            periods.add(str(cur))
            cur = cur + 1
        if periods:
            mask = df["_Period"].astype(str).isin(periods)
            return df[mask].copy()

    # Last resort: date string parsing
    for col in ["Comment Date", "Post Date"]:
        if col in df.columns:
            dt = _parse_date_col(df[col])
            mask = (dt >= start) & (dt <= end)
            return df[mask].copy()
    return df


# ==============================================================================
# Main entry: get_range_data
# ==============================================================================
_RANGE_CACHE: dict[str, tuple[float, RangeData]] = {}
_CACHE_TTL = 300

PERIOD_SHEETS = [
    "FB Wall Post Performance",
    "FB Pivot (Category)", "FB Pivot (Post Type)", "FB Pivot (Pillar)",
    "FB Key Metrics", "FB Key Metric (CM)",
    "Category Performance - BAU", "Sub-Category Performance - PNP",
    "Pillar Performance - CRM", "Pillar Performance - Ecommerce",
    "Pillar Performance - GNC", "Pillar Performance - Branding",
    "Pillar Performance - Category", "Pillar Performance - Sales",
    "Pillar Performance - Others",
    "IG Wall Post Performance", "IG Story Performance",
    "IG Pivot (Pillar)", "IG Story Pivot",
    "IG Key Metrics", "IG Key Metric (CM)",
    "Master Comments Base",
    "Sentiment Summary (Category)", "Sentiment Summary (Type)",
    "Unique Page View", "FB Followers", "IG Followers",
    "FB Reach Funnel",
]


def get_range_data(start_str: str, end_str: str) -> RangeData:
    cache_key = f"{start_str}_{end_str}"
    now = time.time()
    cached = _RANGE_CACHE.get(cache_key)
    if cached and (now - cached[0]) < _CACHE_TTL:
        return cached[1]

    start = pd.to_datetime(start_str)
    end = pd.to_datetime(end_str)
    end = end + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)  # include full end day

    xl = _get_excel_file(settings.excel_path)

    # Generate label
    if start.strftime("%Y-%m") == end.strftime("%Y-%m") and start.day == 1 and end.day == end.days_in_month:
        label = f"{MONTH_NAMES.get(start.month, '')} {start.year}"
    elif start == end:
        label = start.strftime("%d/%m/%Y")
    else:
        label = f"{start.strftime('%d/%m/%Y')} – {end.strftime('%d/%m/%Y')}"

    rd = RangeData(start=start_str, end=end_str, label=label, period_str=label)

    for sheet_name in PERIOD_SHEETS:
        if sheet_name not in xl.sheet_names:
            continue
        try:
            raw = pd.read_excel(xl, sheet_name)

            if sheet_name in POST_LEVEL_SHEETS:
                date_col = POST_LEVEL_SHEETS[sheet_name]
                filtered = _filter_by_date_range(raw, date_col, start, end)
            elif sheet_name in DAILY_SHEETS:
                filtered = _filter_daily(raw, start, end)
            elif sheet_name == "FB Reach Funnel":
                filtered = _filter_reach_funnel(raw, start, end)
            elif sheet_name in PIVOT_SHEETS:
                source_sheet, group_col = PIVOT_SHEETS[sheet_name]
                source_df = rd.get(source_sheet)
                filtered = _reaggregate_pivot(source_df, group_col)
            elif sheet_name == "Master Comments Base":
                filtered = _filter_comments(raw, start, end)
            elif sheet_name.startswith("Sentiment Summary"):
                comments = rd.get("Master Comments Base")
                filtered = _reaggregate_sentiment(raw, comments, sheet_name, start, end)
            elif sheet_name.startswith("FB Key Metric") or sheet_name.startswith("IG Key Metric"):
                filtered = raw  # Keep raw; API layer handles period selection
            else:
                filtered = raw

            rd.sheets[sheet_name] = filtered
        except Exception as e:
            logger.warning(f"Failed to read sheet '{sheet_name}': {e}")
            rd.sheets[sheet_name] = pd.DataFrame()

    _RANGE_CACHE[cache_key] = (now, rd)
    return rd


def _reaggregate_sentiment(raw_template: pd.DataFrame, comments: pd.DataFrame,
                            sheet_name: str, start: pd.Timestamp,
                            end: pd.Timestamp) -> pd.DataFrame:
    """Re-aggregate sentiment summary from filtered comments."""
    if comments.empty or "Sentiment" not in comments.columns:
        return pd.DataFrame()
    sent_col = "Sentiment"
    group_col = "Category" if "Category" in sheet_name else ("Type" if "type" in sheet_name.lower() else None)
    if group_col is None or group_col not in comments.columns:
        return pd.DataFrame()

    df = comments.copy()
    df[sent_col] = df[sent_col].astype(str).str.strip().str.title().replace("N/a", "N/A")
    df = df[df[sent_col].isin(["Positive", "Neutral", "Negative", "N/A"])]
    if df.empty:
        return pd.DataFrame()

    pivot = df.groupby([group_col, sent_col]).size().unstack(fill_value=0).reset_index()
    for s in ["Positive", "Neutral", "Negative", "N/A"]:
        if s not in pivot.columns:
            pivot[s] = 0
    pivot["Total"] = pivot[["Positive", "Neutral", "Negative", "N/A"]].sum(axis=1)
    pivot = pivot.sort_values("Total", ascending=False)
    pivot.insert(0, "Period", "Range")
    return pivot


# ==============================================================================
# Default range + available periods
# ==============================================================================
def default_range() -> tuple[str, str]:
    """Return (start, end) for last month from data."""
    today = date.today()
    if today.month > 1:
        lm_year, lm_month = today.year, today.month - 1
    else:
        lm_year, lm_month = today.year - 1, 12

    try:
        xl = _get_excel_file(settings.excel_path)
        # Check which periods have data
        available = set()
        for sheet in ["FB Pivot (Category)", "FB Wall Post Performance"]:
            if sheet in xl.sheet_names:
                df = pd.read_excel(xl, sheet)
                if "_Period" in df.columns:
                    available.update(df["_Period"].dropna().unique())
        target = f"{lm_year}-{lm_month:02d}"
        if target not in available and available:
            target = sorted(available)[-1]
            lm_year = int(target.split("-")[0])
            lm_month = int(target.split("-")[1])
    except Exception:
        pass

    start = date(lm_year, lm_month, 1)
    if lm_month == 12:
        end = date(lm_year, 12, 31)
    else:
        end = date(lm_year, lm_month + 1, 1) - timedelta(days=1)
    return start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")


def get_available_periods() -> list[dict]:
    """Return available periods as list of {value, label}."""
    try:
        xl = _get_excel_file(settings.excel_path)
    except FileNotFoundError:
        return []
    for sheet in ["FB Pivot (Category)", "FB Wall Post Performance"]:
        if sheet in xl.sheet_names:
            df = pd.read_excel(xl, sheet)
            if "_Period" in df.columns:
                ps = sorted(df["_Period"].dropna().unique())
                return [
                    {"value": str(p), "label": f"{MONTH_NAMES.get(int(str(p).split('-')[1]), '')} {str(p).split('-')[0]}"}
                    for p in ps
                ]
    return []
