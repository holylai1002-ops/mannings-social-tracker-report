"""
Excel reader — loads dashboard data from Mannings_FB_IG_Dashboard_Feed.xlsx.

Caches DataFrames in memory keyed by (path, mtime).
Provides period-filtered access to all 19+ sheets.
"""

import time
import logging
from datetime import date
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path

import pandas as pd

from app.config import settings

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


def _get_metric(key_df: pd.DataFrame, metric_name: str, period_str: str) -> int:
    if key_df.empty or "Metric" not in key_df.columns or period_str not in key_df.columns:
        return 0
    row = key_df[key_df["Metric"] == metric_name]
    if not row.empty:
        try:
            val = row[period_str].values[0]
            if pd.notna(val):
                return int(val)
        except (ValueError, TypeError):
            pass
    return 0


@dataclass
class PeriodData:
    year: int
    month: int
    period_str: str
    sheets: dict = field(default_factory=dict)

    def get(self, name: str) -> pd.DataFrame:
        return self.sheets.get(name, pd.DataFrame())

    def kpis(self) -> dict:
        fb_key = self.get("FB Key Metrics")
        ig_key = self.get("IG Key Metrics")

        return {
            "fb_followers": _get_metric(fb_key, "Total Followers", self.period_str),
            "fb_growth": _get_metric(fb_key, "Net Followers Growth", self.period_str),
            "fb_wall_posts": _get_metric(fb_key, "No. of Wall Post", self.period_str),
            "fb_interactions": _get_metric(fb_key, "Total Interaction*", self.period_str),
            "ig_followers": _get_metric(ig_key, "Total Followers", self.period_str),
            "ig_growth": _get_metric(ig_key, "Net Followers Growth", self.period_str),
            "ig_reach": _get_metric(ig_key, "Total Post Reach", self.period_str),
            "ig_interactions": _get_metric(ig_key, "Total Post Interaction^", self.period_str),
        }


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


def get_excel_for_period(year: int, month: int) -> pd.ExcelFile:
    base = Path(settings.excel_path)
    period_file = base.parent / f"Mannings_FB_IG_Dashboard_Feed_{year}_{month:02d}.xlsx"
    if period_file.exists():
        return _get_excel_file(str(period_file))
    if base.exists():
        return _get_excel_file(str(base))
    raise FileNotFoundError(f"No Excel file found for {year}-{month:02d}")


def _filter_by_date(df: pd.DataFrame, year: int, month: int, date_col: str) -> pd.DataFrame:
    if df.empty or date_col not in df.columns:
        return pd.DataFrame()
    pt = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
    return df[(pt.dt.year == year) & (pt.dt.month == month)].copy()


def _filter_by_period(df: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    p = f"{year}-{month:02d}"
    for c in ["Period", "_Period"]:
        if c in df.columns:
            return df[df[c] == p].copy()
    return pd.DataFrame()


def _filter_comments(df: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
    if df.empty:
        return df
    if "Month Note" in df.columns:
        mn = df["Month Note"].astype(str).str.strip().str.title()
        month_match = mn.map(MONTH_MAP) == month
        if "_Period" in df.columns:
            year_match = df["_Period"].astype(str).str.startswith(str(year))
            return df[month_match & year_match].copy()
        return df[month_match].copy()
    return df


_PERIOD_CACHE: dict[str, tuple[float, PeriodData]] = {}
_CACHE_TTL = 300


def get_period_data(year: int, month: int) -> PeriodData:
    cache_key = f"{year}-{month:02d}"
    now = time.time()
    cached = _PERIOD_CACHE.get(cache_key)
    if cached and (now - cached[0]) < _CACHE_TTL:
        return cached[1]

    xl = get_excel_for_period(year, month)
    period_str = f"{year}-{month:02d}"
    pd_obj = PeriodData(year=year, month=month, period_str=period_str)

    # If a period-specific file exists, it is already pre-filtered — skip date/period filtering
    base = Path(settings.excel_path)
    period_file = base.parent / f"Mannings_FB_IG_Dashboard_Feed_{year}_{month:02d}.xlsx"
    is_period_file = period_file.exists()

    sheet_map = {
        "FB Wall Post Performance": ("date", "Publish time"),
        "FB Pivot (Category)": ("period", None),
        "FB Pivot (Post Type)": ("period", None),
        "FB Pivot (Pillar)": ("period", None),
        "FB Key Metrics": ("raw", None),
        "Category Performance - BAU": ("date", "Publish time"),
        "Sub-Category Performance - PNP": ("date", "Publish time"),
        "Pillar Performance - CRM": ("date", "Publish time"),
        "Pillar Performance - Ecommerce": ("date", "Publish time"),
        "Pillar Performance - GNC": ("date", "Publish time"),
        "Pillar Performance - Others": ("date", "Publish time"),
        "IG Wall Post Performance": ("date", "Post Date"),
        "IG Story Performance": ("date", "Publish time"),
        "IG Pivot (Pillar)": ("period", None),
        "IG Story Pivot": ("period", None),
        "IG Key Metrics": ("raw", None),
        "IG Engagement Pivot": ("raw", None),
        "IG Story Category Pivot": ("raw", None),
        "Master Comments Base": ("comments", None),
        "Sentiment Summary (Category)": ("period", None),
        "Sentiment Summary (Type)": ("period", None),
    }

    for sheet_name, (filter_type, date_col) in sheet_map.items():
        if sheet_name not in xl.sheet_names:
            continue
        try:
            raw = pd.read_excel(xl, sheet_name)
            if is_period_file:
                filtered = raw
            elif filter_type == "date" and date_col:
                filtered = _filter_by_date(raw, year, month, date_col)
            elif filter_type == "period":
                filtered = _filter_by_period(raw, year, month)
            elif filter_type == "comments":
                filtered = _filter_comments(raw, year, month)
            else:
                filtered = raw
            pd_obj.sheets[sheet_name] = filtered
        except Exception as e:
            logger.warning(f"Failed to read sheet '{sheet_name}': {e}")
            pd_obj.sheets[sheet_name] = pd.DataFrame()

    _PERIOD_CACHE[cache_key] = (now, pd_obj)
    return pd_obj


def get_available_periods() -> list[tuple[int, int]]:
    try:
        xl = get_excel_for_period(2025, 1)
    except FileNotFoundError:
        return []
    for sheet in ["Sentiment Summary (Category)", "FB Pivot (Category)", "FB Pivot (Pillar)"]:
        if sheet in xl.sheet_names:
            df = pd.read_excel(xl, sheet)
            for c in ["Period", "_Period"]:
                if c in df.columns:
                    ps = sorted(df[c].dropna().unique())
                    if ps:
                        return [(int(str(p).split("-")[0]), int(str(p).split("-")[1])) for p in ps]
    return []


def default_period() -> tuple[int, int]:
    base = Path(settings.excel_path)
    periods = get_available_periods()
    # Prefer the latest period that has a complete period-specific file
    for y, m in reversed(periods):
        period_file = base.parent / f"Mannings_FB_IG_Dashboard_Feed_{y}_{m:02d}.xlsx"
        if period_file.exists():
            return y, m
    # Fallback to last month
    today = date.today()
    default_year = today.year if today.month > 1 else today.year - 1
    default_month = today.month - 1 if today.month > 1 else 12
    if (default_year, default_month) in periods:
        return default_year, default_month
    if periods:
        return periods[-1]
    return default_year, default_month
