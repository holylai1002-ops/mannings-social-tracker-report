"""
Data API routes — exposes Excel data as JSON for the frontend.
"""

import math
import yaml
import warnings
from pathlib import Path
from datetime import date

import pandas as pd
from fastapi import APIRouter, Request
from fastapi.responses import JSONResponse

from app.config import settings, BASE_DIR
from app.db.reader import (
    get_period_data, get_available_periods, default_period, MONTH_NAMES,
)

router = APIRouter()

_FPK_DROP_COLS = ["Profile-ID", "Link", "External Links", "Image Link", "Network", "Neutral"]

COMPANIES = [
    {"key": "sasa", "keywords": ["sasa", "莎莎"], "color": "#DD178C", "logo": "sasa.png"},
    {"key": "mannings", "keywords": ["manning", "萬寧"], "color": "#FE8301", "logo": "Mannings.png"},
    {"key": "matsumoto", "keywords": ["matsumoto", "松本清"], "color": "#F2EA40", "logo": "Matsumotokiyoshi.png"},
    {"key": "watsons", "keywords": ["watson"], "color": "#0E9F9F", "logo": "watsons.png"},
    {"key": "hktvmall", "keywords": ["hktv"], "color": "#B9D74D", "logo": "hktvmall.png"},
    {"key": "lungfung", "keywords": ["龍豐", "lungfung", "lung fung"], "color": "#DBDBDB", "logo": "lungfung.png"},
]


def _match_company(profile_name: str):
    name_lower = str(profile_name).lower()
    for c in COMPANIES:
        for kw in c["keywords"]:
            if kw.lower() in name_lower:
                return c
    return None


def _logo_url(logo_file: str) -> str:
    return f"/competitors/company_logo/{logo_file}"


def _is_date_col(col_name) -> bool:
    if not col_name or str(col_name).startswith("Unnamed"):
        return False
    try:
        pd.to_datetime(str(col_name))
        return True
    except Exception:
        return False


def _safe_int(val, default=0):
    try:
        if pd.isna(val):
            return default
        return int(val)
    except (ValueError, TypeError):
        return default


def _fmt_int(val):
    v = _safe_int(val)
    return f"{v:,}"


def _df_to_records(df: pd.DataFrame, max_rows: int = 500) -> list[dict]:
    if df.empty:
        return []
    df = df.head(max_rows).copy()
    # Drop _Period column (used for filtering, not display)
    if "_Period" in df.columns:
        df = df.drop(columns=["_Period"])
    for c in df.columns:
        if df[c].dtype == "float64":
            df[c] = df[c].apply(lambda x: round(x, 4) if pd.notna(x) else None)
        elif df[c].dtype == "int64":
            df[c] = df[c].apply(lambda x: int(x) if pd.notna(x) else None)
    drop_cols = [c for c in ["ai_insights"] if c in df.columns]
    if drop_cols:
        df = df.drop(columns=drop_cols)
    return df.fillna("").to_dict(orient="records")


def _filter_key_metrics(df: pd.DataFrame, period_str: str) -> pd.DataFrame:
    """Keep only Metric + the selected period column from a Key Metrics sheet."""
    if df.empty:
        return df
    keep = ["Metric", period_str]
    available = [c for c in keep if c in df.columns]
    if len(available) <= 1:
        return df
    return df[available]


def _api_followers_growth(pd_obj) -> dict | None:
    """Extract daily followers growth data for the selected period."""
    df = pd_obj.get("FB Followers")
    if df.empty or "Date" not in df.columns:
        return None
    dates = pd.to_datetime(df["Date"], format="%d/%m/%Y", errors="coerce")
    gain_col = "Followers Gain" if "Followers Gain" in df.columns else "Gain"
    loss_col = "Followers Loss" if "Followers Loss" in df.columns else "Loss"
    net_col = "Followers Net" if "Followers Net" in df.columns else "Net"
    net_total = int(pd.to_numeric(df[net_col], errors="coerce").fillna(0).sum())
    return {
        "dates": dates.dt.strftime("%d/%m/%Y").tolist(),
        "gain": [int(x) for x in pd.to_numeric(df[gain_col], errors="coerce").fillna(0)],
        "loss": [int(x) for x in pd.to_numeric(df[loss_col], errors="coerce").fillna(0)],
        "net": [int(x) for x in pd.to_numeric(df[net_col], errors="coerce").fillna(0)],
        "monthly_net": net_total,
    }


def _api_total_reach(pd_obj) -> dict | None:
    """Extract daily total reach data for the selected period."""
    df = pd_obj.get("Unique Page View")
    if df.empty or "Date" not in df.columns:
        return None
    dates = pd.to_datetime(df["Date"], format="%d/%m/%Y", errors="coerce")
    reach_col = "Unique Page View" if "Unique Page View" in df.columns else "Total Reach"
    reach_vals = pd.to_numeric(df[reach_col], errors="coerce").fillna(0).astype(int)
    monthly_total = int(reach_vals.sum())
    return {
        "dates": dates.dt.strftime("%d/%m/%Y").tolist(),
        "reach": reach_vals.tolist(),
        "monthly_total": monthly_total,
    }


def _api_reach_funnel(pd_obj) -> dict | None:
    """Extract organic vs paid reach totals for the selected period."""
    df = pd_obj.get("FB Reach Funnel")
    if df is None:
        df = pd_obj.get("Reach Funnel")
    if df is None or df.empty:
        return None
    organic = int(pd.to_numeric(df["Organic reach"], errors="coerce").fillna(0).sum())
    paid = int(pd.to_numeric(df["Paid reach"], errors="coerce").fillna(0).sum())
    return {
        "organic": organic,
        "paid": paid,
        "total": organic + paid,
    }


def _api_ig_followers(pd_obj) -> dict | None:
    """Extract daily IG followers data for the selected period."""
    df = pd_obj.get("IG Followers")
    if df.empty or "Date" not in df.columns:
        return None
    dates = pd.to_datetime(df["Date"], format="%d/%m/%Y", errors="coerce")
    net_col = "Followers Net" if "Followers Net" in df.columns else "Net"
    net_vals = pd.to_numeric(df[net_col], errors="coerce").fillna(0).astype(int)
    monthly_net = int(net_vals.sum())
    total_col = "Total Followers" if "Total Followers" in df.columns else None
    total_vals = pd.to_numeric(df[total_col], errors="coerce").fillna(0).astype(int).tolist() if total_col else []
    return {
        "dates": dates.dt.strftime("%d/%m/%Y").tolist(),
        "net": net_vals.tolist(),
        "monthly_net": monthly_net,
        "total": total_vals,
    }


def _parse_fpk_xlsx(filepath: Path) -> list[dict]:
    """Parse a FPK XLSX export: skip 4 metadata rows, drop link/image columns."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df = pd.read_excel(filepath, header=4)
    # Drop unnamed columns
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    # Drop metadata columns
    drop = [c for c in _FPK_DROP_COLS if c in df.columns]
    if drop:
        df = df.drop(columns=drop)
    # Drop fully empty rows
    df = df.dropna(how="all")
    # Round floats
    for c in df.columns:
        if df[c].dtype == "float64":
            df[c] = df[c].apply(lambda x: round(x, 4) if pd.notna(x) else None)
    return df.fillna("").to_dict(orient="records")


@router.get("/api/periods")
def api_periods():
    periods = get_available_periods()
    dy, dm = default_period()
    return {
        "periods": [{"year": y, "month": m, "label": f"{MONTH_NAMES[m]} {y}"} for y, m in periods],
        "default": {"year": dy, "month": dm},
    }


@router.get("/api/data/kpis")
def api_kpis(year: int = 0, month: int = 0):
    if year == 0 or month == 0:
        year, month = default_period()
    pd_obj = get_period_data(year, month)
    kpis = pd_obj.kpis()
    period_label = f"{MONTH_NAMES.get(month, '')} {year}"
    return {"period": pd_obj.period_str, "label": period_label, **kpis}


@router.get("/api/data/fb_page")
def api_fb_page(year: int = 0, month: int = 0):
    if year == 0 or month == 0:
        year, month = default_period()
    pd_obj = get_period_data(year, month)

    comments = pd_obj.get("Master Comments Base")
    sentiment_data = []
    if not comments.empty and "Sentiment" in comments.columns:
        sent = comments["Sentiment"].astype(str).str.strip().str.title().replace("N/a", "N/A")
        sent = sent[sent.isin(["Positive", "Neutral", "Negative", "N/A"])]
        counts = sent.value_counts()
        sentiment_data = [{"name": s, "value": int(counts.get(s, 0))} for s in ["Positive", "Neutral", "Negative", "N/A"] if counts.get(s, 0) > 0]

    sentiment_by_cat = []
    if not comments.empty and "Sentiment" in comments.columns and "Category" in comments.columns:
        cat_sent = comments.copy()
        cat_sent["Sentiment"] = cat_sent["Sentiment"].astype(str).str.strip().str.title()
        cat_sent = cat_sent[cat_sent["Sentiment"].isin(["Positive", "Neutral", "Negative"])]
        if not cat_sent.empty:
            pivot = cat_sent.groupby("Category")["Sentiment"].value_counts().unstack(fill_value=0).reset_index()
            for _, row in pivot.iterrows():
                sentiment_by_cat.append({
                    "category": str(row["Category"]),
                    "Positive": int(row.get("Positive", 0)),
                    "Neutral": int(row.get("Neutral", 0)),
                    "Negative": int(row.get("Negative", 0)),
                })

    sentiment_by_type = []
    type_col = None
    if not comments.empty:
        for c in ["Type", "type"]:
            if c in comments.columns:
                type_col = c
                break
    if type_col and "Sentiment" in comments.columns:
        t_sent = comments.copy()
        t_sent["Sentiment"] = t_sent["Sentiment"].astype(str).str.strip().str.title()
        t_sent[type_col] = t_sent[type_col].astype(str).str.strip().str.title()
        t_sent = t_sent[t_sent["Sentiment"].isin(["Positive", "Neutral", "Negative"])]
        if not t_sent.empty:
            pivot = t_sent.groupby(type_col)["Sentiment"].value_counts().unstack(fill_value=0).reset_index()
            for _, row in pivot.iterrows():
                sentiment_by_type.append({
                    "type": str(row[type_col]),
                    "Positive": int(row.get("Positive", 0)),
                    "Neutral": int(row.get("Neutral", 0)),
                    "Negative": int(row.get("Negative", 0)),
                })

    # FB Key Metrics — filtered to selected period column
    fb_key = pd_obj.get("FB Key Metrics")
    fb_key = _filter_key_metrics(fb_key, pd_obj.period_str)
    fb_key_metrics = _df_to_records(fb_key)

    return {
        "period": pd_obj.period_str,
        "label": f"{MONTH_NAMES.get(month, '')} {year}",
        "sentiment_distribution": sentiment_data,
        "sentiment_by_category": sentiment_by_cat,
        "sentiment_by_type": sentiment_by_type,
        "fb_key_metrics": fb_key_metrics,
        "followers_growth": _api_followers_growth(pd_obj),
        "total_reach": _api_total_reach(pd_obj),
        "reach_funnel": _api_reach_funnel(pd_obj),
        "ig_followers": _api_ig_followers(pd_obj),
    }


@router.get("/api/data/fb_posts")
def api_fb_posts(year: int = 0, month: int = 0):
    if year == 0 or month == 0:
        year, month = default_period()
    pd_obj = get_period_data(year, month)

    pillar = pd_obj.get("FB Pivot (Pillar)")
    cat = pd_obj.get("FB Pivot (Category)")
    ptype = pd_obj.get("FB Pivot (Post Type)")
    wall = pd_obj.get("FB Wall Post Performance")

    pillar_donut = []
    if not pillar.empty and "Pillar" in pillar.columns and "No. of Posts" in pillar.columns:
        pillar_donut = [{"name": str(r["Pillar"]), "value": _safe_int(r["No. of Posts"])} for _, r in pillar.iterrows()]

    type_donut = []
    if not ptype.empty and "Post type" in ptype.columns and "No. of Posts" in ptype.columns:
        type_donut = [{"name": str(r["Post type"]), "value": _safe_int(r["No. of Posts"])} for _, r in ptype.iterrows()]

    pillar_interactions = []
    if not pillar.empty and "Pillar" in pillar.columns and "Total Interactions" in pillar.columns:
        df_sorted = pillar.sort_values("Total Interactions", ascending=False)
        pillar_interactions = [{"name": str(r["Pillar"]), "value": _safe_int(r["Total Interactions"])} for _, r in df_sorted.iterrows()]

    cat_interactions = []
    if not cat.empty and "Category" in cat.columns and "Total Interactions" in cat.columns:
        df_sorted = cat.sort_values("Total Interactions", ascending=False)
        cat_interactions = [{"name": str(r["Category"]), "value": _safe_int(r["Total Interactions"])} for _, r in df_sorted.iterrows()]

    breakdowns = {}
    breakdown_sheets = [
        "Category Performance - BAU", "Sub-Category Performance - PNP",
        "Pillar Performance - CRM", "Pillar Performance - Ecommerce",
        "Pillar Performance - GNC",
        "Pillar Performance - Branding",
        "Pillar Performance - Category",
        "Pillar Performance - Sales",
        "Pillar Performance - Others",
    ]
    perf_cols = ["Permalink", "Description", "Post type", "Pillar", "Category", "Sub-Category",
                 "Campaign Name", "Publish time", "Reactions", "Comments", "Shares",
                 "Interactions", "Reach", "Views", "Link Clicks", "Photo Clicks"]
    for sheet in breakdown_sheets:
        df = pd_obj.get(sheet)
        if df.empty:
            continue
        display_cols = [c for c in perf_cols if c in df.columns]
        breakdowns[sheet] = {
            "count": len(df),
            "reach": _safe_int(df["Reach"].sum()) if "Reach" in df.columns else 0,
            "interactions": _safe_int(df["Interactions"].sum()) if "Interactions" in df.columns else 0,
            "records": _df_to_records(df[display_cols]) if display_cols else _df_to_records(df),
        }

    return {
        "period": pd_obj.period_str,
        "label": f"{MONTH_NAMES.get(month, '')} {year}",
        "pillar_donut": pillar_donut,
        "type_donut": type_donut,
        "pillar_interactions": pillar_interactions,
        "cat_interactions": cat_interactions,
        "wall_posts": _df_to_records(wall),
        "breakdowns": breakdowns,
    }


@router.get("/api/data/instagram")
def api_instagram(year: int = 0, month: int = 0):
    if year == 0 or month == 0:
        year, month = default_period()
    pd_obj = get_period_data(year, month)

    # IG Key Metrics — filtered to selected period column
    ig_key = _filter_key_metrics(pd_obj.get("IG Key Metrics"), pd_obj.period_str)
    ig_pivot = pd_obj.get("IG Pivot (Pillar)")
    ig_story_pivot = pd_obj.get("IG Story Pivot")
    ig_eng = pd_obj.get("IG Engagement Pivot")
    ig_story_cat = pd_obj.get("IG Story Category Pivot")
    ig_wall = pd_obj.get("IG Wall Post Performance")

    # IG Key Metrics — filtered to selected period column
    ig_key_metrics = _df_to_records(ig_key)

    ig_pillar_donut = []
    if not ig_pivot.empty and "Pillar" in ig_pivot.columns and "No. of Posts" in ig_pivot.columns:
        ig_pillar_donut = [{"name": str(r["Pillar"]), "value": _safe_int(r["No. of Posts"])} for _, r in ig_pivot.iterrows()]

    pillar_interactions = []
    if not ig_pivot.empty and "Pillar" in ig_pivot.columns and "Total Interactions" in ig_pivot.columns:
        pillar_interactions = [{"name": str(r["Pillar"]), "value": _safe_int(r["Total Interactions"])} for _, r in ig_pivot.iterrows()]

    eng_donut = []
    if not ig_eng.empty and "Engagement Type" in ig_eng.columns and "Count" in ig_eng.columns:
        eng_donut = [{"name": str(r["Engagement Type"]), "value": _safe_int(r["Count"])} for _, r in ig_eng.iterrows()]

    story_cat_donut = []
    if not ig_story_cat.empty and "Category" in ig_story_cat.columns and "Count" in ig_story_cat.columns:
        story_cat_donut = [{"name": str(r["Category"]), "value": _safe_int(r["Count"])} for _, r in ig_story_cat.iterrows()]

    story_clicks = []
    if not ig_story_pivot.empty and "Pillar" in ig_story_pivot.columns:
        click_col = "SUM of Link clicks" if "SUM of Link clicks" in ig_story_pivot.columns else ("Total Interactions" if "Total Interactions" in ig_story_pivot.columns else None)
        if click_col:
            df_sorted = ig_story_pivot.sort_values(click_col, ascending=False)
            story_clicks = [{"name": str(r["Pillar"]), "value": _safe_int(r[click_col])} for _, r in df_sorted.iterrows()]

    return {
        "period": pd_obj.period_str,
        "label": f"{MONTH_NAMES.get(month, '')} {year}",
        "ig_key_metrics": ig_key_metrics,
        "pillar_donut": ig_pillar_donut,
        "pillar_interactions": pillar_interactions,
        "engagement_donut": eng_donut,
        "story_cat_donut": story_cat_donut,
        "story_clicks": story_clicks,
        "wall_posts": _df_to_records(ig_wall),
        "ig_story_posts": _df_to_records(pd_obj.get("IG Story Performance")),
        "ig_followers": _api_ig_followers(pd_obj),
    }


@router.get("/api/data/fpk")
def api_fpk(year: int = 0, month: int = 0, platform: str = "fb"):
    """Return chart-ready FPK competitor data for a period."""
    if year == 0 or month == 0:
        y, m = default_period()
    else:
        y, m = year, month

    base = Path(settings.competitors_dir)
    folder = base / f"{platform.upper()} competitor"
    ym = f"{y}{m:02d}"
    pf = platform.upper()

    result = {"period": f"{y}-{m:02d}", "platform": platform}

    # ── Growth Trend → line chart ──
    gt_file = folder / f"{pf} Competitor Fans Growth Trend_{ym}.xlsx"
    if gt_file.exists():
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df = pd.read_excel(gt_file, header=4)
        df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
        date_cols = [c for c in df.columns if _is_date_col(c)]
        series = []
        for _, row in df.iterrows():
            comp = _match_company(row.get("Profile", ""))
            if not comp:
                continue
            series.append({
                "key": comp["key"],
                "name": str(row.get("Profile", comp["key"])),
                "color": comp["color"],
                "logo": _logo_url(comp["logo"]),
                "data": [int(row[c]) if pd.notna(row[c]) else None for c in date_cols],
            })
        result["growth_trend"] = {"dates": [str(c) for c in date_cols], "series": series}

    # ── Competitors Overview → table ──
    co_file = folder / f"{pf} Competitors Overview_{ym}.xlsx"
    if co_file.exists():
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df = pd.read_excel(co_file, header=4)
        df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
        keep_cols = [c for c in ["Profile", "Follower", "Follower Growth (absolute)",
                                  "Follower Growth (in %)", "Number of posts",
                                  "Reactions, Comments & Shares"] if c in df.columns]
        records = []
        for _, row in df.iterrows():
            comp = _match_company(row.get("Profile", ""))
            if not comp:
                continue
            rec = {"key": comp["key"], "color": comp["color"], "logo": _logo_url(comp["logo"])}
            for c in keep_cols:
                val = row[c]
                if c == "Follower Growth (in %)" and pd.notna(val):
                    rec[c] = f"{val * 100:.2g}%"
                elif pd.notna(val):
                    rec[c] = round(val, 4) if isinstance(val, float) else val
                else:
                    rec[c] = ""
            records.append(rec)
        result["competitors_overview"] = records

    # ── KPI Comparison → bubble chart ──
    kpi_file = folder / f"{pf} Key Performance Indexes Comparison_{ym}.xlsx"
    if kpi_file.exists():
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df = pd.read_excel(kpi_file, header=4)
        df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
        posts_col = "Number of posts"
        react_col = "Reactions, Comments & Shares"
        items = []
        for _, row in df.iterrows():
            comp = _match_company(row.get("Profile", ""))
            if not comp:
                continue
            items.append({
                "key": comp["key"],
                "name": str(row.get("Profile", comp["key"])),
                "color": comp["color"],
                "logo": _logo_url(comp["logo"]),
                "posts": int(row[posts_col]) if posts_col in df.columns and pd.notna(row[posts_col]) else 0,
                "reactions": int(row[react_col]) if react_col in df.columns and pd.notna(row[react_col]) else 0,
            })
        avg_posts = sum(i["posts"] for i in items) / len(items) if items else 0
        avg_reactions = sum(i["reactions"] for i in items) / len(items) if items else 0
        result["kpi_comparison"] = {"avg_posts": round(avg_posts, 1), "avg_reactions": round(avg_reactions, 1), "items": items}

    # ── Top 50 Words → tag cloud (IG only) ──
    tw_file = folder / f"IG Top 50 Words Post interaction rate_{ym}.xlsx"
    if platform == "ig" and tw_file.exists():
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df = pd.read_excel(tw_file, header=4)
        df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
        words = []
        for _, row in df.iterrows():
            word = row.get("Profile", "")
            if not word or pd.isna(word):
                continue
            words.append({
                "word": str(word),
                "value": int(row["value"]) if "value" in df.columns and pd.notna(row["value"]) else 0,
                "times_above_avg": int(row["Times above average"]) if "Times above average" in df.columns and pd.notna(row["Times above average"]) else 0,
            })
        result["top_words"] = words

    return result


@router.get("/api/data/linkedin")
def api_linkedin(year: int = 0, month: int = 0):
    if year == 0 or month == 0:
        year, month = default_period()
    pd_obj = get_period_data(year, month)

    fl = pd_obj.get("LinkedIn Follower Log")
    pp = pd_obj.get("LinkedIn Page Perf")
    posts = pd_obj.get("LinkedIn Posts Perf")

    # Normalize column name (pipeline renames "Date Polled"→"Date"; manual Excel may not)
    if not fl.empty and "Date Polled" in fl.columns and "Date" not in fl.columns:
        fl = fl.rename(columns={"Date Polled": "Date"})

    _MONTH_ABBR = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
                   7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"}
    _abbr_val = {v: k for k, v in _MONTH_ABBR.items()}
    month_abbr = _MONTH_ABBR.get(month, "")

    # ── Follower Log: filter Organic/Paid by Month column ──
    fl_organic = []
    fl_paid = []
    if not fl.empty:
        fl_org_paid = fl[fl["Category"].astype(str).str.strip().str.lower().isin(["organic", "paid"])].copy()
        if "Month" in fl_org_paid.columns:
            fl_org_paid = fl_org_paid[fl_org_paid["Month"].astype(str).str.strip().str.title() == month_abbr]
        if not fl_org_paid.empty:
            for _, row in fl_org_paid.iterrows():
                cat = str(row.get("Category", "")).strip().lower()
                val = _safe_int(row.get("Total Followers", 0))
                date_str = str(row.get("Date", ""))
                entry = {"date": date_str, "value": val}
                if cat == "organic":
                    fl_organic.append(entry)
                elif cat == "paid":
                    fl_paid.append(entry)

    total_organic = sum(d["value"] for d in fl_organic)
    total_paid = sum(d["value"] for d in fl_paid)
    total_new = total_organic + total_paid

    # ── Cumulative total from Net category (latest value in months <= selected) ──
    cumulative = 0
    if not fl.empty:
        fl_net = fl[fl["Category"].astype(str).str.strip().str.lower() == "net"].copy()
        if "Month" in fl_net.columns and not fl_net.empty:
            fl_net = fl_net[fl_net["Month"].astype(str).str.strip().str.title().map(lambda x: _abbr_val.get(x, 0)) <= month]
        if not fl_net.empty:
            fl_net["_dt"] = pd.to_datetime(fl_net["Date"], format="mixed", errors="coerce")
            latest_row = fl_net.loc[fl_net["_dt"].idxmax()]
            cumulative = int(latest_row["Total Followers"])

    # ── Organic vs Paid donut ──
    org_paid_donut = [
        {"name": "Organic", "value": total_organic},
        {"name": "Paid", "value": total_paid},
    ]

    # ── Net Followers bar (daily organic + paid grouped) ──
    net_followers = None
    if fl_organic or fl_paid:
        org_map = {d["date"]: d["value"] for d in fl_organic}
        paid_map = {d["date"]: d["value"] for d in fl_paid}
        all_dates = sorted(set(list(org_map.keys()) + list(paid_map.keys())),
                           key=lambda x: pd.to_datetime(x, format="mixed", errors="coerce"))
        net_followers = {
            "dates": [pd.to_datetime(d, format="mixed", errors="coerce").strftime("%Y-%m-%d") if d else d for d in all_dates],
            "organic": [org_map.get(d, 0) for d in all_dates],
            "paid": [paid_map.get(d, 0) for d in all_dates],
        }

    # ── Demographics (latest snapshot) ──
    demo_data = {"location": [], "company_size": [], "seniority": [], "job_function": [], "industry": []}
    if not fl.empty and "Category" in fl.columns:
        prefix_map = {"Location": "location", "Company Size": "company_size",
                      "Seniority": "seniority", "Job Function": "job_function", "Industry": "industry"}
        for _, row in fl.iterrows():
            cat = str(row.get("Category", "")).strip()
            demo = str(row.get("Demographic", "")).strip()
            val = _safe_int(row.get("Total Followers", 0))
            log_date = str(row.get("Log Date", "")).strip()
            for prefix, key in prefix_map.items():
                if cat.startswith("Top ") and cat.endswith(prefix):
                    rank_str = cat.replace("Top ", "").replace(" " + prefix, "").strip()
                    demo_data[key].append({"name": demo, "value": val, "rank": rank_str, "log_date": log_date})

    def _latest_top5(items):
        if not items:
            return []
        latest = max(i["log_date"] for i in items if i["log_date"])
        latest_items = [i for i in items if i["log_date"] == latest]
        latest_items.sort(key=lambda x: int(x["rank"]) if str(x["rank"]).isdigit() else 99)
        seen_ranks = set()
        result = []
        for i in latest_items:
            if i["rank"] not in seen_ranks and i["name"]:
                seen_ranks.add(i["rank"])
                result.append({"name": i["name"], "value": i["value"]})
        return result[:5]

    top_countries = _latest_top5(demo_data["location"])
    top_company_size = _latest_top5(demo_data["company_size"])
    top_seniority = _latest_top5(demo_data["seniority"])
    top_job_function = _latest_top5(demo_data["job_function"])
    top_industry = _latest_top5(demo_data["industry"])

    # ── Page Perf: filter by Month column ──
    clicks_data = None
    impressions_data = None
    social_actions = None
    visitor_metrics = None
    if not pp.empty:
        pp_filtered = pp
        if "Month" in pp.columns:
            pp_filtered = pp[pp["Month"].astype(str).str.strip().str.title() == month_abbr].copy()
        if not pp_filtered.empty:
            _pp_dt = pd.to_datetime(pp_filtered["Date"], format="mixed", errors="coerce")
            pp_filtered = pp_filtered.copy()
            pp_filtered["_dt"] = _pp_dt
            pp_filtered = pp_filtered.sort_values("_dt")
            pp_dates = _pp_dt.dt.strftime("%Y-%m-%d").tolist() if "Date" in pp_filtered.columns else []

            clicks_data = {
                "dates": pp_dates,
                "values": [int(x) for x in pd.to_numeric(pp_filtered["Clicks"], errors="coerce").fillna(0)] if "Clicks" in pp_filtered.columns else [],
            }
            impressions_data = {
                "dates": pp_dates,
                "values": [int(x) for x in pd.to_numeric(pp_filtered["Impressions"], errors="coerce").fillna(0)] if "Impressions" in pp_filtered.columns else [],
            }
            social_actions = {
                "dates": pp_dates,
                "comments": [int(x) for x in pd.to_numeric(pp_filtered["Comments"], errors="coerce").fillna(0)] if "Comments" in pp_filtered.columns else [],
                "likes": [int(x) for x in pd.to_numeric(pp_filtered["Likes"], errors="coerce").fillna(0)] if "Likes" in pp_filtered.columns else [],
                "shares": [int(x) for x in pd.to_numeric(pp_filtered["Shares"], errors="coerce").fillna(0)] if "Shares" in pp_filtered.columns else [],
            }
            visitor_metrics = {
                "dates": pp_dates,
                "desktop": [int(x) for x in pd.to_numeric(pp_filtered["Desktop Views"], errors="coerce").fillna(0)] if "Desktop Views" in pp_filtered.columns else [],
                "mobile": [int(x) for x in pd.to_numeric(pp_filtered["Mobile Views"], errors="coerce").fillna(0)] if "Mobile Views" in pp_filtered.columns else [],
            }

    # ── Posts table (filter by Month column) ──
    posts_filtered = pd.DataFrame()
    if not posts.empty:
        if "Month" in posts.columns:
            posts_filtered = posts[posts["Month"].astype(str).str.strip().str.title() == month_abbr].copy()

    posts_records = _df_to_records(posts_filtered) if not posts_filtered.empty else []

    return {
        "period": pd_obj.period_str,
        "label": f"{MONTH_NAMES.get(month, '')} {year}",
        "total_followers": cumulative,
        "new_followers": total_new,
        "organic_followers": total_organic,
        "paid_followers": total_paid,
        "org_paid_donut": org_paid_donut,
        "net_followers": net_followers,
        "clicks": clicks_data,
        "impressions": impressions_data,
        "social_actions": social_actions,
        "visitor_metrics": visitor_metrics,
        "top_countries": top_countries,
        "top_company_size": top_company_size,
        "top_seniority": top_seniority,
        "top_job_function": top_job_function,
        "top_industry": top_industry,
        "posts": posts_records,
        "post_count": len(posts_records),
        "total_impressions": sum(_safe_int(r.get("Impressions", 0)) for r in posts_records),
        "total_interactions": sum(_safe_int(r.get("Total Interactions", 0)) for r in posts_records),
    }


@router.get("/api/images")
def api_images(year: int = 0, month: int = 0):
    """List FPK screenshots for a period."""
    if year == 0 or month == 0:
        y, m = default_period()
    else:
        y, m = year, month

    period_dir = Path(settings.competitors_dir) / f"{y}_{m:02d}"
    config_path = Path(settings.screenshots_config)

    configured = {}
    if config_path.exists():
        with open(config_path, "r", encoding="utf-8") as f:
            cfg = yaml.safe_load(f) or {}
        configured = {s["file"]: s for s in cfg.get("screenshots", [])}

    images = []
    if period_dir.exists():
        for filepath in sorted(period_dir.iterdir()):
            if filepath.suffix.lower() not in [".png", ".jpg", ".jpeg", ".webp"]:
                continue
            stem = filepath.stem
            cfg = configured.get(stem, {})
            images.append({
                "file": filepath.name,
                "url": f"/competitors/{y}_{m:02d}/{filepath.name}",
                "label": cfg.get("label", stem.replace("_", " ").title()),
                "description": cfg.get("description", ""),
                "configured": bool(cfg),
            })
    return {"period": f"{y}-{m:02d}", "images": images}
