"""
Data API routes for app_v2 — date-range based filtering.
"""

import warnings
from pathlib import Path
from datetime import date, timedelta

import pandas as pd
from fastapi import APIRouter
from fastapi.responses import JSONResponse

from app_v2.config import settings, BASE_DIR
from app_v2.db.range_reader import get_range_data, default_range, MONTH_NAMES

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


def _df_to_records(df: pd.DataFrame, max_rows: int = 500) -> list[dict]:
    if df.empty:
        return []
    df = df.head(max_rows).copy()
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


def _resolve_dates(start: str = "", end: str = "") -> tuple[str, str]:
    if not start or not end:
        return default_range()
    return start, end


def _compute_fb_key_metrics(rd) -> list[dict]:
    """Compute FB Key Metrics from filtered post data."""
    df = rd.get("FB Wall Post Performance")
    fb_fol = rd.get("FB Followers")
    label = rd.label

    def _sum(col):
        if df.empty or col not in df.columns:
            return 0
        return int(pd.to_numeric(df[col].astype(str).str.replace(",", ""), errors="coerce").fillna(0).sum())

    def _count_val(col, val):
        if df.empty or col not in df.columns:
            return 0
        return int((df[col].astype(str).str.strip() == val).sum())

    feeds = len(df)
    dark = _count_val("Dark Post", "Y")
    organic = _count_val("Organic", "Organic")
    paid = _count_val("Paid", "Paid")

    vv_col = "3-second video views"
    no_video = 0
    if not df.empty and vv_col in df.columns:
        col_raw = df[vv_col].astype(str).str.strip()
        vv = pd.to_numeric(col_raw, errors="coerce")
        no_video = int(((col_raw.str.len() > 0) & vv.notna()).sum())

    total_followers = 0
    net_growth = 0
    if not fb_fol.empty:
        if "Total Followers" in fb_fol.columns:
            total_followers = int(pd.to_numeric(fb_fol["Total Followers"], errors="coerce").fillna(0).iloc[0])
        if "Followers Net" in fb_fol.columns:
            net_growth = int(pd.to_numeric(fb_fol["Followers Net"], errors="coerce").fillna(0).sum())

    d = no_video if no_video > 0 else 1
    metrics = {
        "Total Followers": total_followers,
        "Net Followers Growth": net_growth,
        "Feeds Posted": feeds,
        "Total Dark Posts": dark,
        "Total Organic": organic,
        "Organic %": round(organic / feeds, 4) if feeds > 0 else 0,
        "Total Paid": paid,
        "Paid %": round(paid / feeds, 4) if feeds > 0 else 0,
        "No. of Video Post": no_video,
        "Total Interactions": _sum("Interactions"),
        "Total Reactions": _sum("Reactions"),
        "Total Comments": _sum("Comments"),
        "Total Shares": _sum("Shares"),
        "Total Video Views": _sum("3-second video views"),
        "Average Organic Reach": int(round(_sum("Organic reach") / d)),
        "Average Paid Reach": int(round(_sum("Paid reach") / d)),
        "Average Interaction": int(round(_sum("Interactions") / d)),
        "Average Reactions": int(round(_sum("Reactions") / d)),
        "Average Comments": int(round(_sum("Comments") / d)),
        "Average Shares": int(round(_sum("Shares") / d)),
        "Average Video Views": int(round(_sum("3-second video views") / d)),
    }
    return [{"Metric": k, label: v} for k, v in metrics.items()]


def _compute_ig_key_metrics(rd) -> list[dict]:
    """Compute IG Key Metrics from filtered post data."""
    df_post = rd.get("IG Wall Post Performance")
    df_story = rd.get("IG Story Performance")
    ig_fol = rd.get("IG Followers")
    label = rd.label

    def _sum(df, col):
        if df.empty or col not in df.columns:
            return 0
        return int(pd.to_numeric(df[col].astype(str).str.replace(",", ""), errors="coerce").fillna(0).sum())

    total_posts = len(df_post)
    total_followers = 0
    net_growth = 0
    if not ig_fol.empty:
        if "Total Followers" in ig_fol.columns:
            total_followers = int(pd.to_numeric(ig_fol["Total Followers"], errors="coerce").fillna(0).iloc[0])
        if "Followers Net" in ig_fol.columns:
            net_growth = int(pd.to_numeric(ig_fol["Followers Net"], errors="coerce").fillna(0).sum())

    total_likes = _sum(df_post, "Likes")
    total_comments = _sum(df_post, "Comments")
    total_shares = _sum(df_post, "Shares")
    total_saves = _sum(df_post, "Saves")
    total_interaction = total_likes + total_comments + total_shares + total_saves
    total_reach = _sum(df_post, "Total Post Reach")

    total_stories = len(df_story)
    total_story_reach = _sum(df_story, "Total Reach")
    total_story_shares = _sum(df_story, "Shares")
    total_story_clicks = _sum(df_story, "Link clicks")

    d = total_posts if total_posts > 0 else 1
    ds = total_stories if total_stories > 0 else 1
    metrics = {
        "Total Followers": total_followers,
        "Net Followers Growth": net_growth,
        "Feeds Posted": total_posts,
        "Total Post Likes": total_likes,
        "Total Post Comments": total_comments,
        "Total Post Shares": total_shares,
        "Total Post Saves": total_saves,
        "Total Post Interaction": total_interaction,
        "Total Post Reach": total_reach,
        "Average Post Likes": int(round(total_likes / d)),
        "Average Post Comments": int(round(total_comments / d)),
        "Average Post Shares": int(round(total_shares / d)),
        "Average Post Saves": int(round(total_saves / d)),
        "Average Post Interaction": int(round(total_interaction / d)),
        "Average Post Reach": int(round(total_reach / d)),
        "Stories Posted": total_stories,
        "Total Stories Link Clicks": total_story_clicks,
        "Total Stories Shares": total_story_shares,
        "Total Stories Reach": total_story_reach,
        "Average Stories Link Clicks": int(round(total_story_clicks / ds)),
        "Average Stories Shares": int(round(total_story_shares / ds)),
        "Average Stories Reach": int(round(total_story_reach / ds)),
    }
    return [{"Metric": k, label: v} for k, v in metrics.items()]


def _api_followers_growth(rd) -> dict | None:
    df = rd.get("FB Followers")
    if df.empty or "Date" not in df.columns:
        return None
    dates = pd.to_datetime(df["Date"], format="%d/%m/%Y", errors="coerce")
    gain_col = "Followers Gain" if "Followers Gain" in df.columns else None
    loss_col = "Followers Loss" if "Followers Loss" in df.columns else None
    net_col = "Followers Net" if "Followers Net" in df.columns else None
    net_total = int(pd.to_numeric(df[net_col], errors="coerce").fillna(0).sum()) if net_col else 0
    return {
        "dates": dates.dt.strftime("%d/%m/%Y").tolist(),
        "gain": [int(x) for x in pd.to_numeric(df[gain_col], errors="coerce").fillna(0)] if gain_col else [],
        "loss": [int(x) for x in pd.to_numeric(df[loss_col], errors="coerce").fillna(0)] if loss_col else [],
        "net": [int(x) for x in pd.to_numeric(df[net_col], errors="coerce").fillna(0)] if net_col else [],
        "monthly_net": net_total,
    }


def _api_total_reach(rd) -> dict | None:
    df = rd.get("Unique Page View")
    if df.empty or "Date" not in df.columns:
        return None
    dates = pd.to_datetime(df["Date"], format="%d/%m/%Y", errors="coerce")
    reach_vals = pd.to_numeric(df["Unique Page View"], errors="coerce").fillna(0).astype(int)
    return {
        "dates": dates.dt.strftime("%d/%m/%Y").tolist(),
        "reach": reach_vals.tolist(),
        "monthly_total": int(reach_vals.sum()),
    }


def _api_reach_funnel(rd) -> dict | None:
    df = rd.get("FB Reach Funnel")
    if df is None or df.empty:
        return None
    organic = int(pd.to_numeric(df["Organic reach"], errors="coerce").fillna(0).sum())
    paid = int(pd.to_numeric(df["Paid reach"], errors="coerce").fillna(0).sum())
    return {"organic": organic, "paid": paid, "total": organic + paid}


def _api_ig_followers(rd) -> dict | None:
    df = rd.get("IG Followers")
    if df.empty or "Date" not in df.columns:
        return None
    dates = pd.to_datetime(df["Date"], format="%d/%m/%Y", errors="coerce")
    net_col = "Followers Net" if "Followers Net" in df.columns else None
    net_vals = pd.to_numeric(df[net_col], errors="coerce").fillna(0).astype(int) if net_col else pd.Series([0]*len(df))
    monthly_net = int(net_vals.sum())
    total_col = "Total Followers" if "Total Followers" in df.columns else None
    total_vals = pd.to_numeric(df[total_col], errors="coerce").fillna(0).astype(int).tolist() if total_col else []
    return {
        "dates": dates.dt.strftime("%d/%m/%Y").tolist(),
        "net": net_vals.tolist(),
        "monthly_net": monthly_net,
        "total": total_vals,
    }


@router.get("/api/periods")
def api_periods():
    from app_v2.db.range_reader import get_available_periods
    ds, de = default_range()
    return {
        "periods": get_available_periods(),
        "default_range": {"start": ds, "end": de},
    }


@router.get("/api/data/kpis")
def api_kpis(start: str = "", end: str = ""):
    s, e = _resolve_dates(start, end)
    rd = get_range_data(s, e)
    kpis = rd.kpis()
    return {"period": rd.label, "label": rd.label, **kpis}


@router.get("/api/data/fb_page")
def api_fb_page(start: str = "", end: str = ""):
    s, e = _resolve_dates(start, end)
    rd = get_range_data(s, e)

    comments = rd.get("Master Comments Base")
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

    return {
        "period": rd.label,
        "label": rd.label,
        "sentiment_distribution": sentiment_data,
        "sentiment_by_category": sentiment_by_cat,
        "fb_key_metrics": _compute_fb_key_metrics(rd),
        "followers_growth": _api_followers_growth(rd),
        "total_reach": _api_total_reach(rd),
        "reach_funnel": _api_reach_funnel(rd),
        "ig_followers": _api_ig_followers(rd),
    }


@router.get("/api/data/fb_posts")
def api_fb_posts(start: str = "", end: str = ""):
    s, e = _resolve_dates(start, end)
    rd = get_range_data(s, e)

    pillar = rd.get("FB Pivot (Pillar)")
    cat = rd.get("FB Pivot (Category)")
    ptype = rd.get("FB Pivot (Post Type)")
    wall = rd.get("FB Wall Post Performance")

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
        "Pillar Performance - GNC", "Pillar Performance - Branding",
        "Pillar Performance - Category", "Pillar Performance - Sales",
        "Pillar Performance - Others",
    ]
    perf_cols = ["Permalink", "Description", "Post type", "Pillar", "Category", "Sub-Category",
                 "Campaign Name", "Publish time", "Reactions", "Comments", "Shares",
                 "Interactions", "Reach", "Views", "Link Clicks", "Photo Clicks"]
    for sheet in breakdown_sheets:
        df = rd.get(sheet)
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
        "period": rd.label,
        "label": rd.label,
        "pillar_donut": pillar_donut,
        "type_donut": type_donut,
        "pillar_interactions": pillar_interactions,
        "cat_interactions": cat_interactions,
        "wall_posts": _df_to_records(wall),
        "breakdowns": breakdowns,
    }


@router.get("/api/data/instagram")
def api_instagram(start: str = "", end: str = ""):
    s, e = _resolve_dates(start, end)
    rd = get_range_data(s, e)

    ig_pivot = rd.get("IG Pivot (Pillar)")
    ig_story_pivot = rd.get("IG Story Pivot")
    ig_wall = rd.get("IG Wall Post Performance")

    ig_pillar_donut = []
    if not ig_pivot.empty and "Pillar" in ig_pivot.columns and "No. of Posts" in ig_pivot.columns:
        ig_pillar_donut = [{"name": str(r["Pillar"]), "value": _safe_int(r["No. of Posts"])} for _, r in ig_pivot.iterrows()]

    pillar_interactions = []
    if not ig_pivot.empty and "Pillar" in ig_pivot.columns and "Total Interactions" in ig_pivot.columns:
        pillar_interactions = [{"name": str(r["Pillar"]), "value": _safe_int(r["Total Interactions"])} for _, r in ig_pivot.iterrows()]

    story_clicks = []
    if not ig_story_pivot.empty and "Pillar" in ig_story_pivot.columns:
        click_col = "SUM of Link clicks" if "SUM of Link clicks" in ig_story_pivot.columns else ("Total Interactions" if "Total Interactions" in ig_story_pivot.columns else None)
        if click_col:
            df_sorted = ig_story_pivot.sort_values(click_col, ascending=False)
            story_clicks = [{"name": str(r["Pillar"]), "value": _safe_int(r[click_col])} for _, r in df_sorted.iterrows()]

    return {
        "period": rd.label,
        "label": rd.label,
        "ig_key_metrics": _compute_ig_key_metrics(rd),
        "pillar_donut": ig_pillar_donut,
        "pillar_interactions": pillar_interactions,
        "story_clicks": story_clicks,
        "wall_posts": _df_to_records(ig_wall),
        "ig_story_posts": _df_to_records(rd.get("IG Story Performance")),
        "ig_followers": _api_ig_followers(rd),
    }


@router.get("/api/data/fpk")
def api_fpk(start: str = "", end: str = "", platform: str = "fb"):
    """Return FPK competitor data — uses period-aligned months from the range."""
    s, e = _resolve_dates(start, end)
    start_dt = pd.to_datetime(s)
    end_dt = pd.to_datetime(e)

    base = Path(settings.competitors_dir)
    folder = base / f"{platform.upper()} competitor"
    pf = platform.upper()

    result = {"period": rd_label(s, e), "platform": platform}

    # Try each month in the range for FPK files (they're per-month)
    months = pd.date_range(start_dt.replace(day=1), end_dt, freq="MS")
    ym = f"{months[0].year}{months[0].month:02d}"  # Use first month

    # ── Growth Trend ──
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
                "key": comp["key"], "name": str(row.get("Profile", comp["key"])),
                "color": comp["color"], "logo": _logo_url(comp["logo"]),
                "data": [int(row[c]) if pd.notna(row[c]) else None for c in date_cols],
            })
        result["growth_trend"] = {"dates": [str(c) for c in date_cols], "series": series}

    # ── Competitors Overview ──
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
                "key": comp["key"], "name": str(row.get("Profile", comp["key"])),
                "color": comp["color"], "logo": _logo_url(comp["logo"]),
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


def rd_label(s: str, e: str) -> str:
    start = pd.to_datetime(s)
    end = pd.to_datetime(e)
    if start.strftime("%Y-%m") == end.strftime("%Y-%m"):
        return f"{MONTH_NAMES.get(start.month, '')} {start.year}"
    return f"{start.strftime('%d/%m/%Y')} – {end.strftime('%d/%m/%Y')}"


@router.get("/api/images")
def api_images(start: str = "", end: str = ""):
    """List FPK screenshots — uses first month in range."""
    import yaml
    s, e = _resolve_dates(start, end)
    start_dt = pd.to_datetime(s)

    period_dir = Path(settings.competitors_dir) / f"{start_dt.year}_{start_dt.month:02d}"
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
                "url": f"/competitors/{start_dt.year}_{start_dt.month:02d}/{filepath.name}",
                "label": cfg.get("label", stem.replace("_", " ").title()),
                "description": cfg.get("description", ""),
                "configured": bool(cfg),
            })
    return {"period": f"{start_dt.year}-{start_dt.month:02d}", "images": images}
