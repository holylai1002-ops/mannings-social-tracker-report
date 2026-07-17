"""
Context builder — transforms PeriodData into a markdown summary for the LLM.
Includes competitor data from separate Excel files.
"""

import warnings
import pandas as pd
from pathlib import Path

from app.config import settings
from app.db.reader import PeriodData, MONTH_NAMES


def build_context(pd_obj: PeriodData) -> str:
    lines = []
    period_label = f"{MONTH_NAMES.get(pd_obj.month, '')} {pd_obj.year}"
    lines.append(f"# {period_label} Data Summary\n")

    kpis = pd_obj.kpis()
    lines.append("## KPI Overview")
    lines.append(f"- Facebook Followers: {kpis['fb_followers']:,}")
    lines.append(f"- Facebook Follower Growth: {kpis['fb_growth']:,}")
    lines.append(f"- Facebook Posts: {kpis['fb_wall_posts']}")
    lines.append(f"- Facebook Interactions: {kpis['fb_interactions']:,}")
    lines.append(f"- Instagram Followers: {kpis['ig_followers']:,}")
    lines.append(f"- Instagram Follower Growth: {kpis['ig_growth']:,}")
    lines.append(f"- Instagram Reach: {kpis['ig_reach']:,}")
    lines.append(f"- Instagram Interactions: {kpis['ig_interactions']:,}")
    lines.append("")

    _add_sentiment(lines, pd_obj)
    _add_fb_pivot(lines, pd_obj, "FB Pivot (Pillar)", "Pillar", "Facebook Pillar Breakdown")
    _add_fb_pivot(lines, pd_obj, "FB Pivot (Category)", "Category", "Facebook Category Breakdown")
    _add_fb_pivot(lines, pd_obj, "FB Pivot (Post Type)", "Post type", "Facebook Post Type Breakdown")
    _add_ig_section(lines, pd_obj)
    _add_breakdowns(lines, pd_obj)
    _add_competitor_section(lines, pd_obj)

    return "\n".join(lines)


def _add_sentiment(lines: list[str], pd_obj: PeriodData):
    comments = pd_obj.get("Master Comments Base")
    if comments.empty or "Sentiment" not in comments.columns:
        return
    sent = comments["Sentiment"].astype(str).str.strip().str.title().replace("N/a", "N/A")
    sent = sent[sent.isin(["Positive", "Neutral", "Negative", "N/A"])]
    if sent.empty:
        return
    counts = sent.value_counts()
    lines.append("## Comment Sentiment Analysis")
    for s in ["Positive", "Neutral", "Negative", "N/A"]:
        if s in counts.index:
            lines.append(f"- {s}: {counts[s]}")
    lines.append("")

    if "Category" in comments.columns:
        cat_sent = comments.copy()
        cat_sent["Sentiment"] = cat_sent["Sentiment"].astype(str).str.strip().str.title()
        cat_sent = cat_sent[cat_sent["Sentiment"].isin(["Positive", "Neutral", "Negative"])]
        if not cat_sent.empty:
            pivot = cat_sent.groupby("Category")["Sentiment"].value_counts().unstack(fill_value=0)
            lines.append("### Sentiment by Category")
            for cat in pivot.index[:8]:
                parts = []
                for s in ["Positive", "Neutral", "Negative"]:
                    if s in pivot.columns:
                        parts.append(f"{s}={int(pivot.loc[cat, s])}")
                lines.append(f"- {cat}: {', '.join(parts)}")
            lines.append("")

    if "Type" in comments.columns:
        type_sent = comments.copy()
        type_sent["Sentiment"] = type_sent["Sentiment"].astype(str).str.strip().str.title()
        type_sent = type_sent[type_sent["Sentiment"].isin(["Positive", "Neutral", "Negative"])]
        if not type_sent.empty:
            pivot = type_sent.groupby("Type")["Sentiment"].value_counts().unstack(fill_value=0)
            lines.append("### Sentiment by Comment Type")
            for tp in pivot.index[:6]:
                parts = []
                for s in ["Positive", "Neutral", "Negative"]:
                    if s in pivot.columns:
                        parts.append(f"{s}={int(pivot.loc[tp, s])}")
                lines.append(f"- {tp}: {', '.join(parts)}")
            lines.append("")


def _add_fb_pivot(lines: list[str], pd_obj: PeriodData, sheet: str, label_col: str, title: str):
    df = pd_obj.get(sheet)
    if df.empty:
        return
    lines.append(f"## {title}")
    int_col = "Total Interactions"
    posts_col = "No. of Posts"
    for _, row in df.iterrows():
        name = str(row.get(label_col, ""))
        posts = int(row.get(posts_col, 0)) if posts_col in df.columns else ""
        interactions = int(row.get(int_col, 0)) if int_col in df.columns else ""
        reach = int(row.get("Total Reach", 0)) if "Total Reach" in df.columns else ""
        line_parts = [f"- {name}"]
        if posts != "":
            line_parts.append(f"posts={posts}")
        if reach != "":
            line_parts.append(f"reach={reach:,}")
        if interactions != "":
            line_parts.append(f"interactions={interactions:,}")
        lines.append(" ".join(line_parts))
    lines.append("")


def _add_ig_section(lines: list[str], pd_obj: PeriodData):
    ig_key = pd_obj.get("IG Key Metrics")
    if ig_key.empty:
        return
    lines.append("## Instagram Key Metrics")
    period_col = pd_obj.period_str if pd_obj.period_str in ig_key.columns else ig_key.columns[-1]
    if "Metric" in ig_key.columns:
        for _, row in ig_key.iterrows():
            metric = str(row.get("Metric", ""))
            val = row.get(period_col, "")
            try:
                val = int(val) if pd.notna(val) else "N/A"
            except (ValueError, TypeError):
                val = str(val) if pd.notna(val) else "N/A"
            lines.append(f"- {metric}: {val}")
    lines.append("")

    ig_pivot = pd_obj.get("IG Pivot (Pillar)")
    if not ig_pivot.empty:
        lines.append("### IG Pillar Breakdown")
        for _, row in ig_pivot.iterrows():
            name = str(row.get("Pillar", ""))
            interactions = int(row.get("Total Interactions", 0)) if "Total Interactions" in ig_pivot.columns else ""
            lines.append(f"- {name}: interactions={interactions:,}" if interactions != "" else f"- {name}")
        lines.append("")


def _add_breakdowns(lines: list[str], pd_obj: PeriodData):
    breakdown_sheets = [
        "Category Performance - BAU",
        "Sub-Category Performance - PNP",
        "Pillar Performance - CRM",
        "Pillar Performance - Ecommerce",
        "Pillar Performance - GNC",
    ]
    lines.append("## Category/Pillar Performance")
    for sheet in breakdown_sheets:
        df = pd_obj.get(sheet)
        if df.empty:
            continue
        posts = len(df)
        interactions = int(df["Interactions"].sum()) if "Interactions" in df.columns else 0
        reach = int(df["Reach"].sum()) if "Reach" in df.columns else 0
        lines.append(f"- {sheet}: {posts} posts, reach={reach:,}, interactions={interactions:,}")
    lines.append("")


def _add_competitor_section(lines: list[str], pd_obj: PeriodData):
    """Read competitor data from Excel files and add to context."""
    base = Path(settings.competitors_dir)
    y, m = pd_obj.year, pd_obj.month
    ym = f"{y}{m:02d}"

    has_data = False
    for platform in ["FB", "IG"]:
        folder = base / f"{platform} competitor"
        co_file = folder / f"{platform} Competitors Overview_{ym}.xlsx"

        if not co_file.exists():
            continue

        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                df = pd.read_excel(co_file, header=4)
            df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

            if not has_data:
                lines.append("## Competitor Analysis")
                has_data = True

            lines.append(f"### {platform} Competitors Overview")
            keep_cols = [c for c in [
                "Profile", "Follower", "Follower Growth (absolute)",
                "Follower Growth (in %)", "Number of posts",
                "Reactions, Comments & Shares"
            ] if c in df.columns]

            for _, row in df.iterrows():
                profile = str(row.get("Profile", "")).strip()
                if not profile or profile == "nan":
                    continue
                parts = []
                for c in keep_cols:
                    if c == "Profile":
                        continue
                    val = row.get(c, "")
                    if pd.notna(val) and str(val).strip() not in ("", "nan"):
                        parts.append(f"{c}={val}")
                if parts:
                    lines.append(f"- {profile}: {', '.join(parts)}")
            lines.append("")
        except Exception:
            pass

        gt_file = folder / f"{platform} Competitor Fans Growth Trend_{ym}.xlsx"
        if gt_file.exists():
            try:
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    df_gt = pd.read_excel(gt_file, header=4)
                df_gt = df_gt.loc[:, ~df_gt.columns.astype(str).str.startswith("Unnamed")]

                date_cols = [c for c in df_gt.columns if _is_date_col(c)]
                if date_cols:
                    lines.append(f"### {platform} Competitor Fan Growth Trend (monthly snapshots)")
                    for _, row in df_gt.iterrows():
                        profile = str(row.get("Profile", "")).strip()
                        if not profile or profile == "nan":
                            continue
                        vals = []
                        for c in date_cols:
                            v = row.get(c)
                            if pd.notna(v):
                                vals.append(str(int(v)))
                        if vals:
                            first_val = vals[0]
                            last_val = vals[-1]
                            try:
                                growth = int(last_val) - int(first_val)
                                lines.append(f"- {profile}: start={first_val}, end={last_val}, growth={growth:+d} ({len(vals)} data points)")
                            except ValueError:
                                lines.append(f"- {profile}: {' -> '.join(vals[:3])} ... {' -> '.join(vals[-2:])}")
                    lines.append("")
            except Exception:
                pass

    if not has_data:
        lines.append("## Competitor Analysis")
        lines.append("- No competitor data files found for this period.")
        lines.append("")


def _is_date_col(col) -> bool:
    """Detect date columns like 'Jun 1, 2026', '2026-01-15', '1/5/2026'."""
    import re
    s = str(col)
    has_date_sep = any(sep in s for sep in ["/", "-", ","])
    has_digit = any(c.isdigit() for c in s)
    has_month = bool(re.search(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)", s))
    return (has_date_sep and has_digit) or has_month
