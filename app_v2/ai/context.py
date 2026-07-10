"""
Context builder — transforms PeriodData into a markdown summary for the LLM.

Gemini uses this context to answer user questions about the data.
"""

import pandas as pd
from app_v2.db.range_reader import RangeData as PeriodData, MONTH_NAMES


def build_context(pd_obj: PeriodData) -> str:
    lines = []
    period_label = pd_obj.label
    lines.append(f"# {period_label} 數據摘要\n")

    kpis = pd_obj.kpis()
    lines.append("## KPI 總覽")
    lines.append(f"- Facebook 粉絲數: {kpis['fb_followers']:,}")
    lines.append(f"- Facebook 粉絲增長: {kpis['fb_growth']:,}")
    lines.append(f"- Facebook 貼文數: {kpis['fb_wall_posts']}")
    lines.append(f"- Facebook 互動總數: {kpis['fb_interactions']:,}")
    lines.append(f"- Instagram 粉絲數: {kpis['ig_followers']:,}")
    lines.append(f"- Instagram 粉絲增長: {kpis['ig_growth']:,}")
    lines.append(f"- Instagram 觸及: {kpis['ig_reach']:,}")
    lines.append(f"- Instagram 互動: {kpis['ig_interactions']:,}")
    lines.append("")

    _add_sentiment(lines, pd_obj)
    _add_fb_pivot(lines, pd_obj, "FB Pivot (Pillar)", "Pillar", "Facebook Pillar 分析")
    _add_fb_pivot(lines, pd_obj, "FB Pivot (Category)", "Category", "Facebook 分類分析")
    _add_ig_section(lines, pd_obj)
    _add_breakdowns(lines, pd_obj)

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
    lines.append("## 留言情緒分析")
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
            lines.append("### 情緒 × 分類")
            for cat in pivot.index[:8]:
                parts = []
                for s in ["Positive", "Neutral", "Negative"]:
                    if s in pivot.columns:
                        parts.append(f"{s}={int(pivot.loc[cat, s])}")
                lines.append(f"- {cat}: {', '.join(parts)}")
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
            line_parts.append(f"貼文={posts}")
        if reach != "":
            line_parts.append(f"觸及={reach:,}")
        if interactions != "":
            line_parts.append(f"互動={interactions:,}")
        lines.append(" ".join(line_parts))
    lines.append("")


def _add_ig_section(lines: list[str], pd_obj: PeriodData):
    ig_key = pd_obj.get("IG Key Metrics")
    if ig_key.empty:
        return
    lines.append("## Instagram 關鍵指標")
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
        lines.append("### IG Pillar 分析")
        for _, row in ig_pivot.iterrows():
            name = str(row.get("Pillar", ""))
            interactions = int(row.get("Total Interactions", 0)) if "Total Interactions" in ig_pivot.columns else ""
            lines.append(f"- {name}: 互動={interactions:,}" if interactions != "" else f"- {name}")
        lines.append("")


def _add_breakdowns(lines: list[str], pd_obj: PeriodData):
    breakdown_sheets = [
        "Category Performance - BAU",
        "Sub-Category Performance - PNP",
        "Pillar Performance - CRM",
        "Pillar Performance - Ecommerce",
        "Pillar Performance - GNC",
    ]
    for sheet in breakdown_sheets:
        df = pd_obj.get(sheet)
        if df.empty:
            continue
        posts = len(df)
        interactions = int(df["Interactions"].sum()) if "Interactions" in df.columns else 0
        reach = int(df["Reach"].sum()) if "Reach" in df.columns else 0
        lines.append(f"- {sheet}: {posts} 貼文, 觸及={reach:,}, 互動={interactions:,}")


