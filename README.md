---
title: Mannings Social Dashboard
emoji: 📊
colorFrom: blue
colorTo: indigo
sdk: docker
app_port: 7860
pinned: false
---

# Mannings Social Media Dashboard

FastAPI + Jinja2 + HTMX + Alpine.js + ECharts dashboard for Mannings Facebook & Instagram analytics.

## Features
- FB Page: Sentiment analysis, Key Metrics, Competitor Analysis
- FB Posts: Posts by Pillar/Type, Interactions, Wall Post Performance
- Instagram: Posts by Pillar, Engagement, Stories, Competitor Analysis
- AI Insights: Per-chart Key Takeaway + Actionable Insights (OpenRouter)
- Excel export on all tables/charts
- Print to PDF (A4 landscape)

## Setup

1. Set the following Secrets in HF Space Settings:
   - `OPENROUTER_API_KEY` — your OpenRouter API key (get free key at https://openrouter.ai/keys)
   - `OPENROUTER_MODEL` — model name (default: `openai/gpt-oss-120b:free`)

2. The app starts automatically on port 7860.

## Local Development

```bash
pip install -r requirements_deploy.txt
python -m uvicorn app.main:app --host 127.0.0.1 --port 8000 --reload
```
