"""
Chat API route — streaming LLM chatbot + per-chart AI insights.
"""

import json
import logging
from fastapi import APIRouter, Request
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

from app.db.reader import get_period_data, default_period, MONTH_NAMES
from app.ai.llm_client import stream_chat, stream_insights

logger = logging.getLogger(__name__)

router = APIRouter()


class ChatRequest(BaseModel):
    message: str
    history: list[dict] = []
    year: int = 0
    month: int = 0


class InsightsRequest(BaseModel):
    chart_title: str
    chart_type: str = ""
    data_summary: str
    year: int = 0
    month: int = 0


@router.post("/api/ai/chat")
async def chat(req: ChatRequest):
    if req.year == 0 or req.month == 0:
        year, month = default_period()
    else:
        year, month = req.year, req.month

    pd_obj = get_period_data(year, month)

    messages = []
    for m in req.history[-6:]:
        role = "user" if m.get("role") == "user" else "model"
        messages.append({"role": role, "content": m.get("content", "")})
    messages.append({"role": "user", "content": req.message})

    async def generate():
        async for chunk in stream_chat(pd_obj, messages):
            yield chunk

    return StreamingResponse(generate(), media_type="text/plain; charset=utf-8")


@router.post("/api/ai/insights")
async def insights(req: InsightsRequest):
    """Generate Key Takeaway + Actionable Insights for a specific chart/table."""
    async def generate():
        async for chunk in stream_insights(req.chart_title, req.chart_type, req.data_summary):
            yield chunk

    return StreamingResponse(generate(), media_type="text/plain; charset=utf-8")
