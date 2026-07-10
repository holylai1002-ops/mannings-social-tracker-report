"""
Chat API route for app_v2 — streaming LLM with date-range context.
"""

import logging
from fastapi import APIRouter
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

from app_v2.db.range_reader import get_range_data, default_range, MONTH_NAMES
from app_v2.ai.llm_client import stream_chat, stream_insights

logger = logging.getLogger(__name__)

router = APIRouter()


class ChatRequest(BaseModel):
    message: str
    history: list[dict] = []
    start: str = ""
    end: str = ""


class InsightsRequest(BaseModel):
    chart_title: str
    chart_type: str = ""
    data_summary: str


@router.post("/api/ai/chat")
async def chat(req: ChatRequest):
    s = req.start or ""
    e = req.end or ""
    if not s or not e:
        s, e = default_range()
    rd = get_range_data(s, e)

    messages = []
    for m in req.history[-6:]:
        role = "user" if m.get("role") == "user" else "model"
        messages.append({"role": role, "content": m.get("content", "")})
    messages.append({"role": "user", "content": req.message})

    async def generate():
        async for chunk in stream_chat(rd, messages):
            yield chunk

    return StreamingResponse(generate(), media_type="text/plain; charset=utf-8")


@router.post("/api/ai/insights")
async def insights(req: InsightsRequest):
    async def generate():
        async for chunk in stream_insights(req.chart_title, req.chart_type, req.data_summary):
            yield chunk

    return StreamingResponse(generate(), media_type="text/plain; charset=utf-8")
