"""
LLM client — streaming chat via OpenRouter (free models, works in Hong Kong).

Uses the OpenAI SDK pointed at OpenRouter's API.
Default model: deepseek/deepseek-chat-v3-0324:free
"""

import logging
from typing import AsyncGenerator

from app.config import settings
from app.ai.prompts import SYSTEM_PROMPT
from app.ai.context import build_context
from app.db.reader import PeriodData

logger = logging.getLogger(__name__)

_client = None


def get_client():
    """Lazily create the OpenAI client configured for OpenRouter."""
    global _client
    if _client is None:
        if not settings.openrouter_api_key:
            logger.warning("OPENROUTER_API_KEY not set — chatbot will return fallback text")
            return None
        from openai import AsyncOpenAI
        _client = AsyncOpenAI(
            api_key=settings.openrouter_api_key,
            base_url=settings.openrouter_base_url,
            default_headers={
                "HTTP-Referer": settings.openrouter_referer,
                "X-Title": settings.openrouter_title,
            },
            timeout=60.0,
        )
    return _client


async def stream_chat(
    pd_obj: PeriodData,
    messages: list[dict],
) -> AsyncGenerator[str, None]:
    """
    Stream chat response from OpenRouter.

    Args:
        pd_obj: PeriodData for the current period
        messages: [{"role": "user"/"assistant", "content": "..."}]

    Yields:
        Text chunks as they arrive
    """
    client = get_client()
    context = build_context(pd_obj)
    system = SYSTEM_PROMPT.format(period=pd_obj.period_str, context=context)

    if client is None:
        yield (
            "AI 聊天功能需要設定 OPENROUTER_API_KEY。\n\n"
            "1. 前往 https://openrouter.ai/keys 免費註冊並取得 API key\n"
            "2. 在 `.env` 檔案中設定 `OPENROUTER_API_KEY=你的key`\n"
            "3. 重啟伺服器"
        )
        return

    # Build OpenAI-style messages: system + conversation history
    api_messages = [{"role": "system", "content": system}]
    for m in messages:
        role = m.get("role", "user")
        if role == "user":
            api_messages.append({"role": "user", "content": m["content"]})
        else:
            api_messages.append({"role": "assistant", "content": m["content"]})

    try:
        response = await client.chat.completions.create(
            model=settings.openrouter_model,
            messages=api_messages,
            temperature=0.4,
            max_tokens=4000,
            stream=True,
        )
        async for chunk in response:
            delta = chunk.choices[0].delta if chunk.choices else None
            if delta and delta.content:
                yield delta.content
    except Exception as e:
        logger.error(f"OpenRouter streaming error: {e}")
        yield f"抱歉，生成回覆時發生錯誤：{e}"


INSIGHTS_PROMPT = """你是 Mannings（萬寧）社交媒體數據分析專家。請分析以下圖表/表格數據，用繁體中文提供：

**Key Takeaway**（2-3 個重點摘要）
**Actionable Insights**（2-3 個可執行的建議）

要求：
1. 所有數字必須來自提供的數據
2. 簡潔有力，用 markdown 列表格式
3. 總字數不超過 200 字

圖表名稱：{title}
圖表類型：{ctype}

數據摘要：
{summary}
"""


async def stream_insights(
    chart_title: str,
    chart_type: str,
    data_summary: str,
) -> AsyncGenerator[str, None]:
    """Stream AI insights for a specific chart/table."""
    client = get_client()
    if client is None:
        yield "AI 分析功能需要設定 OPENROUTER_API_KEY。"
        return

    prompt = INSIGHTS_PROMPT.format(title=chart_title, ctype=chart_type, summary=data_summary)

    try:
        response = await client.chat.completions.create(
            model=settings.openrouter_model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.4,
            max_tokens=600,
            stream=True,
        )
        async for chunk in response:
            delta = chunk.choices[0].delta if chunk.choices else None
            if delta and delta.content:
                yield delta.content
    except Exception as e:
        logger.error(f"OpenRouter insights error: {e}")
        yield f"抱歉，生成分析時發生錯誤：{e}"
