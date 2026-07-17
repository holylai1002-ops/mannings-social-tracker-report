"""
LLM client — streaming chat + insights via OpenRouter (free models, works in Hong Kong).

Uses the OpenAI SDK pointed at OpenRouter's API.
Primary model: nvidia/nemotron-3-super-120b-a12b:free (120B MoE, 12B active, 1M context)
Fallback chain: Tencent HY3 → Nemotron Nano 30B → Llama 3.3 70B → Gemma 4 31B →
                Gemma 4 26B A4B → Nemotron Nano 9B → Llama 3.2 3B
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
            timeout=90.0,
        )
    return _client


# Fallback chain ranked by suitability for Chinese/English social media analysis.
# Primary model is settings.openrouter_model (nemotron-3-super-120b).
FALLBACK_MODELS = [
    "tencent/hy3:free",
    "nvidia/nemotron-3-nano-30b-a3b:free",
    "meta-llama/llama-3.3-70b-instruct:free",
    "google/gemma-4-31b-it:free",
    "google/gemma-4-26b-a4b-it:free",
    "nvidia/nemotron-nano-9b-v2:free",
    "meta-llama/llama-3.2-3b-instruct:free",
]

_NO_KEY_MSG = (
    "AI features require an OpenRouter API key.\n\n"
    "1. Get a free key at https://openrouter.ai/keys\n"
    "2. Set OPENROUTER_API_KEY in your .env file\n"
    "3. Restart the server"
)


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
        yield _NO_KEY_MSG
        return

    api_messages = [{"role": "system", "content": system}]
    for m in messages:
        role = m.get("role", "user")
        if role == "user":
            api_messages.append({"role": "user", "content": m["content"]})
        else:
            api_messages.append({"role": "assistant", "content": m["content"]})

    models_to_try = [settings.openrouter_model] + FALLBACK_MODELS
    for i, model in enumerate(models_to_try):
        try:
            response = await client.chat.completions.create(
                model=model,
                messages=api_messages,
                temperature=0.4,
                max_tokens=4000,
                stream=True,
            )
            async for chunk in response:
                delta = chunk.choices[0].delta if chunk.choices else None
                if delta and delta.content:
                    yield delta.content
            return
        except Exception as e:
            logger.warning(f"Model {model} failed: {e}")
            if i < len(models_to_try) - 1:
                continue
            yield f"Sorry, an error occurred while generating a response: {e}"


INSIGHTS_PROMPT = """You are the Mannings social media analytics expert. Analyze the chart/table data below and provide insights in English.

**Key Takeaway** (2-3 main points)
**Actionable Insights** (2-3 actionable recommendations)

Rules:
1. All numbers must come from the provided data — do not fabricate
2. Be concise and impactful, use markdown bullet lists
3. Total response under 300 words

Chart name: {title}
Chart type: {ctype}

Data summary:
{summary}
"""


async def _try_stream(client, model, messages, temperature, max_tokens):
    """Attempt a streaming completion. Returns the streaming response object."""
    response = await client.chat.completions.create(
        model=model,
        messages=messages,
        temperature=temperature,
        max_tokens=max_tokens,
        stream=True,
    )
    return response


async def stream_insights(
    chart_title: str,
    chart_type: str,
    data_summary: str,
) -> AsyncGenerator[str, None]:
    """Stream AI insights for a specific chart/table."""
    client = get_client()
    if client is None:
        yield _NO_KEY_MSG
        return

    prompt = INSIGHTS_PROMPT.format(title=chart_title, ctype=chart_type, summary=data_summary)
    msg = [{"role": "user", "content": prompt}]

    models_to_try = [settings.openrouter_model] + FALLBACK_MODELS
    for i, model in enumerate(models_to_try):
        try:
            response = await _try_stream(client, model, msg, 0.4, 1500)
            async for chunk in response:
                delta = chunk.choices[0].delta if chunk.choices else None
                if delta and delta.content:
                    yield delta.content
            return
        except Exception as e:
            logger.warning(f"Model {model} failed: {e}")
            if i < len(models_to_try) - 1:
                continue
            yield f"Sorry, an error occurred while generating insights: {e}"
