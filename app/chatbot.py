"""chatbot.py - Claude tool-use loop for the in-app data chatbot.

The model never sees raw rows: it calls the read-only tools in chat_tools.py and
phrases the returned aggregates. No Streamlit imports, so it is unit-testable
with a fake client.
"""
from __future__ import annotations

import json
import logging
import os
from datetime import datetime
from typing import Any, Callable, Optional

from chat_tools import TOOL_SCHEMAS

_logger = logging.getLogger(__name__)

DEFAULT_MODEL = "claude-haiku-5-5"
CHAT_MODELS = ["claude-haiku-5-5", "claude-sonnet-5-5"]
MAX_TOOL_ROUNDS = 5
MAX_HISTORY_MESSAGES = 10
MAX_TOKENS = 1024
API_KEY_ENV = "ANTHROPIC_API_KEY"

EXAMPLE_QUESTIONS = [
    "What was total revenue and GP% in 2025, and how does it compare to 2024?",
    "Who are the top 5 customers by revenue this year?",
    "Show monthly Tablet revenue for 2025.",
    "What is the full-year forecast vs budget?",
]


def resolve_api_key(settings: Optional[dict] = None) -> str:
    """Environment variable wins; fall back to the key saved in settings.json."""
    key = os.environ.get(API_KEY_ENV, "").strip()
    if not key and settings:
        key = str(settings.get("chat_api_key") or "").strip()
    return key


def build_system_prompt(today: Optional[datetime] = None) -> str:
    today = today or datetime.now()
    return (
        "You are the data assistant inside the Sales Report Tool. Answer short, factual "
        "questions about the company's sales data.\n"
        f"Today is {today:%Y-%m-%d}; the fiscal year equals the calendar year.\n"
        "Rules:\n"
        "- ALWAYS use the tools for any number. Never compute or recall figures yourself.\n"
        "- Amounts are in TWD. Say so, and format large numbers readably "
        "(e.g. 12.3M TWD or NT$12,345,678).\n"
        "- Categories: Tablet, CDR, Tablet ACC, CDR ACC, AI_SW, Signify, Others. "
        "qty only counts CDR and Tablet units.\n"
        "- 'This year' means the current calendar year. Shipping data is actuals; "
        "get_forecast blends actuals (past months) with FCST (current month onward) and Budget.\n"
        "- If a customer name matches several customers, say which ones you used.\n"
        "- If a tool returns an error or no rows, say so plainly; do not guess.\n"
        "- If the question is not about this sales data, politely say you can only answer "
        "questions about it.\n"
        "- Keep answers concise. Use a small markdown table for rankings or trends."
    )


class ChatError(Exception):
    """User-presentable failure (missing key, network, rate limit...)."""


def _friendly_error(exc: Exception) -> str:
    name = type(exc).__name__
    if name == "AuthenticationError":
        return "The API key was rejected. Check it in Settings ▸ Chatbot."
    if name == "PermissionDeniedError":
        return "This API key is not allowed to use that model."
    if name == "NotFoundError":
        return "The selected model was not found. Pick another one in Settings ▸ Chatbot."
    if name == "RateLimitError":
        return "Rate limit reached. Please wait a moment and try again."
    if name in ("APIConnectionError", "APITimeoutError"):
        return "Could not reach the Anthropic API. Check your internet connection."
    if name == "BadRequestError":
        return f"The request was rejected: {exc}"
    return f"The assistant ran into an error ({name}). Please try again."


def _make_client(api_key: str):
    try:
        import anthropic
    except ImportError as exc:  # packaging problem / dependency not installed
        raise ChatError(
            "The 'anthropic' package is not installed. Run `pip install anthropic`."
        ) from exc
    return anthropic.Anthropic(api_key=api_key, timeout=60.0, max_retries=2)


def ask(
    question: str,
    history: list[dict],
    executor: Callable[[str, dict], dict],
    api_key: str,
    model: str = DEFAULT_MODEL,
    client: Any = None,
) -> tuple[str, list[dict]]:
    """Answer *question*. Returns (answer_text, tool_trace).

    *history* is prior plain-text turns [{"role": "user"|"assistant", "content": str}].
    *tool_trace* lists {"tool", "input", "output"} for each tool call, so the UI
    can show where the numbers came from. Raises ChatError on failure.
    """
    if not api_key and client is None:
        raise ChatError(
            "No API key configured. Add one in Settings ▸ Chatbot, or set the "
            f"{API_KEY_ENV} environment variable."
        )
    if client is None:
        client = _make_client(api_key)

    messages: list[dict] = [
        {"role": m["role"], "content": m["content"]}
        for m in history[-MAX_HISTORY_MESSAGES:]
        if m.get("role") in ("user", "assistant") and m.get("content")
    ]
    # The API requires the first message to be from the user.
    while messages and messages[0]["role"] != "user":
        messages.pop(0)
    messages.append({"role": "user", "content": question})

    trace: list[dict] = []
    system = build_system_prompt()
    try:
        for _ in range(MAX_TOOL_ROUNDS + 1):
            resp = client.messages.create(
                model=model,
                max_tokens=MAX_TOKENS,
                system=system,
                tools=TOOL_SCHEMAS,
                messages=messages,
            )
            if resp.stop_reason != "tool_use":
                text = "".join(
                    b.text for b in resp.content if getattr(b, "type", "") == "text"
                ).strip()
                return (text or "(No answer returned.)"), trace

            messages.append({"role": "assistant", "content": resp.content})
            results = []
            for block in resp.content:
                if getattr(block, "type", "") != "tool_use":
                    continue
                output = executor(block.name, dict(block.input or {}))
                trace.append({"tool": block.name, "input": dict(block.input or {}),
                              "output": output})
                results.append({
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": json.dumps(output, ensure_ascii=False, default=str),
                    **({"is_error": True} if isinstance(output, dict) and "error" in output else {}),
                })
            messages.append({"role": "user", "content": results})
    except ChatError:
        raise
    except Exception as exc:
        _logger.warning("Chatbot request failed", exc_info=True)
        raise ChatError(_friendly_error(exc)) from exc

    raise ChatError(
        "That question needed too many lookups. Try asking something simpler or more specific."
    )
