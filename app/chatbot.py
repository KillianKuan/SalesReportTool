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

# Google AI Studio (Gemini API): free tier needs no credit card. Called over
# plain REST with urllib, so no extra dependency / packaging flags.
PROVIDERS = {"anthropic": "Anthropic (Claude)", "gemini": "Google AI Studio (Gemini)"}
DEFAULT_PROVIDER = "anthropic"
DEFAULT_GEMINI_MODEL = "gemini-3.5-flash"
# Verified against the live API (generateContent). "-latest" aliases follow
# Google's current model but can be briefly overloaded (503).
GEMINI_MODELS = ["gemini-3.5-flash", "gemini-3.5-flash-lite", "gemini-3.1-flash-lite",
                 "gemini-flash-latest"]
# Google stops serving old models to new users (404 "no longer available").
# A saved choice from this list silently falls back to the default.
GEMINI_RETIRED_MODELS = frozenset({"gemini-2.5-flash", "gemini-2.5-flash-lite", "gemini-2.5-pro"})


def normalize_gemini_model(model: Optional[str]) -> str:
    """Saved Gemini model, or the default if empty / known to be retired."""
    model = (model or "").strip()
    return DEFAULT_GEMINI_MODEL if not model or model in GEMINI_RETIRED_MODELS else model
GEMINI_API_KEY_ENVS = ("GEMINI_API_KEY", "GOOGLE_API_KEY")
GEMINI_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
GEMINI_TIMEOUT = 60

EXAMPLE_QUESTIONS = [
    "What was total revenue and GP% in 2025, and how does it compare to 2024?",
    "Who are the top 5 customers by revenue this year?",
    "Show monthly Tablet revenue for 2025.",
    "What is the full-year forecast vs budget?",
]


def resolve_api_key(settings: Optional[dict] = None, provider: str = DEFAULT_PROVIDER) -> str:
    """Environment variable wins; fall back to the key saved in settings.json."""
    if provider == "gemini":
        key = next(
            (v for v in (os.environ.get(e, "").strip() for e in GEMINI_API_KEY_ENVS) if v), ""
        )
        field = "chat_gemini_api_key"
    else:
        key = os.environ.get(API_KEY_ENV, "").strip()
        field = "chat_api_key"
    if not key and settings:
        key = str(settings.get(field) or "").strip()
    return key


def provider_config(settings: Optional[dict] = None) -> tuple[str, str, str]:
    """(provider, api_key, model) for the provider selected in settings."""
    settings = settings or {}
    provider = settings.get("chat_provider")
    if provider not in PROVIDERS:
        provider = DEFAULT_PROVIDER
    if provider == "gemini":
        model = normalize_gemini_model(settings.get("chat_gemini_model"))
    else:
        model = settings.get("chat_model") or DEFAULT_MODEL
    return provider, resolve_api_key(settings, provider), model


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
    model: Optional[str] = None,
    client: Any = None,
    provider: str = DEFAULT_PROVIDER,
) -> tuple[str, list[dict]]:
    """Answer *question* with the chosen provider. Returns (answer_text, tool_trace).

    *history* is prior plain-text turns [{"role": "user"|"assistant", "content": str}].
    *tool_trace* lists {"tool", "input", "output"} for each tool call, so the UI
    can show where the numbers came from. Raises ChatError on failure.
    *client* is a test seam: an Anthropic-like client, or for Gemini a
    callable (url, headers, body) -> dict.
    """
    if provider == "gemini":
        return _ask_gemini(question, history, executor, api_key,
                           model or DEFAULT_GEMINI_MODEL, client)
    return _ask_anthropic(question, history, executor, api_key,
                          model or DEFAULT_MODEL, client)


def _gemini_error_message(status: int, body: str) -> str:
    detail = ""
    try:
        detail = json.loads(body).get("error", {}).get("message", "")
    except (ValueError, AttributeError):
        pass
    if status in (400, 401, 403) and "api key" in (detail or "").lower():
        return "The Gemini API key was rejected. Check it in Settings ▸ Chatbot."
    if status in (401, 403):
        return "The Gemini API key is not allowed to use this model or region."
    if status == 404:
        if "no longer available" in (detail or "").lower():
            return ("The selected Gemini model has been retired by Google. Pick a newer one "
                    "in Settings ▸ Chatbot.")
        return "The selected Gemini model was not found. Pick another one in Settings ▸ Chatbot."
    if status == 429:
        return ("Gemini rate limit reached (the free tier has low per-minute and daily "
                "limits). Wait a moment and try again.")
    if status == 400:
        return f"The request was rejected: {detail or 'bad request'}"
    return f"Gemini returned an error (HTTP {status}). Please try again."


def _gemini_post(url: str, headers: dict, body: dict) -> dict:
    import urllib.error
    import urllib.request

    req = urllib.request.Request(
        url, data=json.dumps(body).encode("utf-8"), headers=headers, method="POST"
    )
    try:
        with urllib.request.urlopen(req, timeout=GEMINI_TIMEOUT) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        raise ChatError(_gemini_error_message(exc.code, exc.read().decode("utf-8", "replace"))) from exc
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        raise ChatError(
            "Could not reach the Gemini API. Check your internet connection."
        ) from exc


def _ask_gemini(question, history, executor, api_key, model, post=None):
    if not api_key and post is None:
        raise ChatError(
            "No Gemini API key configured. Add one in Settings ▸ Chatbot, or set the "
            f"{GEMINI_API_KEY_ENVS[0]} environment variable."
        )
    post = post or _gemini_post
    url = GEMINI_ENDPOINT.format(model=model)
    headers = {"Content-Type": "application/json", "x-goog-api-key": api_key}
    tools = [{"functionDeclarations": [
        {"name": t["name"], "description": t["description"], "parameters": t["input_schema"]}
        for t in TOOL_SCHEMAS
    ]}]

    contents: list[dict] = [
        {"role": "user" if m["role"] == "user" else "model", "parts": [{"text": m["content"]}]}
        for m in history[-MAX_HISTORY_MESSAGES:]
        if m.get("role") in ("user", "assistant") and m.get("content")
    ]
    while contents and contents[0]["role"] != "user":
        contents.pop(0)
    contents.append({"role": "user", "parts": [{"text": question}]})

    trace: list[dict] = []
    system = {"parts": [{"text": build_system_prompt()}]}
    for _ in range(MAX_TOOL_ROUNDS + 1):
        body = {
            "systemInstruction": system,
            "contents": contents,
            "tools": tools,
            "generationConfig": {"maxOutputTokens": MAX_TOKENS * 2},
        }
        try:
            data = post(url, headers, body)
        except ChatError:
            raise
        except Exception as exc:
            _logger.warning("Gemini request failed", exc_info=True)
            raise ChatError("The assistant ran into an error. Please try again.") from exc

        cands = data.get("candidates") or []
        if not cands:
            block = (data.get("promptFeedback") or {}).get("blockReason")
            raise ChatError(
                f"Gemini returned no answer{f' (blocked: {block})' if block else ''}. "
                "Try rephrasing the question."
            )
        content = cands[0].get("content") or {}
        parts = content.get("parts") or []
        calls = [p["functionCall"] for p in parts if "functionCall" in p]
        if not calls:
            text = "".join(p.get("text", "") for p in parts if not p.get("thought")).strip()
            return (text or "(No answer returned.)"), trace

        # Echo the model turn back unchanged (keeps any thought signatures).
        contents.append({"role": "model", "parts": parts})
        responses = []
        for call in calls:
            args = dict(call.get("args") or {})
            output = executor(call["name"], args)
            trace.append({"tool": call["name"], "input": args, "output": output})
            responses.append({"functionResponse": {
                "name": call["name"], "response": {"result": output},
            }})
        contents.append({"role": "user", "parts": responses})

    raise ChatError(
        "That question needed too many lookups. Try asking something simpler or more specific."
    )


def _ask_anthropic(question, history, executor, api_key, model, client):
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
