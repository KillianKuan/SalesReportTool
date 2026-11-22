"""Tests for the chatbot tool-use loop (app/chatbot.py) using a fake client.

`python3 -m unittest discover -s tests` from the repo root.
"""
import os
import sys
import unittest
from pathlib import Path
from types import SimpleNamespace as NS
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

import chatbot  # noqa: E402


def _text(t):
    return NS(type="text", text=t)


def _tool_use(name, args, tid="tu_1"):
    return NS(type="tool_use", name=name, input=args, id=tid)


def _resp(blocks, stop):
    return NS(content=blocks, stop_reason=stop)


class FakeClient:
    """Returns scripted responses; records the create() calls."""

    def __init__(self, responses):
        self._responses = list(responses)
        self.calls = []
        self.messages = self

    def create(self, **kwargs):
        self.calls.append(kwargs)
        return self._responses.pop(0)


class AskTests(unittest.TestCase):
    def test_plain_answer_without_tools(self):
        client = FakeClient([_resp([_text("Hello")], "end_turn")])
        text, trace = chatbot.ask("hi", [], lambda n, a: {}, "k", client=client)
        self.assertEqual(text, "Hello")
        self.assertEqual(trace, [])

    def test_tool_round_trip_and_trace(self):
        client = FakeClient([
            _resp([_tool_use("get_kpis", {"year": 2025})], "tool_use"),
            _resp([_text("Revenue was 100.")], "end_turn"),
        ])
        seen = []

        def executor(name, args):
            seen.append((name, args))
            return {"revenue": 100}

        text, trace = chatbot.ask("rev?", [], executor, "k", client=client)
        self.assertEqual(text, "Revenue was 100.")
        self.assertEqual(seen, [("get_kpis", {"year": 2025})])
        self.assertEqual(trace[0]["output"], {"revenue": 100})
        # second request carries the tool_result back to the model
        last_user = client.calls[1]["messages"][-1]
        self.assertEqual(last_user["content"][0]["type"], "tool_result")
        self.assertEqual(last_user["content"][0]["tool_use_id"], "tu_1")

    def test_tool_error_flagged(self):
        client = FakeClient([
            _resp([_tool_use("query_sales", {})], "tool_use"),
            _resp([_text("No data.")], "end_turn"),
        ])
        chatbot.ask("x", [], lambda n, a: {"error": "boom"}, "k", client=client)
        result = client.calls[1]["messages"][-1]["content"][0]
        self.assertTrue(result.get("is_error"))

    def test_missing_key_raises_friendly_error(self):
        with self.assertRaises(chatbot.ChatError) as cm:
            chatbot.ask("hi", [], lambda n, a: {}, "")
        self.assertIn("API key", str(cm.exception))

    def test_loop_is_capped(self):
        looping = [_resp([_tool_use("get_kpis", {"year": 1})], "tool_use")
                   for _ in range(chatbot.MAX_TOOL_ROUNDS + 2)]
        with self.assertRaises(chatbot.ChatError):
            chatbot.ask("x", [], lambda n, a: {}, "k", client=FakeClient(looping))

    def test_api_exception_becomes_chat_error(self):
        class AuthenticationError(Exception):
            pass

        class Boom:
            messages = None

            def __init__(self):
                self.messages = self

            def create(self, **kw):
                raise AuthenticationError("401")

        with self.assertRaises(chatbot.ChatError) as cm:
            chatbot.ask("x", [], lambda n, a: {}, "k", client=Boom())
        self.assertIn("API key", str(cm.exception))

    def test_history_trimmed_and_starts_with_user(self):
        history = [{"role": "assistant", "content": "stale"}] + [
            {"role": "user" if i % 2 == 0 else "assistant", "content": f"m{i}"}
            for i in range(30)
        ]
        client = FakeClient([_resp([_text("ok")], "end_turn")])
        chatbot.ask("q", history, lambda n, a: {}, "k", client=client)
        msgs = client.calls[0]["messages"]
        self.assertEqual(msgs[0]["role"], "user")
        self.assertLessEqual(len(msgs), chatbot.MAX_HISTORY_MESSAGES + 1)
        self.assertEqual(msgs[-1], {"role": "user", "content": "q"})


class KeyAndPromptTests(unittest.TestCase):
    def test_env_var_wins_over_settings(self):
        with mock.patch.dict(os.environ, {chatbot.API_KEY_ENV: "env-key"}):
            self.assertEqual(chatbot.resolve_api_key({"chat_api_key": "file-key"}), "env-key")

    def test_settings_fallback_and_empty(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            self.assertEqual(chatbot.resolve_api_key({"chat_api_key": " file-key "}), "file-key")
            self.assertEqual(chatbot.resolve_api_key({}), "")
            self.assertEqual(chatbot.resolve_api_key(None), "")

    def test_system_prompt_has_date_and_rules(self):
        from datetime import datetime
        p = chatbot.build_system_prompt(datetime(2026, 5, 3))
        self.assertIn("2026-05-03", p)
        self.assertIn("TWD", p)
        self.assertIn("ALWAYS use the tools", p)


def _gem_text(t):
    return {"candidates": [{"content": {"role": "model", "parts": [{"text": t}]}}]}


def _gem_call(name, args):
    return {"candidates": [{"content": {"role": "model", "parts": [
        {"functionCall": {"name": name, "args": args}, "thoughtSignature": "sig"}]}}]}


class FakePost:
    def __init__(self, responses):
        self._responses = list(responses)
        self.calls = []

    def __call__(self, url, headers, body):
        # deep-ish copy: the loop mutates `contents` after the call
        import copy
        self.calls.append((url, headers, copy.deepcopy(body)))
        return self._responses.pop(0)


class GeminiAskTests(unittest.TestCase):
    def ask(self, post, **kw):
        return chatbot.ask("q", kw.pop("history", []), kw.pop("executor", lambda n, a: {}),
                           "gkey", client=post, provider="gemini", **kw)

    def test_plain_answer_and_request_shape(self):
        post = FakePost([_gem_text("Hi")])
        text, trace = self.ask(post, model="gemini-2.5-flash")
        self.assertEqual((text, trace), ("Hi", []))
        url, headers, body = post.calls[0]
        self.assertIn("gemini-2.5-flash:generateContent", url)
        self.assertEqual(headers["x-goog-api-key"], "gkey")
        names = [d["name"] for d in body["tools"][0]["functionDeclarations"]]
        self.assertEqual(names, ["query_sales", "get_kpis", "get_forecast"])
        self.assertIn("systemInstruction", body)

    def test_function_call_round_trip(self):
        post = FakePost([_gem_call("get_kpis", {"year": 2025}), _gem_text("Revenue 100")])
        seen = []

        def executor(name, args):
            seen.append((name, args))
            return {"revenue": 100}

        text, trace = self.ask(post, executor=executor)
        self.assertEqual(text, "Revenue 100")
        self.assertEqual(seen, [("get_kpis", {"year": 2025})])
        self.assertEqual(trace[0]["output"], {"revenue": 100})
        second = post.calls[1][2]["contents"]
        # model turn echoed unchanged (thought signature kept), then our response
        self.assertEqual(second[-2]["role"], "model")
        self.assertEqual(second[-2]["parts"][0]["thoughtSignature"], "sig")
        fr = second[-1]["parts"][0]["functionResponse"]
        self.assertEqual(fr["name"], "get_kpis")
        self.assertEqual(fr["response"]["result"], {"revenue": 100})

    def test_history_roles_mapped_to_user_and_model(self):
        post = FakePost([_gem_text("ok")])
        hist = [{"role": "user", "content": "a"}, {"role": "assistant", "content": "b"}]
        self.ask(post, history=hist)
        roles = [c["role"] for c in post.calls[0][2]["contents"]]
        self.assertEqual(roles, ["user", "model", "user"])

    def test_missing_key(self):
        with self.assertRaises(chatbot.ChatError) as cm:
            chatbot.ask("q", [], lambda n, a: {}, "", provider="gemini")
        self.assertIn("Gemini", str(cm.exception))

    def test_no_candidates_is_friendly_error(self):
        post = FakePost([{"promptFeedback": {"blockReason": "SAFETY"}}])
        with self.assertRaises(chatbot.ChatError) as cm:
            self.ask(post)
        self.assertIn("SAFETY", str(cm.exception))

    def test_loop_is_capped(self):
        post = FakePost([_gem_call("get_kpis", {"year": 1})
                         for _ in range(chatbot.MAX_TOOL_ROUNDS + 2)])
        with self.assertRaises(chatbot.ChatError):
            self.ask(post)

    def test_http_error_messages(self):
        f = chatbot._gemini_error_message
        self.assertIn("rate limit", f(429, "{}").lower())
        self.assertIn("API key", f(400, '{"error": {"message": "API key not valid."}}'))
        self.assertIn("not found", f(404, "{}"))


class ProviderConfigTests(unittest.TestCase):
    def test_defaults_to_anthropic(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            p, k, m = chatbot.provider_config({"chat_api_key": "a"})
        self.assertEqual((p, k, m), ("anthropic", "a", chatbot.DEFAULT_MODEL))

    def test_gemini_selected_uses_gemini_fields(self):
        s = {"chat_provider": "gemini", "chat_api_key": "a",
             "chat_gemini_api_key": "g", "chat_gemini_model": "gemini-2.5-pro"}
        with mock.patch.dict(os.environ, {}, clear=True):
            self.assertEqual(chatbot.provider_config(s), ("gemini", "g", "gemini-2.5-pro"))

    def test_gemini_env_vars(self):
        with mock.patch.dict(os.environ, {"GOOGLE_API_KEY": "envg"}, clear=True):
            self.assertEqual(chatbot.resolve_api_key({}, "gemini"), "envg")
        with mock.patch.dict(os.environ, {"GEMINI_API_KEY": "e1", "GOOGLE_API_KEY": "e2"}, clear=True):
            self.assertEqual(chatbot.resolve_api_key({}, "gemini"), "e1")

    def test_unknown_provider_falls_back(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            self.assertEqual(chatbot.provider_config({"chat_provider": "x"})[0], "anthropic")


if __name__ == "__main__":
    unittest.main()
