"""Tests for dashboard insights (app/insights.py): rules, AI narration, cache.

`python3 -m unittest discover -s tests` from the repo root.
"""
import json
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace as NS

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

import chatbot  # noqa: E402
import insights  # noqa: E402

TODAY = datetime(2026, 11, 24)  # Jan-Oct 2026 are completed months


def _rows(year, month_rev, customer, cat="Tablet", gp_pct=30.0):
    return [
        {"Ship Date": pd.Timestamp(year, m, 28), "Customer Name": customer, "Category": cat,
         "SALES Total AMT": rev, insights.GP: rev * gp_pct / 100}
        for m, rev in month_rev.items()
    ]


def _sample_df():
    flat = {m: 1_000_000 for m in range(1, 13)}
    cur = {m: 1_000_000 for m in range(1, 11)}
    cur.update({7: 500_000, 8: 500_000, 10: 600_000})  # drops: Jul-Aug, Oct
    rows = _rows(2025, flat, "ACME") + _rows(2025, flat, "BETA", cat="CDR")
    rows += _rows(2026, cur, "ACME")
    rows += _rows(2026, {m: 1_000_000 for m in range(1, 11)}, "BETA", cat="CDR")
    return pd.DataFrame(rows)


class FakeClient:
    """Anthropic-like client returning one scripted text answer (or raising)."""

    def __init__(self, text=None, exc=None):
        self.text, self.exc, self.calls = text, exc, 0
        self.messages = self

    def create(self, **kwargs):
        self.calls += 1
        if self.exc:
            raise self.exc
        return NS(content=[NS(type="text", text=self.text)], stop_reason="end_turn")


class DetectTests(unittest.TestCase):
    def setUp(self):
        self.findings = insights.detect_findings(_sample_df(), 2026, TODAY, "en")

    def test_clusters_july_august_and_october_separately(self):
        ids = [f["id"] for f in self.findings if f["id"].startswith("drop_")]
        self.assertCountEqual(ids, ["drop_2026_07_08", "drop_2026_10_10"])
        titles = " ".join(f["title"] for f in self.findings)
        self.assertIn("Jul–Aug", titles)
        self.assertIn("Oct", titles)

    def test_zh_titles_use_month_labels(self):
        zh = insights.detect_findings(_sample_df(), 2026, TODAY, "zh-TW")
        self.assertTrue(any("7–8月" in f["title"] for f in zh))

    def test_ranked_by_severity_then_impact(self):
        order = [(insights._SEV_RANK[f["severity"]], -abs(f["impact"])) for f in self.findings]
        self.assertEqual(order, sorted(order))

    def test_gross_margin_drop_flagged(self):
        df = _sample_df()
        low = (df["Ship Date"].dt.year == 2026) & (df["Ship Date"].dt.month == 10)
        df.loc[low, insights.GP] = df.loc[low, "SALES Total AMT"] * 0.10
        out = insights.detect_findings(df, 2026, TODAY, "en")
        self.assertTrue(any(f["section"] == "margin" and "Oct" in f["title"] for f in out))

    def test_shrinking_top_customer_flagged(self):
        df = _sample_df()
        df = df[~((df["Customer Name"] == "BETA") & (df["Ship Date"].dt.year == 2026))]
        out = insights.detect_findings(df, 2026, TODAY, "en")
        beta = [f for f in out if f["id"] == "cust_2026_BETA"]
        self.assertEqual(len(beta), 1)
        self.assertEqual(beta[0]["severity"], "critical")

    def test_budget_pace_and_forecast_gap(self):
        budget = {"ytd_rev": 10_000_000, "fy_budget": 40_000_000, "fy_forecast": 30_000_000}
        out = insights.detect_findings(_sample_df(), 2026, TODAY, "en", budget)
        ids = {f["id"] for f in out}
        self.assertIn("pace_2026", ids)
        self.assertIn("fc_2026", ids)

    def test_month_the_data_has_not_finished_is_not_judged(self):
        df = _sample_df()
        # Data stops on Oct 10 (stale file): October is partial, not a collapse.
        df = df[~((df["Ship Date"].dt.year == 2026) & (df["Ship Date"].dt.month == 10))]
        df = pd.concat([df, pd.DataFrame(_rows(2026, {10: 100_000}, "ACME")).assign(
            **{"Ship Date": pd.Timestamp(2026, 10, 10)})])
        self.assertEqual(insights._completed_months(df, 2026, TODAY), 9)
        out = insights.detect_findings(df, 2026, TODAY, "en")
        self.assertFalse(any(f["id"] == "drop_2026_10_10" for f in out))

    def test_no_findings_for_empty_or_unstarted_year(self):
        self.assertEqual(insights.detect_findings(_sample_df().iloc[0:0], 2026, TODAY), [])
        self.assertEqual(insights.detect_findings(_sample_df(), 2027, TODAY), [])


class AiAndCacheTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.cache = Path(self.tmp.name) / "cache.json"
        insights._failed.clear()
        self.df = _sample_df()
        self.settings = {"chat_api_key": "k", "insights_lang": "en", "insights_refresh": "weekly"}
        first_id = insights.detect_findings(self.df, 2026, TODAY, "en")[0]["id"]
        self.payload = json.dumps({
            "headline": "Revenue is steady but Jul-Aug and Oct dipped.",
            "top3": [{"id": first_id, "text": "Look here first."}, {"id": "bogus", "text": "x"}],
            "sections": {"trend": "Dips in Jul, Aug, Oct."},
        })

    def tearDown(self):
        self.tmp.cleanup()

    def _get(self, client, **kw):
        return insights.get_insights(
            self.df, 2026, settings=self.settings, today=TODAY,
            client=client, cache_path=self.cache, **kw)

    def test_ai_summary_used_and_unknown_ids_dropped(self):
        client = FakeClient(self.payload)
        res = self._get(client)
        self.assertEqual(res["source"], "ai")
        self.assertEqual(len(res["summary"]["top3"]), 1)
        self.assertEqual(res["summary"]["sections"]["trend"], "Dips in Jul, Aug, Oct.")
        self.assertEqual(client.calls, 1)

    def test_sections_the_model_left_empty_fall_back_to_rule_text(self):
        res = self._get(FakeClient(self.payload))
        rule = insights.rule_summary(res["findings"], insights.build_facts(self.df, 2026, TODAY), "en")
        self.assertEqual(res["summary"]["sections"]["customers"], rule["sections"]["customers"])

    def test_second_call_hits_cache_not_api(self):
        client = FakeClient(self.payload)
        self._get(client)
        self._get(client)
        self.assertEqual(client.calls, 1)

    def test_freshness_by_schedule(self):
        later = datetime(2026, 11, 30)
        old = {"generated_at": "2026-11-01T00:00:00"}
        new = {"generated_at": "2026-11-25T00:00:00"}
        self.assertFalse(insights.is_fresh(old, "weekly", later))
        self.assertTrue(insights.is_fresh(new, "weekly", later))
        self.assertTrue(insights.is_fresh(old, "monthly", later))
        self.assertTrue(insights.is_fresh(old, "manual", later))
        self.assertFalse(insights.is_fresh({}, "weekly", later))

    def test_stale_cache_triggers_regeneration(self):
        client = FakeClient(self.payload)
        self._get(client)
        cache = insights.load_cache(self.cache)
        for entry in cache.values():
            entry["generated_at"] = "2026-10-01T00:00:00"
        insights.save_cache(cache, self.cache)
        self._get(client)
        self.assertEqual(client.calls, 2)

    def test_api_failure_falls_back_to_rules_and_cools_down(self):
        client = FakeClient(exc=RuntimeError("boom"))
        res = self._get(client)
        self.assertEqual(res["source"], "rules")
        self.assertTrue(res["error"])
        self.assertTrue(res["summary"]["headline"])
        self._get(client)
        self.assertEqual(client.calls, 1)  # cooldown: no hammering
        self._get(client, force=True)
        self.assertEqual(client.calls, 2)  # Regenerate overrides the cooldown

    def test_unparseable_reply_is_a_friendly_failure(self):
        res = self._get(FakeClient("not json at all"))
        self.assertEqual(res["source"], "rules")
        self.assertIn("could not be parsed", res["error"])

    def test_no_key_gives_rule_based_summary_without_api(self):
        res = insights.get_insights(
            self.df, 2026, settings={"insights_lang": "en"}, today=TODAY,
            cache_path=self.cache)
        self.assertEqual(res["source"], "rules")
        self.assertIsNone(res["error"])
        self.assertTrue(res["summary"]["top3"])

    def test_json_in_code_fence_is_parsed(self):
        res = self._get(FakeClient("```json\n" + self.payload + "\n```"))
        self.assertEqual(res["source"], "ai")


class GenerateTextTests(unittest.TestCase):
    def test_gemini_path_returns_text_and_requests_json(self):
        seen = {}

        def post(url, headers, body):
            seen["body"] = body
            return {"candidates": [{"content": {"parts": [{"text": "hi"}]}}]}

        out = chatbot.generate_text("p", "s", "key", provider="gemini", client=post, json_mode=True)
        self.assertEqual(out, "hi")
        self.assertEqual(seen["body"]["generationConfig"]["responseMimeType"], "application/json")

    def test_empty_answer_raises_chat_error(self):
        with self.assertRaises(chatbot.ChatError):
            chatbot.generate_text("p", "s", "key", client=FakeClient(""))


if __name__ == "__main__":
    unittest.main()
