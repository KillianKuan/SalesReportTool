"""Tests for the chatbot data tools (app/chat_tools.py).

Pure stdlib (unittest): `python3 -m unittest tests.test_chat_tools -v`
or `python3 -m unittest discover -s tests` from the repo root.
"""
import json
import sys
import unittest
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

import chat_tools  # noqa: E402
from utils import GP_COL  # noqa: E402


def _df():
    rows = [
        # date,        customer,    category,   qty, amt,  gp
        ("2025-01-15", "Acme Inc",  "Tablet",    10, 1000, 300),
        ("2025-02-15", "Acme Inc",  "CDR ACC",    5,  200,  50),
        ("2025-02-20", "Beta Co",   "Tablet",     2, 4000, 800),
        ("2024-03-15", "Acme Inc",  "Tablet",     4,  500, 100),
    ]
    df = pd.DataFrame(rows, columns=["Ship Date", "Customer Name", "Category", "QTY",
                                     "SALES Total AMT", GP_COL])
    df["Ship Date"] = pd.to_datetime(df["Ship Date"])
    df["Month"] = df["Ship Date"].dt.strftime("%Y-%m")
    df["Part Number"] = ["P1", "P2", "P1", "P1"]
    return df


class QuerySalesTests(unittest.TestCase):
    def test_top_customers_sorted_by_revenue(self):
        r = chat_tools.query_sales(_df(), group_by="customer", years=[2025])
        self.assertEqual([x["customer"] for x in r["rows"]], ["Beta Co", "Acme Inc"])
        self.assertEqual(r["rows"][0]["revenue"], 4000)
        self.assertEqual(r["rows"][1]["revenue"], 1200)

    def test_qty_counts_only_tablet_and_cdr(self):
        r = chat_tools.query_sales(_df(), group_by="customer", years=[2025],
                                   customer="acme")
        # 10 Tablet counted, 5 CDR ACC excluded
        self.assertEqual(r["rows"][0]["qty"], 10)

    def test_gp_pct(self):
        r = chat_tools.query_sales(_df(), group_by="none", years=[2025])
        self.assertAlmostEqual(r["rows"][0]["gp_pct"], 1150 / 5200 * 100, places=1)

    def test_month_grouping_is_chronological(self):
        r = chat_tools.query_sales(_df(), group_by="month")
        months = [x["month"] for x in r["rows"]]
        self.assertEqual(months, sorted(months))

    def test_category_filter_case_insensitive(self):
        r = chat_tools.query_sales(_df(), group_by="none", category="tablet", years=[2025])
        self.assertEqual(r["rows"][0]["revenue"], 5000)

    def test_unknown_customer_returns_error(self):
        r = chat_tools.query_sales(_df(), group_by="none", customer="nobody")
        self.assertIn("error", r)

    def test_invalid_group_by_returns_error(self):
        self.assertIn("error", chat_tools.query_sales(_df(), group_by="salary"))

    def test_top_n_capped_and_noted(self):
        r = chat_tools.query_sales(_df(), group_by="customer", top_n=1)
        self.assertEqual(len(r["rows"]), 1)
        self.assertIn("note", r)

    def test_no_match_returns_empty_rows(self):
        r = chat_tools.query_sales(_df(), group_by="customer", years=[1999])
        self.assertEqual(r["rows"], [])


class KpiAndExecutorTests(unittest.TestCase):
    def test_kpis_with_yoy(self):
        k = chat_tools.get_kpis(_df(), 2025)
        self.assertEqual(k["revenue"], 5200)
        self.assertAlmostEqual(k["revenue_yoy"], (5200 - 500) / 500 * 100, places=1)

    def test_kpis_missing_year(self):
        self.assertIn("error", chat_tools.get_kpis(_df(), 2030))

    def test_forecast_without_fcst(self):
        self.assertIn("error", chat_tools.get_forecast(_df(), None, 2025, 3))

    def test_executor_handles_bad_args_and_unknown_tool(self):
        run = chat_tools.make_executor(_df(), None, 2025, 3)
        self.assertIn("error", run("query_sales", {"bogus": 1}))
        self.assertIn("error", run("drop_tables", {}))
        self.assertIn("rows", run("query_sales", {"group_by": "category"}))

    def test_results_are_json_serialisable(self):
        run = chat_tools.make_executor(_df(), None, 2025, 3)
        json.dumps(run("query_sales", {"group_by": "part_number"}))
        json.dumps(run("get_kpis", {"year": 2025}))


class ShowChartTests(unittest.TestCase):
    def test_returns_chart_spec_and_rows(self):
        r = chat_tools.show_chart(_df(), "bar", "customer", years=[2025])
        self.assertEqual(r["chart"]["type"], "bar")
        self.assertEqual(r["chart"]["points"][0], {"label": "Beta Co", "value": 4000})

    def test_month_line_is_chronological(self):
        r = chat_tools.show_chart(_df(), "line", "month", metric="gp")
        labels = [p["label"] for p in r["chart"]["points"]]
        self.assertEqual(labels, sorted(labels))
        self.assertTrue(r["chart"]["chronological"])

    def test_rejects_bad_input(self):
        self.assertIn("error", chat_tools.show_chart(_df(), "radar", "customer"))
        self.assertIn("error", chat_tools.show_chart(_df(), "bar", "none"))
        self.assertIn("error", chat_tools.show_chart(_df(), "bar", "customer", metric="foo"))

    def test_no_rows_has_no_chart(self):
        r = chat_tools.show_chart(_df(), "bar", "customer", years=[1999])
        self.assertNotIn("chart", r)

    def test_executor_dispatch_and_renderer(self):
        run = chat_tools.make_executor(_df(), None, 2025, 6)
        out = run("show_chart", {"chart_type": "pie", "group_by": "category"})
        self.assertIn("chart", out)
        import charts
        for t in ("bar", "line", "pie"):
            spec = dict(out["chart"], type=t)
            charts.chat_chart(spec).to_dict()


if __name__ == "__main__":
    unittest.main()
