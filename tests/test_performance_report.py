"""Tests for the Performance Report workflow: single-year filtering,
unconditional category output, combined ACC merging, and the report-state
options tuple used for stale-result detection.

Pure stdlib (unittest) - no pytest dependency:
`python3 -m unittest tests.test_performance_report -v` from the repo root,
or `python3 -m unittest discover -s tests`.
"""
import sys
import unittest
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

import utils  # noqa: E402


def _make_df(rows):
    """Build a minimal report-ready DataFrame from (year, month, category,
    qty, amt, gp) tuples."""
    return pd.DataFrame({
        "Ship Date": pd.to_datetime([f"{y}-{m.split('-')[1]}-15" for y, m, *_ in rows]),
        "Month": [m for _, m, *_ in rows],
        "Customer Name": ["Acme"] * len(rows),
        "Category": [c for _, _, c, *_ in rows],
        "QTY": [q for *_, q, a, g in rows],
        "SALES Total AMT": [a for *_, q, a, g in rows],
        utils.GP_COL: [g for *_, q, a, g in rows],
        "Part Number": [f"P{i}" for i in range(len(rows))],
    })


class SingleYearFilteringTestCase(unittest.TestCase):
    """The Performance Report now filters by exactly one selected year
    (a scalar equality check), replacing the old multiselect .isin() call."""

    def setUp(self):
        self.df = _make_df([
            (2023, "2023-12", "CDR", 1, 10, 2),
            (2024, "2024-01", "CDR", 2, 20, 4),
            (2024, "2024-02", "Tablet", 3, 30, 6),
        ])

    def test_scalar_year_filter_keeps_only_that_year(self):
        filtered = self.df[self.df["Ship Date"].dt.year == 2024].copy()
        self.assertEqual(set(filtered["Ship Date"].dt.year.unique()), {2024})
        self.assertEqual(len(filtered), 2)

    def test_scalar_year_filter_excludes_other_years(self):
        filtered = self.df[self.df["Ship Date"].dt.year == 2023].copy()
        self.assertEqual(len(filtered), 1)
        self.assertNotIn(2024, filtered["Ship Date"].dt.year.unique())

    def test_year_not_in_data_yields_empty(self):
        filtered = self.df[self.df["Ship Date"].dt.year == 1999].copy()
        self.assertTrue(filtered.empty)


class UnconditionalCategoryOutputTestCase(unittest.TestCase):
    """By Category results (and the ByCategory export sheet) are now always
    generated for any valid (non-empty) report run - no `by_cat` flag."""

    def test_build_bycat_nonempty_for_single_category_base(self):
        base = _make_df([
            (2024, "2024-01", "CDR", 5, 50, 10),
        ])
        long_bycat = utils.build_bycat(base, qty_only=True, merge_acc=True)
        self.assertFalse(long_bycat.empty)
        self.assertIn("CDR", long_bycat["Category"].tolist())

    def test_build_bycat_nonempty_for_multi_category_base(self):
        base = _make_df([
            (2024, "2024-01", "CDR", 5, 50, 10),
            (2024, "2024-01", "Signify", 1, 5, 1),
            (2024, "2024-02", "Others", 2, 8, 1),
        ])
        long_bycat = utils.build_bycat(base, qty_only=False, merge_acc=True)
        self.assertFalse(long_bycat.empty)
        self.assertEqual(
            set(long_bycat["Category"].unique()), {"CDR", "Signify", "Others"}
        )

    def test_sorted_cats_and_wide_export_always_produce_rows(self):
        base = _make_df([(2024, "2024-01", "AI_SW", 1, 100, 20)])
        long_bycat = utils.build_bycat(base, qty_only=False, merge_acc=True)
        all_months = sorted(long_bycat["Month"].unique().tolist())
        frames = [
            utils.to_wide_one_cat(long_bycat, cat, all_months)
            for cat in utils.sorted_cats(long_bycat)
        ]
        combined = pd.concat(frames, ignore_index=True)
        self.assertFalse(combined.empty)


class MergeAccTestCase(unittest.TestCase):
    """A single `merge_acc` flag replaces the separate CDR/Tablet ACC merge
    checkboxes and must merge both pairs together, summing QTY/AMT/GP."""

    def setUp(self):
        self.base = _make_df([
            (2024, "2024-01", "CDR", 10, 100, 20),
            (2024, "2024-01", "CDR ACC", 3, 30, 6),
            (2024, "2024-01", "Tablet", 5, 50, 10),
            (2024, "2024-01", "Tablet ACC", 2, 20, 4),
        ])

    def test_merge_acc_true_combines_cdr_and_tablet_acc(self):
        long_bycat = utils.build_bycat(self.base, qty_only=False, merge_acc=True)
        cats = set(long_bycat["Category"].unique())
        self.assertEqual(cats, {"CDR", "Tablet"})

        cdr_row = long_bycat[long_bycat["Category"] == "CDR"].iloc[0]
        self.assertEqual(cdr_row["SALES Total AMT"], 130)
        self.assertEqual(cdr_row["final GP(NTD)"], 26)

        tablet_row = long_bycat[long_bycat["Category"] == "Tablet"].iloc[0]
        self.assertEqual(tablet_row["SALES Total AMT"], 70)
        self.assertEqual(tablet_row["final GP(NTD)"], 14)

    def test_merge_acc_false_keeps_categories_separate(self):
        long_bycat = utils.build_bycat(self.base, qty_only=False, merge_acc=False)
        cats = set(long_bycat["Category"].unique())
        self.assertEqual(cats, {"CDR", "CDR ACC", "Tablet", "Tablet ACC"})

    def test_merge_acc_true_qty_only_excludes_acc_from_qty(self):
        long_bycat = utils.build_bycat(self.base, qty_only=True, merge_acc=True)
        cdr_row = long_bycat[long_bycat["Category"] == "CDR"].iloc[0]
        # qty_only masks QTY to rows whose *original* Category was CDR/Tablet,
        # so the merged-in ACC qty (3) must not be counted.
        self.assertEqual(cdr_row["QTY (All)"], 10)


class ReportOptionsStateTestCase(unittest.TestCase):
    """The `_opts` tuple app.py caches as `rpt_opts` drives the stale-result
    warning; it must change whenever any control (qty_only, merge_acc,
    year, or the selected customers) changes."""

    @staticmethod
    def _opts(qty_only, merge_acc, year, selected):
        return (qty_only, merge_acc, year, tuple(sorted(selected)))

    def test_identical_controls_produce_equal_opts(self):
        a = self._opts(True, True, 2024, ["Acme", "Beta"])
        b = self._opts(True, True, 2024, ["Beta", "Acme"])
        self.assertEqual(a, b)

    def test_year_change_is_detected_as_stale(self):
        a = self._opts(True, True, 2024, ["Acme"])
        b = self._opts(True, True, 2023, ["Acme"])
        self.assertNotEqual(a, b)

    def test_qty_only_change_is_detected_as_stale(self):
        a = self._opts(True, True, 2024, ["Acme"])
        b = self._opts(False, True, 2024, ["Acme"])
        self.assertNotEqual(a, b)

    def test_merge_acc_change_is_detected_as_stale(self):
        a = self._opts(True, True, 2024, ["Acme"])
        b = self._opts(True, False, 2024, ["Acme"])
        self.assertNotEqual(a, b)

    def test_selected_customers_change_is_detected_as_stale(self):
        a = self._opts(True, True, 2024, ["Acme"])
        b = self._opts(True, True, 2024, ["Acme", "Beta"])
        self.assertNotEqual(a, b)


class StyleReportTableTestCase(unittest.TestCase):
    """The reusable table styling helper left-aligns value columns and
    themes headers via CSS variables (not hard-coded colors)."""

    def test_style_report_table_left_aligns_value_columns_only(self):
        df = pd.DataFrame({
            "Metric": ["QTY (All)", "GP%"],
            "2024-01": [10, "20.0%"],
        })
        styled = utils.style_report_table(df)
        styled._compute()
        ctx = styled._translate(False, False)
        cell_props = [p for entry in ctx.get("cellstyle", []) for p in entry["props"]]
        self.assertIn(("text-align", "left"), cell_props)

    def test_style_report_table_header_uses_theme_css_variables(self):
        df = pd.DataFrame({"Metric": ["QTY (All)"], "2024-01": [10]})
        styled = utils.style_report_table(df)
        ctx = styled._translate(False, False)
        header_props = [p for s in ctx.get("table_styles", []) for p in s["props"]]
        css_values = [str(v) for _, v in header_props]
        self.assertTrue(any("var(--" in v for v in css_values))
        self.assertFalse(any(v.strip().startswith("#") for v in css_values))


if __name__ == "__main__":
    unittest.main()
