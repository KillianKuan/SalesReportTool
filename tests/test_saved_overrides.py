"""Tests for the saved-overrides manager: override_key_id, set_override,
reset_override, and the unchanged QTY eligibility rule for rows whose
Category came from a saved override.

Pure stdlib (unittest) - no pytest dependency:
`python3 -m unittest tests.test_saved_overrides -v` from the repo root,
or `python3 -m unittest discover -s tests`.
"""
import sys
import unittest
import unittest.mock
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

import utils  # noqa: E402


def _key(customer="Acme", pn="P1", month="2024-01", des=""):
    return utils.override_key(customer, pn, month, des)


class OverrideKeyIdTestCase(unittest.TestCase):
    """override_key_id() is the smallest reusable representation of an
    override key used for both display and Streamlit widget keys."""

    def test_deterministic_for_same_key(self):
        k = _key()
        self.assertEqual(utils.override_key_id(k), utils.override_key_id(k))

    def test_differs_for_different_keys(self):
        self.assertNotEqual(
            utils.override_key_id(_key(pn="P1")),
            utils.override_key_id(_key(pn="P2")),
        )

    def test_safe_for_fields_that_would_collide_under_naive_join(self):
        # A naive "__".join() id (as the existing Others-panel widget keys
        # use) collides here; the hash-based id must not.
        a = utils.override_key_id(("Acme__X", "", "2024-01", ""))
        b = utils.override_key_id(("Acme", "X", "2024-01", ""))
        self.assertNotEqual(a, b)


class OverrideAssignmentOptionsTestCase(unittest.TestCase):
    """The saved-overrides manager's category selector must offer exactly
    the same options as the existing Others-panel reassignment control."""

    def test_starts_with_reset_sentinel(self):
        opts = utils.override_assignment_options()
        self.assertEqual(opts[0], utils.RESET_OVERRIDE_CHOICE)

    def test_excludes_others_and_matches_cat_order(self):
        opts = utils.override_assignment_options()
        self.assertNotIn("Others", opts)
        self.assertEqual(
            set(opts[1:]), {c for c in utils.CAT_ORDER if c != "Others"}
        )


class SetOverrideTestCase(unittest.TestCase):
    """set_override() backs both the Others-panel selector and the saved-
    overrides manager's per-row editor: it must touch only the targeted key
    and persist through the existing save_overrides() helper."""

    def setUp(self):
        self.overrides = {}
        patcher = unittest.mock.patch.object(utils, "save_overrides")
        self.mock_save = patcher.start()
        self.addCleanup(patcher.stop)

    def test_new_assignment_is_added_and_saved(self):
        k = _key()
        changed = utils.set_override(self.overrides, k, "CDR")
        self.assertTrue(changed)
        self.assertEqual(self.overrides[k], "CDR")
        self.mock_save.assert_called_once_with(self.overrides)

    def test_editing_only_updates_the_targeted_key(self):
        k1, k2 = _key(pn="P1"), _key(pn="P2")
        self.overrides[k1] = "CDR"
        self.overrides[k2] = "Tablet"
        utils.set_override(self.overrides, k1, "Signify")
        self.assertEqual(self.overrides[k1], "Signify")
        self.assertEqual(self.overrides[k2], "Tablet")

    def test_reassigning_to_the_same_category_is_a_noop(self):
        k = _key()
        self.overrides[k] = "CDR"
        self.mock_save.reset_mock()
        changed = utils.set_override(self.overrides, k, "CDR")
        self.assertFalse(changed)
        self.mock_save.assert_not_called()

    def test_reset_sentinel_removes_the_entry(self):
        k = _key()
        self.overrides[k] = "CDR"
        changed = utils.set_override(self.overrides, k, utils.RESET_OVERRIDE_CHOICE)
        self.assertTrue(changed)
        self.assertNotIn(k, self.overrides)
        self.mock_save.assert_called_once_with(self.overrides)

    def test_reset_on_a_key_not_present_is_a_noop(self):
        k = _key()
        changed = utils.set_override(self.overrides, k, utils.RESET_OVERRIDE_CHOICE)
        self.assertFalse(changed)
        self.mock_save.assert_not_called()


class ResetOverrideTestCase(unittest.TestCase):
    """The manager's per-row Reset action deletes only that override and
    persists the change immediately."""

    def setUp(self):
        self.overrides = {}
        patcher = unittest.mock.patch.object(utils, "save_overrides")
        self.mock_save = patcher.start()
        self.addCleanup(patcher.stop)

    def test_deletes_only_the_targeted_override(self):
        k1, k2 = _key(pn="P1"), _key(pn="P2")
        self.overrides[k1] = "CDR"
        self.overrides[k2] = "Tablet"
        changed = utils.reset_override(self.overrides, k1)
        self.assertTrue(changed)
        self.assertNotIn(k1, self.overrides)
        self.assertEqual(self.overrides[k2], "Tablet")

    def test_persists_via_the_existing_save_overrides_helper(self):
        k = _key()
        self.overrides[k] = "CDR"
        utils.reset_override(self.overrides, k)
        self.mock_save.assert_called_once_with(self.overrides)


class QtyEligibilityUnaffectedByOverridesTestCase(unittest.TestCase):
    """Overrides only change which Category a row is classified into before
    build_bycat() runs; the QTY-only inclusion rule itself must be
    unchanged: CDR ACC / Tablet ACC stay excluded from QTY even after being
    merged into CDR / Tablet, and a row an override reassigned into CDR is
    counted exactly like a row that was natively CDR."""

    def _make_df(self, rows):
        return pd.DataFrame({
            "Ship Date": pd.to_datetime(
                [f"2024-{m.split('-')[1]}-15" for m, *_ in rows]
            ),
            "Month": [m for m, *_ in rows],
            "Customer Name": ["Acme"] * len(rows),
            "Category": [c for _, c, *_ in rows],
            "QTY": [q for *_, q, a, g in rows],
            "SALES Total AMT": [a for *_, q, a, g in rows],
            utils.GP_COL: [g for *_, q, a, g in rows],
            "Part Number": [f"P{i}" for i in range(len(rows))],
        })

    def test_override_reassigned_row_counts_like_a_native_category_row(self):
        # As in app.py, an override is applied to the Category column
        # before build_bycat() ever runs, so a reassigned "Others" row and
        # a natively-CDR row must produce identical QTY output.
        overridden = self._make_df([("2024-01", "CDR", 4, 40, 8)])
        native = self._make_df([("2024-01", "CDR", 4, 40, 8)])
        overridden_out = utils.build_bycat(overridden, qty_only=True, merge_acc=True)
        native_out = utils.build_bycat(native, qty_only=True, merge_acc=True)
        pd.testing.assert_frame_equal(
            overridden_out.reset_index(drop=True), native_out.reset_index(drop=True)
        )

    def test_acc_still_excluded_from_qty_after_merge_regardless_of_overrides(self):
        base = self._make_df([
            ("2024-01", "CDR", 10, 100, 20),
            ("2024-01", "CDR ACC", 3, 30, 6),
        ])
        long_bycat = utils.build_bycat(base, qty_only=True, merge_acc=True)
        cdr_row = long_bycat[long_bycat["Category"] == "CDR"].iloc[0]
        self.assertEqual(cdr_row["QTY (All)"], 10)
        # Sales/GP are still merged even though QTY excludes the ACC portion.
        self.assertEqual(cdr_row["SALES Total AMT"], 130)


if __name__ == "__main__":
    unittest.main()
