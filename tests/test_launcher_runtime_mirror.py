"""Tests for launcher.py's atomic macOS runtime-app mirroring.

Pure stdlib (unittest) - no pytest dependency, so this runs anywhere with
python3: `python3 -m unittest tests.test_launcher_runtime_mirror -v` from the
repo root, or `python3 -m unittest discover -s tests`.

These exercise the platform-neutral core (`mirror_app_dir()` and its
helpers) directly with tmp directories, without touching Application
Support, ``sys.platform``, or PyInstaller-only state - so they run the same
on Linux/CI as on a real Mac.
"""
import contextlib
import json
import shutil
import sys
import tempfile
import unittest
import unittest.mock
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import launcher  # noqa: E402


REQUIRED_FILES = launcher.REQUIRED_RUNTIME_FILES  # ("app.py", "utils.py", "charts.py", "fcst_loader.py")


def _write_bundle(root: Path, version: str = "v1", extra_files: dict | None = None) -> None:
    """Create a fake bundled app/ directory under root with the required
    runtime files, an aliases.json (bundled default), and placeholder
    overrides.json/settings.json (as ship in the real repo)."""
    root.mkdir(parents=True, exist_ok=True)
    for name in REQUIRED_FILES:
        (root / name).write_text(f"# {name} {version}\n", encoding="utf-8")
    (root / "aliases.json").write_text(
        json.dumps({"customer": {}, "fcst_customer": {f"bundled-{version}": "Bundled Co."}}),
        encoding="utf-8",
    )
    (root / "overrides.json").write_text("[]", encoding="utf-8")
    (root / "settings.json").write_text("{}", encoding="utf-8")
    for name, content in (extra_files or {}).items():
        (root / name).write_text(content, encoding="utf-8")


class MirrorAppDirTestCase(unittest.TestCase):
    def setUp(self):
        self._tmp = Path(self.enterContext(_temp_dir()))
        self.source_dir = self._tmp / "bundle" / "app"
        self.active_dir = self._tmp / "support" / "app"
        _write_bundle(self.source_dir, "v1")

    # ---- first launch ----------------------------------------------------

    def test_first_launch_creates_active_dir_from_bundle(self):
        result = launcher.mirror_app_dir(self.source_dir, self.active_dir)

        self.assertEqual(result, self.active_dir)
        for name in REQUIRED_FILES:
            self.assertTrue((self.active_dir / name).is_file())
        self.assertEqual(
            json.loads((self.active_dir / "aliases.json").read_text())["fcst_customer"],
            {"bundled-v1": "Bundled Co."},
        )
        # No leftover staging/backup directories after a successful mirror.
        self.assertFalse((self.active_dir.parent / "app.staging").exists())
        self.assertFalse((self.active_dir.parent / "app.previous").exists())

    def test_first_launch_missing_required_file_raises(self):
        (self.source_dir / "utils.py").unlink()

        with self.assertRaises(launcher.RuntimeMirrorError):
            launcher.mirror_app_dir(self.source_dir, self.active_dir)

        # No half-written active dir left behind.
        self.assertFalse(self.active_dir.exists())

    # ---- restart (no bundle change) ---------------------------------------

    def test_restart_is_idempotent(self):
        launcher.mirror_app_dir(self.source_dir, self.active_dir)
        (self.active_dir / "settings.json").write_text(
            json.dumps({"customer_aliases": {"Foo": "Foo Inc."}}), encoding="utf-8"
        )

        result = launcher.mirror_app_dir(self.source_dir, self.active_dir)

        self.assertEqual(result, self.active_dir)
        self.assertEqual(
            json.loads((self.active_dir / "settings.json").read_text())["customer_aliases"],
            {"Foo": "Foo Inc."},
        )
        self.assertFalse((self.active_dir.parent / "app.staging").exists())
        self.assertFalse((self.active_dir.parent / "app.previous").exists())

    # ---- upgrade ------------------------------------------------------------

    def test_upgrade_refreshes_bundled_files_and_preserves_user_state(self):
        launcher.mirror_app_dir(self.source_dir, self.active_dir)

        custom_settings = {
            "customer_aliases": {"Weird Customer LLC": "Normalized Co."},
            "fcst_customer_aliases": {"FCST Weird": "Normalized Co."},
        }
        (self.active_dir / "settings.json").write_text(json.dumps(custom_settings), encoding="utf-8")
        (self.active_dir / "overrides.json").write_text(
            json.dumps([[["Cust", "PN", "2024-01", ""], 42]]), encoding="utf-8"
        )

        # Ship a new bundle version with updated code and updated defaults.
        _write_bundle(self.source_dir, "v2")

        launcher.mirror_app_dir(self.source_dir, self.active_dir)

        self.assertIn("v2", (self.active_dir / "app.py").read_text())
        self.assertEqual(
            json.loads((self.active_dir / "aliases.json").read_text())["fcst_customer"],
            {"bundled-v2": "Bundled Co."},
        )
        # User overrides survive the upgrade unchanged.
        self.assertEqual(
            json.loads((self.active_dir / "settings.json").read_text()), custom_settings
        )
        self.assertEqual(
            json.loads((self.active_dir / "overrides.json").read_text()),
            [[["Cust", "PN", "2024-01", ""], 42]],
        )

    def test_deleted_bundled_file_is_removed_from_active_copy(self):
        _write_bundle(self.source_dir, "v1", extra_files={"legacy_module.py": "# old\n"})
        launcher.mirror_app_dir(self.source_dir, self.active_dir)
        self.assertTrue((self.active_dir / "legacy_module.py").exists())

        # New version drops the module entirely.
        shutil.rmtree(self.source_dir)
        _write_bundle(self.source_dir, "v2")

        launcher.mirror_app_dir(self.source_dir, self.active_dir)

        self.assertFalse((self.active_dir / "legacy_module.py").exists())
        self.assertIn("v2", (self.active_dir / "app.py").read_text())

    # ---- copy failure --------------------------------------------------------

    def test_copy_failure_keeps_previous_valid_copy(self):
        launcher.mirror_app_dir(self.source_dir, self.active_dir)
        original_app_py = (self.active_dir / "app.py").read_text()

        _write_bundle(self.source_dir, "v2")

        def _boom(*_args, **_kwargs):
            raise OSError("simulated disk failure")

        with unittest.mock.patch.object(launcher.shutil, "copytree", side_effect=_boom):
            result = launcher.mirror_app_dir(self.source_dir, self.active_dir)

        # Falls back to the previous valid copy rather than raising.
        self.assertEqual(result, self.active_dir)
        self.assertEqual((self.active_dir / "app.py").read_text(), original_app_py)
        self.assertFalse((self.active_dir.parent / "app.staging").exists())
        self.assertFalse((self.active_dir.parent / "app.previous").exists())

    def test_copy_failure_on_first_launch_raises_actionable_error(self):
        def _boom(*_args, **_kwargs):
            raise OSError("simulated disk failure")

        with unittest.mock.patch.object(launcher.shutil, "copytree", side_effect=_boom):
            with self.assertRaises(launcher.RuntimeMirrorError):
                launcher.mirror_app_dir(self.source_dir, self.active_dir)

        self.assertFalse(self.active_dir.exists())

    def test_activation_failure_rolls_back_to_previous_copy(self):
        launcher.mirror_app_dir(self.source_dir, self.active_dir)
        original_app_py = (self.active_dir / "app.py").read_text()
        _write_bundle(self.source_dir, "v2")

        def _boom(_src, _dst):
            raise OSError("simulated rename failure")

        # Fail only the second rename (staging -> active), after the first
        # rename (active -> backup) has already happened.
        real_rename = launcher.os.rename
        calls = []

        def _flaky_rename(src, dst):
            calls.append((src, dst))
            if len(calls) == 2:
                raise OSError("simulated rename failure")
            return real_rename(src, dst)

        with unittest.mock.patch.object(launcher.os, "rename", side_effect=_flaky_rename):
            result = launcher.mirror_app_dir(self.source_dir, self.active_dir)

        self.assertEqual(result, self.active_dir)
        self.assertEqual((self.active_dir / "app.py").read_text(), original_app_py)
        self.assertFalse((self.active_dir.parent / "app.previous").exists())

    # ---- crash recovery --------------------------------------------------

    def test_recovers_from_backup_when_active_dir_missing(self):
        launcher.mirror_app_dir(self.source_dir, self.active_dir)
        original_app_py = (self.active_dir / "app.py").read_text()

        # Simulate a crash between the two renames of a previous swap:
        # active_dir is gone, backup still holds the last known-good copy.
        backup_dir = launcher._backup_dir_for(self.active_dir)
        shutil.move(str(self.active_dir), str(backup_dir))
        self.assertFalse(self.active_dir.exists())

        launcher._recover_interrupted_swap(self.active_dir)

        self.assertTrue(self.active_dir.exists())
        self.assertEqual((self.active_dir / "app.py").read_text(), original_app_py)

    # ---- validation helper -------------------------------------------------

    def test_validate_runtime_dir_flags_missing_and_empty_files(self):
        _write_bundle(self.source_dir, "v1")
        (self.source_dir / "charts.py").write_text("", encoding="utf-8")   # empty
        (self.source_dir / "fcst_loader.py").unlink()                      # missing

        problems = launcher._validate_runtime_dir(self.source_dir)

        self.assertIn("charts.py", problems)
        self.assertIn("fcst_loader.py", problems)
        self.assertNotIn("app.py", problems)


class DevSourceRunTestCase(unittest.TestCase):
    """When source_dir resolves to the same path as active_dir (dev/source
    run, no bundling involved), mirror_app_dir must be a safe no-op that
    still guarantees the user-state files exist."""

    def test_same_source_and_active_dir_is_noop_with_defaults(self):
        tmp = Path(self.enterContext(_temp_dir()))
        app_dir = tmp / "app"
        _write_bundle(app_dir, "v1")
        (app_dir / "settings.json").unlink()

        result = launcher.mirror_app_dir(app_dir, app_dir)

        self.assertEqual(result, app_dir)
        self.assertTrue((app_dir / "settings.json").exists())
        self.assertEqual((app_dir / "settings.json").read_text(), "{}")

    def test_no_source_dir_is_noop_with_defaults(self):
        tmp = Path(self.enterContext(_temp_dir()))
        app_dir = tmp / "app"
        app_dir.mkdir()

        result = launcher.mirror_app_dir(None, app_dir)

        self.assertEqual(result, app_dir)
        self.assertTrue((app_dir / "overrides.json").exists())
        self.assertTrue((app_dir / "settings.json").exists())


@contextlib.contextmanager
def _temp_dir():
    d = tempfile.mkdtemp(prefix="salesreport-launcher-test-")
    try:
        yield d
    finally:
        shutil.rmtree(d, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
