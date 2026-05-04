"""Tests for ``dss_tools_updater`` (Windows process wait helper)."""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path


class TestDssToolsUpdater(unittest.TestCase):
    def test_parent_pid_invalid(self) -> None:
        from dss_tools_updater import parent_process_exists_windows

        self.assertFalse(parent_process_exists_windows(0))
        self.assertFalse(parent_process_exists_windows(-1))

    @unittest.skipUnless(sys.platform == "win32", "Windows-only")
    def test_parent_exists_self(self) -> None:
        from dss_tools_updater import parent_process_exists_windows

        self.assertTrue(parent_process_exists_windows(os.getpid()))

    @unittest.skipUnless(sys.platform == "win32", "Windows-only")
    def test_parent_absent_unlikely_pid(self) -> None:
        from dss_tools_updater import parent_process_exists_windows

        # Avoid huge integers: ``c_uint32`` truncates and could match a live PID.
        self.assertFalse(parent_process_exists_windows(999_999_999))

    @unittest.skipUnless(sys.platform == "win32", "Windows-only")
    def test_terminate_process_tree_nonexistent_pid_skips_taskkill(self) -> None:
        from dss_tools_updater import terminate_process_tree_windows

        code, detail = terminate_process_tree_windows(999_999_999, str(Path(sys.executable).resolve()))
        self.assertEqual(code, 0)
        self.assertIn("Skipped", detail)
        self.assertIn("already exited", detail)

    @unittest.skipUnless(sys.platform == "win32", "Windows-only")
    def test_should_taskkill_parent_matches_own_image_path(self) -> None:
        from dss_tools_updater import get_process_image_path_windows, should_taskkill_parent_process

        img = get_process_image_path_windows(os.getpid())
        self.assertIsNotNone(img)
        ok, msg = should_taskkill_parent_process(os.getpid(), img)
        self.assertTrue(ok)
        self.assertEqual(msg, "")

    @unittest.skipUnless(sys.platform == "win32", "Windows-only")
    def test_should_taskkill_rejects_wrong_parent_path(self) -> None:
        from dss_tools_updater import should_taskkill_parent_process

        ok, msg = should_taskkill_parent_process(os.getpid(), r"C:\Windows\System32\notepad.exe")
        self.assertFalse(ok)
        self.assertIn("not the updater parent", msg)

    @unittest.skipUnless(sys.platform == "win32", "Windows-only")
    def test_get_process_image_path_self(self) -> None:
        from dss_tools_updater import get_process_image_path_windows

        img = get_process_image_path_windows(os.getpid())
        self.assertIsNotNone(img)
        self.assertTrue(len(img) > 3)

    def test_clean_transient_app_data_preserves_config(self) -> None:
        from dss_tools_updater import CallableLog, clean_transient_app_data

        lines: list[str] = []
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "dss_hours_tracker_config.json").write_text("{}", encoding="utf-8")
            cache = root / "cache"
            cache.mkdir()
            (cache / "x.json").write_text("[]", encoding="utf-8")
            (root / "updates").mkdir()
            (root / "updates" / "old.exe").write_bytes(b"")
            (root / "update_handoff.log").write_text("x", encoding="utf-8")
            clean_transient_app_data(root, CallableLog(lines.append))
            self.assertTrue((root / "dss_hours_tracker_config.json").is_file())
            self.assertFalse(cache.exists())
            self.assertFalse((root / "updates").exists())
            self.assertFalse((root / "update_handoff.log").exists())


if __name__ == "__main__":
    unittest.main()
