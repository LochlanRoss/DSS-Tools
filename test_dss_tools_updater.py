"""Tests for ``dss_tools_updater`` (Windows process wait helper)."""

from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path
from unittest.mock import Mock, patch


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

    def test_parse_uninstall_executable_expandvars_and_quotes(self) -> None:
        from dss_tools_updater import parse_uninstall_executable

        fake_unins = Path(r"C:\Program Files\DSS Tools\unins000.exe")
        line = f'"{fake_unins}" /VERYSILENT /SUPPRESSMSGBOXES'
        with patch("pathlib.Path.is_file", return_value=True):
            got = parse_uninstall_executable(line)
        self.assertIsNotNone(got)
        self.assertEqual(got, fake_unins)

    def test_row_matches_dss_tools_fuzzy_inno(self) -> None:
        from dss_tools_updater import _row_matches_dss_tools_fuzzy_inno

        self.assertTrue(
            _row_matches_dss_tools_fuzzy_inno(
                "DSS Tools",
                "Other",
                "",
                r'"C:\Program Files\DSS Tools\unins000.exe" /SILENT',
            )
        )
        self.assertTrue(
            _row_matches_dss_tools_fuzzy_inno(
                "",
                "DSS Tools",
                "",
                r'"C:\Program Files\DSS Tools\unins000.exe"',
            )
        )
        self.assertTrue(
            _row_matches_dss_tools_fuzzy_inno(
                "",
                "",
                r"C:\Program Files\DSS Tools",
                r'"C:\Program Files\DSS Tools\unins000.exe"',
            )
        )
        self.assertFalse(
            _row_matches_dss_tools_fuzzy_inno(
                "Other App",
                "Other",
                r"C:\Other",
                r'"C:\Other\unins000.exe"',
            )
        )
        self.assertFalse(
            _row_matches_dss_tools_fuzzy_inno(
                "DSS Tools",
                "DSS Tools",
                "",
                r'"C:\Python313\python.exe" -m pip',
            )
        )

    def test_clean_transient_app_data_preserves_config(self) -> None:
        from dss_tools_updater import CallableLog, clean_transient_app_data

        lines: list[str] = []
        current_exe = Path(sys.executable).resolve()

        def fake_child(name: str, *, is_dir: bool, resolved: Path | None = None) -> Mock:
            child = Mock()
            child.name = name
            child.is_dir.return_value = is_dir
            child.resolve.return_value = resolved or Path(rf"C:\fake\{name}")
            child.unlink = Mock()
            return child

        config = fake_child("dss_hours_tracker_config.json", is_dir=False)
        cache = fake_child("cache", is_dir=True)
        updates = fake_child("updates", is_dir=True)
        handoff = fake_child("update_handoff.log", is_dir=False)
        updater = fake_child("DSSToolsUpdater.exe", is_dir=False, resolved=current_exe)
        stray = fake_child("odd_folder", is_dir=True)
        root = Mock()
        root.is_dir.return_value = True
        root.iterdir.return_value = [config, cache, updates, handoff, updater, stray]

        with patch("dss_tools_updater.shutil.rmtree") as rmtree:
            clean_transient_app_data(root, CallableLog(lines.append))

        handoff.unlink.assert_called_once_with(missing_ok=True)
        updater.unlink.assert_not_called()
        rmtree.assert_any_call(cache, ignore_errors=False)
        rmtree.assert_any_call(updates, ignore_errors=False)
        rmtree.assert_any_call(stray, ignore_errors=False)
        self.assertTrue(any("Skip deleting running updater" in line for line in lines))

    def test_terminate_process_tree_with_retries_waits_before_final_attempt(self) -> None:
        from dss_tools_updater import terminate_process_tree_with_retries_windows

        sleep_calls: list[float] = []
        log_lines: list[str] = []
        terminate_results = [(1, "busy")] * 4 + [(0, "done")]
        still_running = [(True, "")] * 4 + [(False, "Skipped: no process with this PID (main app already exited); taskkill not used.")]

        with (
            patch("dss_tools_updater.terminate_process_tree_windows", side_effect=terminate_results) as terminate_mock,
            patch("dss_tools_updater.should_taskkill_parent_process", side_effect=still_running),
            patch("dss_tools_updater.time.sleep", side_effect=lambda seconds: sleep_calls.append(seconds)),
        ):
            code, detail = terminate_process_tree_with_retries_windows(
                1234,
                r"C:\Program Files\DSS Tools\DSSTools.exe",
                log_lines.append,
            )

        self.assertEqual(code, 0)
        self.assertEqual(terminate_mock.call_count, 5)
        self.assertEqual(sleep_calls, [1.0, 1.0, 1.0, 30.0])
        self.assertIn("Attempt 5/5: done", detail)
        self.assertTrue(any("Waiting 30s before the final retry." in line for line in log_lines))


if __name__ == "__main__":
    unittest.main()
