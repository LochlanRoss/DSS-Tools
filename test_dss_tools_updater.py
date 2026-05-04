"""Tests for ``dss_tools_updater`` (Windows process wait helper)."""

from __future__ import annotations

import os
import sys
import unittest


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


if __name__ == "__main__":
    unittest.main()
