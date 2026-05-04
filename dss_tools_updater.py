"""Small sidecar: wait for DSS Tools to exit, then start the downloaded installer.

Frozen installs run this as ``DSSToolsUpdater.exe`` next to ``DSSTools.exe``::

    DSSToolsUpdater.exe <installer_path> <parent_pid>

During development the main app may invoke::

    python dss_tools_updater.py <installer_path> <parent_pid>
"""

from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path


def parent_process_exists_windows(pid: int) -> bool:
    """Return True if a Windows process with ``pid`` is still running (best-effort)."""
    if sys.platform != "win32" or pid <= 0:
        return False
    import ctypes

    k32 = ctypes.windll.kernel32
    SYNCHRONIZE = 0x00100000
    h = k32.OpenProcess(SYNCHRONIZE, False, ctypes.c_uint32(pid))
    if h:
        k32.CloseHandle(h)
        return True
    err = int(k32.GetLastError())
    # Access denied often means the process exists but ACLs block SYNCHRONIZE.
    if err == 5:
        return True
    return False


def _notify_error(message: str) -> None:
    if sys.platform == "win32":
        try:
            import ctypes

            ctypes.windll.user32.MessageBoxW(0, message, "DSS Tools Updater", 0x10)
        except Exception:
            print(message, file=sys.stderr)
    else:
        print(message, file=sys.stderr)


def main() -> int:
    parser = argparse.ArgumentParser(description="Install DSS Tools after the main app exits.")
    parser.add_argument("installer", type=Path, help="Path to the installer (.exe/.msi)")
    parser.add_argument("parent_pid", type=int, help="PID of the running DSS Tools instance")
    args = parser.parse_args()

    if sys.platform != "win32":
        _notify_error("The DSS Tools updater only runs on Windows.")
        return 1

    inst = args.installer.expanduser().resolve()
    if not inst.is_file():
        _notify_error(f"Installer not found:\n{inst}")
        return 1

    pid = args.parent_pid
    deadline = time.monotonic() + 900.0
    while time.monotonic() < deadline:
        if not parent_process_exists_windows(pid):
            break
        time.sleep(0.35)
    else:
        _notify_error(
            "Timed out waiting for DSS Tools to close.\n\n"
            "Close the application and run the installer manually from your updates folder "
            f"({inst.parent})."
        )
        return 1

    # ``os.startfile`` uses shell association / UAC flow like Explorer double-click.
    try:
        os.startfile(str(inst))  # noqa: S606
    except OSError as exc:
        _notify_error(f"Could not start the installer:\n{exc}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
