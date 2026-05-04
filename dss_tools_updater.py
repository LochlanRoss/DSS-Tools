"""DSS Tools update mini-app (Windows).

After the main app exits, this helper (frozen as ``DSSToolsUpdater.exe``) can:

1. Show a small status window (default).
2. Silently uninstall the previous Inno-installed build (same ``AppId``).
3. Remove transient data under ``%LOCALAPPDATA%\\DSSTools`` (cache, downloads, logs)
   while keeping ``dss_hours_tracker_config.json``.
4. Run the new ``DSSToolsSetup.exe`` with Inno silent flags so the previous install
   directory and Start-menu layout are reused where possible.

CLI (frozen)::

    DSSToolsUpdater.exe <installer_path> <parent_pid>

Development::

    python dss_tools_updater.py <installer_path> <parent_pid>

``--headless`` skips the window and only waits for the parent process then opens the
installer (legacy behaviour for tests).
"""

from __future__ import annotations

import argparse
import math
import os
import shlex
import shutil
import subprocess
import sys
import threading
import time
import tkinter as tk
from collections.abc import Callable
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk

# Must match installer/DSSTools.iss [Setup] AppId (single braces in registry).
INNO_UNINSTALL_SUBKEY = "{E7B8F9A0-1D2C-4E5F-8A9B-0C1D2E3F4A5B}_is1"
APP_DIRNAME = "DSSTools"
CONFIG_FILENAME = "dss_hours_tracker_config.json"
DISPLAY_NAME = "DSS Tools"
INNO_UNINSTALL_SILENT_FLAGS = ("/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART")
INNO_INSTALL_SILENT_FLAGS = ("/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/CLOSEAPPLICATIONS")


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
    if err == 5:
        return True
    return False


def localappdata_dss_tools_root() -> Path | None:
    base = os.environ.get("LOCALAPPDATA", "").strip()
    if not base:
        return None
    return Path(base) / APP_DIRNAME


def _read_inno_uninstall_string() -> str | None:
    if sys.platform != "win32":
        return None
    import winreg

    uninstall_root = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for sam in (
            winreg.KEY_READ | getattr(winreg, "KEY_WOW64_64KEY", 0),
            winreg.KEY_READ | getattr(winreg, "KEY_WOW64_32KEY", 0),
        ):
            try:
                with winreg.OpenKey(hive, uninstall_root, 0, sam) as root:
                    with winreg.OpenKey(root, INNO_UNINSTALL_SUBKEY) as sub:
                        value, _ = winreg.QueryValueEx(sub, "UninstallString")
                        if isinstance(value, str) and value.strip():
                            return value.strip()
            except OSError:
                continue
    return None


def parse_uninstall_executable(uninstall_string: str) -> Path | None:
    try:
        parts = shlex.split(uninstall_string, posix=False)
    except ValueError:
        return None
    if not parts:
        return None
    candidate = Path(parts[0]).expanduser()
    return candidate if candidate.is_file() else None


def read_inno_install_location() -> Path | None:
    if sys.platform != "win32":
        return None
    import winreg

    uninstall_root = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for sam in (
            winreg.KEY_READ | getattr(winreg, "KEY_WOW64_64KEY", 0),
            winreg.KEY_READ | getattr(winreg, "KEY_WOW64_32KEY", 0),
        ):
            try:
                with winreg.OpenKey(hive, uninstall_root, 0, sam) as root:
                    with winreg.OpenKey(root, INNO_UNINSTALL_SUBKEY) as sub:
                        value, _ = winreg.QueryValueEx(sub, "InstallLocation")
                        if isinstance(value, str) and value.strip():
                            loc = Path(value.strip())
                            if loc.is_dir():
                                return loc
            except (OSError, FileNotFoundError):
                continue
    return None


class CallableLog:
    def __init__(self, sink: object | None = None) -> None:
        self._sink = sink

    def __call__(self, line: str) -> None:
        if self._sink is not None:
            self._sink(line)


def _run_command_with_smoothed_progress(
    cmd: list[str],
    timeout: float,
    pct_lo: float,
    pct_hi: float,
    on_progress: Callable[[float], None] | None,
) -> int:
    """Run ``cmd`` while reporting smoothed progress from ``pct_lo`` toward ``pct_hi`` (inclusive)."""
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    proc = subprocess.Popen(cmd, creationflags=flags)
    t0 = time.monotonic()
    span = max(pct_hi - pct_lo, 1e-6)
    # Ease toward just below ``pct_hi`` while the process runs; snap to ``pct_hi`` at exit.
    cap = pct_hi - 0.25
    scale = 28.0 + span * 0.35
    deadline = t0 + timeout
    if on_progress is not None:
        on_progress(pct_lo)
    while proc.poll() is None:
        now = time.monotonic()
        if now > deadline:
            proc.kill()
            try:
                proc.wait(timeout=15)
            except (OSError, subprocess.TimeoutExpired):
                pass
            raise subprocess.TimeoutExpired(cmd, timeout)
        elapsed = now - t0
        frac = 1.0 - math.exp(-elapsed / scale)
        cur = pct_lo + (cap - pct_lo) * min(1.0, frac)
        if on_progress is not None:
            on_progress(min(cur, cap))
        time.sleep(0.1)
    if on_progress is not None:
        on_progress(pct_hi)
    return int(proc.returncode or 0)


def clean_transient_app_data(app_root: Path, log: CallableLog) -> None:
    """Delete cache, updates, logs, diagnostics; keep ``dss_hours_tracker_config.json``."""
    this_exe = Path(sys.executable).resolve()
    if not app_root.is_dir():
        return
    preserve = {CONFIG_FILENAME.casefold()}
    for child in sorted(app_root.iterdir(), key=lambda p: p.name.casefold()):
        name_cf = child.name.casefold()
        if name_cf in preserve:
            continue
        if child.resolve() == this_exe:
            log(f"Skip deleting running updater: {child}")
            continue
        try:
            if child.is_dir():
                if name_cf in ("cache", "updates"):
                    shutil.rmtree(child, ignore_errors=False)
                    log(f"Removed directory: {child.name}\\")
                continue
            if child.suffix.lower() == ".log":
                child.unlink(missing_ok=True)
                log(f"Removed file: {child.name}")
            elif child.name.startswith("diagnostic_snapshot_") and child.suffix.lower() == ".json":
                child.unlink(missing_ok=True)
                log(f"Removed file: {child.name}")
        except OSError as exc:
            log(f"Could not remove {child}: {exc}")


def run_uninstall_silent(log: CallableLog, on_progress: Callable[[float], None] | None = None) -> int:
    raw = _read_inno_uninstall_string()
    if raw is None:
        log("No existing Inno uninstall entry (fresh install or different product id).")
        if on_progress is not None:
            on_progress(50.0)
        return 0
    uninst = parse_uninstall_executable(raw)
    if uninst is None:
        log(f"Could not parse uninstall executable from: {raw!r}")
        return 2
    cmd = [str(uninst), *INNO_UNINSTALL_SILENT_FLAGS]
    log(f"Running uninstall: {uninst.name} {' '.join(INNO_UNINSTALL_SILENT_FLAGS)}")
    try:
        rc = _run_command_with_smoothed_progress(cmd, 600.0, 0.0, 50.0, on_progress)
    except subprocess.TimeoutExpired:
        log("Uninstall timed out.")
        return 1
    log(f"Uninstall finished (exit {rc}).")
    if rc != 0:
        log("Uninstall reported a non-zero exit code; continuing with install anyway.")
    return 0


def run_installer_silent(installer: Path, log: CallableLog, on_progress: Callable[[float], None] | None = None) -> int:
    cmd = [str(installer), *INNO_INSTALL_SILENT_FLAGS]
    log(f"Running installer: {installer.name} {' '.join(INNO_INSTALL_SILENT_FLAGS)}")
    try:
        rc = _run_command_with_smoothed_progress(cmd, 900.0, 50.0, 100.0, on_progress)
    except subprocess.TimeoutExpired:
        log("Installer timed out.")
        return 1
    if rc != 0:
        log(f"Silent installer exit code {rc}.")
        return rc
    log("Installer completed successfully.")
    return 0


def launch_installed_app(log: CallableLog) -> None:
    install_dir = read_inno_install_location()
    if install_dir is None:
        install_dir = Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / APP_DIRNAME
    exe = install_dir / "DSSTools.exe"
    if not exe.is_file():
        log(f"Installed app not found at {exe}")
        return
    log(f"Starting {exe}")
    try:
        os.startfile(str(exe))  # noqa: S606
    except OSError as exc:
        log(f"Could not start {exe}: {exc}")


def run_headless_legacy(installer: Path, parent_pid: int) -> int:
    """Wait for parent then open the installer with shell (no silent chain)."""
    if sys.platform != "win32":
        return 1
    inst = installer.expanduser().resolve()
    if not inst.is_file():
        return 1
    deadline = time.monotonic() + 900.0
    while time.monotonic() < deadline:
        if not parent_process_exists_windows(parent_pid):
            break
        time.sleep(0.35)
    else:
        return 1
    try:
        os.startfile(str(inst))  # noqa: S606
    except OSError:
        return 1
    return 0


class UpdateMiniApp(tk.Tk):
    def __init__(self, installer: Path, parent_pid: int) -> None:
        super().__init__()
        self._installer = installer.resolve()
        self._parent_pid = parent_pid
        self._failed = False
        self.title(f"{DISPLAY_NAME} Update")
        self.geometry("400x228")
        self.minsize(360, 200)
        self._apply_icon()
        self._pct_var = tk.DoubleVar(value=0.0)
        self._pct_text = tk.StringVar(value="0%")
        frm = ttk.Frame(self, padding=(10, 8))
        frm.pack(fill="both", expand=True)
        self._status = tk.StringVar(value="Waiting for the main application to close…")
        ttk.Label(frm, textvariable=self._status, wraplength=360, justify="left", font=("Segoe UI", 9)).pack(anchor="w")
        row = ttk.Frame(frm)
        row.pack(fill="x", pady=(8, 2))
        self._progress = ttk.Progressbar(
            row,
            mode="determinate",
            maximum=100,
            variable=self._pct_var,
            length=300,
        )
        self._progress.pack(side="left", fill="x", expand=True)
        ttk.Label(row, textvariable=self._pct_text, width=5, font=("Segoe UI", 9)).pack(side="right", padx=(8, 0))
        self._log_widget = scrolledtext.ScrolledText(
            frm, height=5, wrap="word", state="disabled", font=("Consolas", 8)
        )
        self._log_widget.pack(fill="both", expand=True, pady=(6, 0))
        ttk.Button(frm, text="Close", command=self.destroy).pack(anchor="e", pady=(6, 0))
        self._set_progress_pct(0.0)
        self.after(200, self._wait_parent_then_start)

    def _set_progress_pct(self, value: float) -> None:
        clamped = max(0.0, min(100.0, float(value)))
        self._pct_var.set(clamped)
        self._pct_text.set(f"{round(clamped)}%")

    def _apply_icon(self) -> None:
        for cand in (
            Path(sys.executable).resolve().parent / "dss_tools.ico",
            Path(sys.executable).resolve().parent / "DSSTools.exe",
        ):
            if cand.is_file() and cand.suffix.lower() == ".ico":
                try:
                    self.iconbitmap(default=str(cand))
                    return
                except tk.TclError:
                    continue

    def _log(self, line: str) -> None:
        self._log_widget.configure(state="normal")
        self._log_widget.insert(tk.END, line + "\n")
        self._log_widget.see(tk.END)
        self._log_widget.configure(state="disabled")

    def _wait_parent_then_start(self) -> None:
        if parent_process_exists_windows(self._parent_pid):
            self.after(250, self._wait_parent_then_start)
            return
        self._status.set("Main application closed. Preparing upgrade…")
        self._set_progress_pct(0.0)
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self) -> None:
        log = CallableLog(lambda m: self.after(0, lambda msg=m: self._log(msg)))

        def report_progress(p: float) -> None:
            self.after(0, lambda v=p: self._set_progress_pct(v))

        try:
            self.after(0, lambda: self._status.set("Uninstalling previous version (0–50%)…"))
            code = run_uninstall_silent(log, report_progress)
            if code == 2:
                self.after(0, lambda: self._finish_install_error("Could not parse the uninstall program from the registry.", 2))
                return
            app_root = localappdata_dss_tools_root()
            if app_root is not None:
                self.after(0, lambda: self._status.set("Cleaning temporary files…"))
                log(f"Cleaning transient data under {app_root}")
                clean_transient_app_data(app_root, log)
            self.after(0, lambda: self._status.set("Installing new version (50–100%)…"))
            code = run_installer_silent(self._installer, log, report_progress)
            if code != 0:
                self.after(
                    0,
                    lambda c=code: self._offer_interactive_install(
                        "Silent install did not complete successfully. Open the installer with the full wizard?",
                        c,
                    ),
                )
                return
            self.after(0, self._finish_ok)
        except Exception as exc:  # pragma: no cover - defensive UI path
            self.after(0, lambda: self._finish_install_error(str(exc), 1))

    def _finish_ok(self) -> None:
        self._set_progress_pct(100.0)
        self._status.set("Update finished.")
        self._log("Launching DSS Tools…")
        launch_installed_app(CallableLog(lambda m: self.after(0, lambda: self._log(m))))
        self.after(800, self.destroy)

    def _finish_install_error(self, message: str, code: int) -> None:
        self._failed = True
        self._status.set("Update failed.")
        self._log(f"{message} (code {code})")
        messagebox.showerror(DISPLAY_NAME + " Update", f"{message}\n\nExit code: {code}", parent=self)

    def _offer_interactive_install(self, message: str, code: int) -> None:
        self._status.set("Silent install failed.")
        self._log(message)
        if messagebox.askyesno(
            DISPLAY_NAME + " Update",
            message + "\n\nThis opens the normal installer window so you can finish manually.",
            parent=self,
        ):
            try:
                os.startfile(str(self._installer))  # noqa: S606
            except OSError as exc:
                messagebox.showerror(DISPLAY_NAME + " Update", str(exc), parent=self)
        else:
            self._failed = True
        self.destroy()


def run_gui(installer: Path, parent_pid: int) -> int:
    if sys.platform != "win32":
        return 1
    inst = installer.expanduser().resolve()
    if not inst.is_file():
        messagebox.showerror(DISPLAY_NAME + " Update", f"Installer not found:\n{inst}")
        return 1
    app = UpdateMiniApp(inst, parent_pid)
    app.mainloop()
    return 1 if getattr(app, "_failed", False) else 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="DSS Tools update helper.")
    parser.add_argument("installer", type=Path, help="Path to DSSToolsSetup.exe")
    parser.add_argument("parent_pid", type=int, help="PID of the running DSS Tools instance")
    parser.add_argument(
        "--headless",
        action="store_true",
        help="No window: wait for parent then launch installer only (legacy / tests).",
    )
    args = parser.parse_args(argv)
    if args.headless:
        return run_headless_legacy(Path(args.installer), args.parent_pid)
    if sys.platform != "win32":
        print("The DSS Tools updater only runs on Windows.", file=sys.stderr)
        return 1
    return run_gui(Path(args.installer), args.parent_pid)


if __name__ == "__main__":
    raise SystemExit(main())
