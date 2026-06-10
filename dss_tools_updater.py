"""DSS Tools update mini-app (Windows).

After the main app exits, this helper (frozen as ``DSSToolsUpdater.exe``) can:

1. Show a small status window (default).
2. Silently uninstall the previous Inno-installed build (registry ``{AppId}_is1`` when present,
   otherwise enumerate ``Uninstall`` for a matching DSS Tools Inno entry).
3. Remove transient data under ``%LOCALAPPDATA%\\DSSTools`` (everything except
   ``dss_hours_tracker_config.json``, including cache, updates, copied updater exe, logs).
4. Run the new ``DSSToolsSetup.exe`` with Inno silent flags so the previous install
   directory and Start-menu layout are reused where possible.

CLI (frozen)::

    DSSToolsUpdater.exe <installer_path> <parent_pid> [<parent_executable>]

The optional *parent_executable* is the resolved ``sys.executable`` of the main app when the
handoff was started. ``taskkill`` runs only if that PID still maps to the same image, so a
recycled PID pointing at another program is never killed.

Development::

    python dss_tools_updater.py <installer_path> <parent_pid> [<parent_executable>]

``--headless`` skips the window, ends the parent process tree with ``taskkill``, then opens the
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
import tempfile
import threading
import time
import tkinter as tk
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk

# Must match installer/DSSTools.iss [Setup] AppId (single braces in registry).
INNO_UNINSTALL_SUBKEY = "{E7B8F9A0-1D2C-4E5F-8A9B-0C1D2E3F4A5B}_is1"
APP_DIRNAME = "DSSTools"
CONFIG_FILENAME = "dss_hours_tracker_config.json"
DISPLAY_NAME = "DSS Tools"
INNO_UNINSTALL_SILENT_FLAGS = ("/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART")
INNO_INSTALL_SILENT_FLAGS = ("/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/CLOSEAPPLICATIONS")
PARENT_TERMINATION_MAX_ATTEMPTS = 5
PARENT_TERMINATION_RETRY_DELAY_SECONDS = 1.0
PARENT_TERMINATION_FINAL_WAIT_SECONDS = 30.0


def _windows_hidden_subprocess_kwargs(extra_creationflags: int = 0) -> dict:
    """Avoid a flashing console when spawning ``taskkill``, ``cmd``, or Inno from a GUI process."""
    if sys.platform != "win32":
        return {}
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) | int(extra_creationflags)
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = getattr(subprocess, "SW_HIDE", 0)
    return {"creationflags": flags, "startupinfo": startupinfo}


def parent_process_exists_windows(pid: int) -> bool:
    """Return True if a Windows process with ``pid`` is still running (best-effort).

    Note: ``OpenProcess`` can return ``ERROR_ACCESS_DENIED`` (5) after the process has exited
    and the PID was reused, or across integrity levels. Treating 5 as \"still running\" caused
    endless waits in the update helper; upgrade flow uses :func:`terminate_process_tree_windows`
    instead of polling this function.
    """
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


def get_process_image_path_windows(pid: int) -> str | None:
    """Return the full Win32 path of the executable for *pid*, or None if unavailable."""
    if sys.platform != "win32" or pid <= 0:
        return None
    import ctypes
    from ctypes import wintypes

    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    k32 = ctypes.windll.kernel32
    k32.QueryFullProcessImageNameW.argtypes = [
        wintypes.HANDLE,
        wintypes.DWORD,
        ctypes.POINTER(ctypes.c_wchar),
        ctypes.POINTER(wintypes.DWORD),
    ]
    k32.QueryFullProcessImageNameW.restype = wintypes.BOOL

    h = k32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, ctypes.c_uint32(pid))
    if not h:
        return None
    try:
        buf = ctypes.create_unicode_buffer(4096)
        size = wintypes.DWORD(len(buf))
        if not k32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
            return None
        path = buf.value.strip()
        return path or None
    finally:
        k32.CloseHandle(h)


def _resolved_path_casefold(path_str: str) -> str | None:
    try:
        return str(Path(path_str).expanduser().resolve()).casefold()
    except OSError:
        return None


def should_taskkill_parent_process(parent_pid: int, parent_executable: str) -> tuple[bool, str]:
    """Decide whether ``taskkill`` is safe for *parent_pid*.

    If the process is already gone, or the PID now belongs to a different image (PID reuse),
    return (False, reason) and **do not** run ``taskkill``.
    """
    if sys.platform != "win32" or parent_pid <= 0:
        return False, "Skipped: not Windows or invalid PID."
    image = get_process_image_path_windows(parent_pid)
    if not image:
        return False, "Skipped: no process with this PID (main app already exited); taskkill not used."

    parent_exe = (parent_executable or "").strip()
    if parent_exe:
        want = _resolved_path_casefold(parent_exe)
        got = _resolved_path_casefold(image)
        if want and got and want == got:
            return True, ""
        return (
            False,
            f"Skipped: PID {parent_pid} is {image!r}, not the updater parent {parent_executable!r} "
            f"(PID reuse or wrong process); taskkill not used.",
        )

    if Path(image).name.casefold() == "dsstools.exe":
        return True, ""
    return (
        False,
        f"Skipped: PID {parent_pid} is {image!r} (not DSSTools.exe) and no parent executable was passed; taskkill not used.",
    )


def terminate_process_tree_windows(parent_pid: int, parent_executable: str = "") -> tuple[int, str]:
    """Force-end the process tree rooted at *parent_pid* using ``taskkill`` (Windows), if identity checks pass.

    *parent_executable* must be the resolved ``sys.executable`` passed from the main app at
    handoff. ``taskkill`` is skipped when the PID is free (clean exit) or maps to another image.
    """
    if sys.platform != "win32" or parent_pid <= 0:
        return 0, ""
    ok, skip_reason = should_taskkill_parent_process(parent_pid, parent_executable)
    if not ok:
        return 0, skip_reason

    try:
        proc = subprocess.run(
            ["taskkill", "/PID", str(parent_pid), "/T", "/F"],
            capture_output=True,
            timeout=60,
            text=True,
            check=False,
            **_windows_hidden_subprocess_kwargs(),
        )
    except subprocess.TimeoutExpired:
        return 1, "taskkill timed out"
    parts = [proc.stdout or "", proc.stderr or ""]
    detail = "\n".join(p.strip() for p in parts if p and p.strip())
    return int(proc.returncode or 0), detail[:800]


def terminate_process_tree_with_retries_windows(
    parent_pid: int,
    parent_executable: str = "",
    log: Callable[[str], None] | None = None,
    *,
    max_attempts: int = PARENT_TERMINATION_MAX_ATTEMPTS,
    retry_delay_seconds: float = PARENT_TERMINATION_RETRY_DELAY_SECONDS,
    final_wait_seconds: float = PARENT_TERMINATION_FINAL_WAIT_SECONDS,
) -> tuple[int, str]:
    """Retry ending the parent process tree before the upgrade proceeds.

    Defaults to 5 total attempts, with a 30 second wait before the final attempt.
    """
    attempts = max(1, int(max_attempts or 1))
    final_code = 0
    details: list[str] = []

    for attempt in range(1, attempts + 1):
        code, detail = terminate_process_tree_windows(parent_pid, parent_executable)
        final_code = code

        if detail:
            details.append(f"Attempt {attempt}/{attempts}: {detail}")
        elif code not in (0, 128):
            details.append(f"Attempt {attempt}/{attempts}: taskkill exit {code}")

        if log:
            if detail:
                if detail.startswith("Skipped:"):
                    log(detail)
                else:
                    log(f"Stop DSS attempt {attempt}/{attempts}: {detail}")
            else:
                log(f"Stop DSS attempt {attempt}/{attempts}: taskkill exit {code}.")

        ok, skip_reason = should_taskkill_parent_process(parent_pid, parent_executable)
        if not ok:
            if skip_reason and (not details or details[-1] != skip_reason):
                details.append(skip_reason)
            return final_code, "\n".join(details).strip()

        if attempt >= attempts:
            break

        wait_seconds = final_wait_seconds if attempt == attempts - 1 else retry_delay_seconds
        if wait_seconds > 0:
            if log:
                if attempt == attempts - 1:
                    log(
                        f"DSS Tools is still running after {attempt} close attempts. "
                        f"Waiting {int(wait_seconds)}s before the final retry."
                    )
                else:
                    log(f"DSS Tools is still running; retrying close in {wait_seconds:.0f}s.")
            time.sleep(wait_seconds)

    return final_code, "\n".join(details).strip()


def localappdata_dss_tools_root() -> Path | None:
    base = os.environ.get("LOCALAPPDATA", "").strip()
    if not base:
        return None
    return Path(base) / APP_DIRNAME


@dataclass(frozen=True)
class InnoDssToolsUninstallInfo:
    uninstall_string: str
    install_location: Path | None
    registry_subkey_name: str


def _row_matches_dss_tools_fuzzy_inno(
    display_name: str, publisher: str, install_location: str, uninstall_string: str
) -> bool:
    """True when registry values identify an Inno uninstaller for DSS Tools (non-AppId lookup)."""
    us = uninstall_string.casefold()
    if "unins" not in us or ".exe" not in us:
        return False
    dn = display_name.strip().casefold()
    pub = publisher.strip().casefold()
    if dn == "dss tools" or dn.startswith("dss tools ") or pub == "dss tools":
        return True
    il = install_location.strip()
    if il:
        try:
            if Path(il.rstrip("\\/")).name.casefold() == "dss tools":
                return True
        except (OSError, ValueError):
            pass
    try:
        parts = shlex.split(uninstall_string, posix=False)
    except ValueError:
        return False
    if not parts:
        return False
    try:
        p = Path(parts[0]).expanduser()
        if len(p.parts) >= 2 and p.parts[-2].casefold() == "dss tools" and "unins" in p.name.casefold():
            return True
    except (OSError, ValueError):
        pass
    return False


def discover_inno_dss_tools_uninstall_info() -> InnoDssToolsUninstallInfo | None:
    """Resolve Inno ``UninstallString`` (and optional ``InstallLocation``) for DSS Tools."""
    if sys.platform != "win32":
        return None
    import winreg

    uninstall_root = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"

    def regstr(sub: object, name: str) -> str:
        try:
            value, _ = winreg.QueryValueEx(sub, name)
            if isinstance(value, str):
                return value
        except OSError:
            pass
        return ""

    def read_exact(root: object) -> InnoDssToolsUninstallInfo | None:
        try:
            with winreg.OpenKey(root, INNO_UNINSTALL_SUBKEY) as sub:
                chosen = ""
                for key in ("UninstallString", "QuietUninstallString"):
                    line = regstr(sub, key).strip()
                    if not line:
                        continue
                    exe = parse_uninstall_executable(line)
                    if exe is None or "unins" not in exe.name.casefold():
                        continue
                    chosen = line
                    break
                if not chosen:
                    return None
                il_raw = regstr(sub, "InstallLocation").strip()
                loc: Path | None = None
                if il_raw:
                    try:
                        cand = Path(il_raw)
                        if cand.is_dir():
                            loc = cand
                    except OSError:
                        pass
                return InnoDssToolsUninstallInfo(chosen, loc, INNO_UNINSTALL_SUBKEY)
        except OSError:
            return None

    def read_fuzzy(root: object, subkey_name: str) -> InnoDssToolsUninstallInfo | None:
        if subkey_name == INNO_UNINSTALL_SUBKEY:
            return None
        try:
            with winreg.OpenKey(root, subkey_name) as sub:
                dname = regstr(sub, "DisplayName")
                pub = regstr(sub, "Publisher")
                il_raw = regstr(sub, "InstallLocation")
                for key in ("UninstallString", "QuietUninstallString"):
                    line = regstr(sub, key).strip()
                    if not line:
                        continue
                    if not _row_matches_dss_tools_fuzzy_inno(dname, pub, il_raw, line):
                        continue
                    exe = parse_uninstall_executable(line)
                    if exe is None or "unins" not in exe.name.casefold():
                        continue
                    il_clean = il_raw.strip()
                    loc: Path | None = None
                    if il_clean:
                        try:
                            cand = Path(il_clean)
                            if cand.is_dir():
                                loc = cand
                        except OSError:
                            pass
                    return InnoDssToolsUninstallInfo(line, loc, subkey_name)
                return None
        except OSError:
            return None

    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for sam in (
            winreg.KEY_READ | getattr(winreg, "KEY_WOW64_64KEY", 0),
            winreg.KEY_READ | getattr(winreg, "KEY_WOW64_32KEY", 0),
        ):
            try:
                with winreg.OpenKey(hive, uninstall_root, 0, sam) as root:
                    hit = read_exact(root)
                    if hit is not None:
                        return hit
            except OSError:
                continue

    fuzzy_hits: list[tuple[int, InnoDssToolsUninstallInfo]] = []
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        hive_pri = 0 if hive == winreg.HKEY_LOCAL_MACHINE else 100
        for sam in (
            winreg.KEY_READ | getattr(winreg, "KEY_WOW64_64KEY", 0),
            winreg.KEY_READ | getattr(winreg, "KEY_WOW64_32KEY", 0),
        ):
            sam_pri = 0 if sam & getattr(winreg, "KEY_WOW64_64KEY", 0) else 1
            try:
                with winreg.OpenKey(hive, uninstall_root, 0, sam) as root:
                    idx = 0
                    while True:
                        try:
                            name = winreg.EnumKey(root, idx)
                        except OSError:
                            break
                        idx += 1
                        info = read_fuzzy(root, name)
                        if info is not None:
                            fuzzy_hits.append((hive_pri + sam_pri, info))
            except OSError:
                continue
    if not fuzzy_hits:
        return None
    fuzzy_hits.sort(key=lambda t: t[0])
    return fuzzy_hits[0][1]


def _read_inno_uninstall_string() -> str | None:
    info = discover_inno_dss_tools_uninstall_info()
    return info.uninstall_string if info else None


def parse_uninstall_executable(uninstall_string: str) -> Path | None:
    """Resolve the first ``*.exe`` path in *uninstall_string* that exists on disk (Inno / registry quirks)."""
    raw = os.path.expandvars((uninstall_string or "").strip())
    if not raw:
        return None

    def try_exe_token(tok: str) -> Path | None:
        t = tok.strip().strip('"').strip("'")
        if not t.lower().endswith(".exe"):
            return None
        try:
            p = Path(t).expanduser()
            if p.is_file():
                return p
            r = p.resolve(strict=False)
            if r.is_file():
                return r
        except OSError:
            return None
        return None

    candidates: list[str] = []
    try:
        candidates.extend(shlex.split(raw, posix=False))
    except ValueError:
        pass
    # Quoted path when shlex fails or splits oddly
    if raw.startswith('"'):
        end = raw.find('"', 1)
        if end > 1:
            candidates.append(raw[1:end])
    if not candidates and raw.split():
        candidates.append(raw.split()[0])
    seen: set[str] = set()
    for tok in candidates:
        k = tok.casefold()
        if k in seen:
            continue
        seen.add(k)
        hit = try_exe_token(tok)
        if hit is not None:
            return hit
    return None


def read_inno_install_location() -> Path | None:
    info = discover_inno_dss_tools_uninstall_info()
    if info is None:
        return None
    if info.install_location is not None and info.install_location.is_dir():
        return info.install_location
    exe = parse_uninstall_executable(info.uninstall_string)
    if exe is not None:
        parent = exe.parent
        if parent.is_dir():
            return parent
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
    proc = subprocess.Popen(
        cmd,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        **_windows_hidden_subprocess_kwargs(),
    )
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
    """Remove everything under *app_root* except ``dss_hours_tracker_config.json`` (user settings)."""
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
                shutil.rmtree(child, ignore_errors=False)
                log(f"Removed directory: {child.name}\\")
            else:
                child.unlink(missing_ok=True)
                log(f"Removed file: {child.name}")
        except OSError as exc:
            log(f"Could not remove {child}: {exc}")


def schedule_delete_path_windows(path: Path) -> None:
    """Best-effort delete after this process can exit (needed for staged *.exe under %%TEMP%%)."""
    if sys.platform != "win32":
        try:
            path.unlink(missing_ok=True)
        except OSError:
            pass
        return
    if not path.is_file():
        return
    ps = str(path)
    if '"' in ps or "\r" in ps or "\n" in ps:
        return
    subprocess.Popen(
        ["cmd", "/c", f'ping 127.0.0.1 -n 2 >nul & del /f /q "{ps}"'],
        close_fds=True,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        **_windows_hidden_subprocess_kwargs(getattr(subprocess, "DETACHED_PROCESS", 0)),
    )


def cleanup_staged_temp_artifacts(installer_path: Path) -> None:
    """Remove staged copies under %%TEMP%% (installer / updater) left from elevation handoff."""
    temp_root = Path(tempfile.gettempdir()).resolve()
    inst = installer_path.resolve()
    try:
        inst.relative_to(temp_root)
    except ValueError:
        pass
    else:
        if inst.is_file() and inst.name.lower().startswith("dss_tools_setup_"):
            schedule_delete_path_windows(inst)
    exe = Path(sys.argv[0]).resolve()
    try:
        exe.relative_to(temp_root)
    except ValueError:
        return
    if exe.is_file() and "dss_tools_updater_" in exe.name.lower():
        schedule_delete_path_windows(exe)


def run_uninstall_silent(log: CallableLog, on_progress: Callable[[float], None] | None = None) -> tuple[int, str]:
    """Return ``(exit_code, empty_string)`` or ``(2, user_message)`` when the uninstall path cannot be used."""
    info = discover_inno_dss_tools_uninstall_info()
    if info is None:
        log(
            "No Inno uninstall entry for DSS Tools under Uninstall (expected subkey "
            f"{INNO_UNINSTALL_SUBKEY!r} missing and no fuzzy match). Portable or non-Inno install?"
        )
        if on_progress is not None:
            on_progress(50.0)
        return 0, ""
    if info.registry_subkey_name != INNO_UNINSTALL_SUBKEY:
        log(
            f"Matched Inno uninstall via registry enumeration ({info.registry_subkey_name!r}); "
            f"expected AppId subkey {INNO_UNINSTALL_SUBKEY!r} was absent or incomplete."
        )
    raw = info.uninstall_string
    uninst = parse_uninstall_executable(raw)
    if uninst is None:
        log(f"Could not parse uninstall executable from registry line: {raw!r}")
        frag = raw if len(raw) <= 280 else raw[:280] + "…"
        hint = (
            "Windows listed an uninstall entry for DSS Tools, but the uninstall command is not a "
            "usable path to an Inno unins*.exe file (or the file is missing).\n\n"
            f"Registry fragment:\n{frag}"
        )
        return 2, hint
    cmd = [str(uninst), *INNO_UNINSTALL_SILENT_FLAGS]
    log(f"Running uninstall: {uninst.name} {' '.join(INNO_UNINSTALL_SILENT_FLAGS)}")
    try:
        rc = _run_command_with_smoothed_progress(cmd, 600.0, 0.0, 50.0, on_progress)
    except subprocess.TimeoutExpired:
        log("Uninstall timed out.")
        return 1, ""
    log(f"Uninstall finished (exit {rc}).")
    if rc != 0:
        log("Uninstall reported a non-zero exit code; continuing with install anyway.")
    return 0, ""


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


def run_headless_legacy(installer: Path, parent_pid: int, parent_executable: str = "") -> int:
    """End the parent process tree (if still the same executable) then open the installer with shell."""
    if sys.platform != "win32":
        return 1
    inst = installer.expanduser().resolve()
    if not inst.is_file():
        return 1
    terminate_process_tree_with_retries_windows(parent_pid, parent_executable)
    time.sleep(1.0)
    try:
        os.startfile(str(inst))  # noqa: S606
    except OSError:
        return 1
    return 0


class UpdateMiniApp(tk.Tk):
    def __init__(self, installer: Path, parent_pid: int, parent_executable: str = "") -> None:
        super().__init__()
        self._installer = installer.resolve()
        self._parent_pid = parent_pid
        self._parent_executable = (parent_executable or "").strip()
        self._failed = False
        self.title(f"{DISPLAY_NAME} Update")
        self.geometry("400x228")
        self.minsize(360, 200)
        self._apply_icon()
        self._pct_var = tk.DoubleVar(value=0.0)
        self._pct_text = tk.StringVar(value="0%")
        frm = ttk.Frame(self, padding=(10, 8))
        frm.pack(fill="both", expand=True)
        self._status = tk.StringVar(value="Stopping the previous application…")
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
        self.after(200, self._terminate_parent_then_start_upgrade)

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

    def _terminate_parent_then_start_upgrade(self) -> None:
        """End the main app process tree (``taskkill``), then run uninstall/install.

        Polling ``OpenProcess`` / ``parent_process_exists_windows`` was unreliable when the
        updater runs elevated (``runas``): access denied on PID reuse looked like \"still running\"
        and the UI never advanced.
        """
        self._status.set("Stopping the previous application…")

        def work() -> None:
            code, detail = terminate_process_tree_with_retries_windows(
                self._parent_pid,
                self._parent_executable,
                lambda msg: self.after(0, lambda line=msg: self._log(line)),
            )
            time.sleep(1.0)

            def on_done() -> None:
                if detail:
                    if detail.startswith("Skipped:"):
                        self._log(detail)
                    else:
                        self._log(f"End prior instance (taskkill exit {code}): {detail}")
                elif code not in (0, 128):
                    self._log(f"End prior instance: taskkill exit {code}")
                self._status.set("Previous application stopped. Preparing upgrade…")
                self._set_progress_pct(0.0)
                threading.Thread(target=self._worker, daemon=True).start()

            self.after(0, on_done)

        threading.Thread(target=work, daemon=True).start()

    def _worker(self) -> None:
        log = CallableLog(lambda m: self.after(0, lambda msg=m: self._log(msg)))

        def report_progress(p: float) -> None:
            self.after(0, lambda v=p: self._set_progress_pct(v))

        try:
            self.after(0, lambda: self._status.set("Uninstalling previous version (0–50%)…"))
            code, uninstall_hint = run_uninstall_silent(log, report_progress)
            if code == 2:
                self.after(
                    0,
                    lambda h=uninstall_hint: self._finish_install_error(
                        h or "Could not parse the uninstall program from the registry.", 2
                    ),
                )
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


def run_gui(installer: Path, parent_pid: int, parent_executable: str = "") -> int:
    if sys.platform != "win32":
        return 1
    inst = installer.expanduser().resolve()
    if not inst.is_file():
        messagebox.showerror(DISPLAY_NAME + " Update", f"Installer not found:\n{inst}")
        return 1
    app = UpdateMiniApp(inst, parent_pid, parent_executable)
    app.mainloop()
    cleanup_staged_temp_artifacts(Path(app._installer))
    return 1 if getattr(app, "_failed", False) else 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="DSS Tools update helper.")
    parser.add_argument("installer", type=Path, help="Path to DSSToolsSetup.exe")
    parser.add_argument("parent_pid", type=int, help="PID of the running DSS Tools instance")
    parser.add_argument(
        "parent_executable",
        nargs="?",
        default="",
        help="Resolved path of main app sys.executable when started (enables safe taskkill).",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="No window: optional safe taskkill then launch installer only (legacy / tests).",
    )
    args = parser.parse_args(argv)
    parent_exe = str(args.parent_executable or "").strip()
    if args.headless:
        return run_headless_legacy(Path(args.installer), args.parent_pid, parent_exe)
    if sys.platform != "win32":
        print("The DSS Tools updater only runs on Windows.", file=sys.stderr)
        return 1
    return run_gui(Path(args.installer), args.parent_pid, parent_exe)


if __name__ == "__main__":
    raise SystemExit(main())
