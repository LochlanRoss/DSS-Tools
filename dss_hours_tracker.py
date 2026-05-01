from __future__ import annotations

import argparse
import concurrent.futures
import difflib
import html
import hashlib
import io
import json
import os
import re
import sys
import tempfile
import threading
import tomllib
import urllib.error
import urllib.request
import zipfile
from importlib import metadata as importlib_metadata
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Callable, Iterable

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

try:
    import pythoncom
    import win32com.client
except ImportError:  # pragma: no cover - optional Windows integration
    pythoncom = None
    win32com = None


DATE_PATTERN = re.compile(r"(\d{4}-\d{2}-\d{2})")
REVISION_PATTERN = re.compile(r"(?:^|[\s_-])(?:rev(?:ision)?[\s_-]*|r[\s_-]*)(\d+)(?=$|[\s_-])", re.IGNORECASE)
PF_PATTERN = re.compile(r"\b(PF\d+(?:-\d+)?)\b", re.IGNORECASE)
WORKBOOK_CALC_ID_PATTERN = re.compile(br'\s(?:calcId|fullCalcOnLoad|forceFullCalc|calcCompleted)="[^"]*"')
BLOCK_START_ROWS = (25, 28, 31, 34)

LEFT_NAME_COLS = tuple(range(20, 28))   # T:AA
LEFT_HOUR_COLS = tuple(range(29, 32))   # AC:AE
RIGHT_NAME_COLS = tuple(range(29, 48))  # AC:AV
RIGHT_HOUR_COLS = tuple(range(50, 53))  # AX:AZ
CONFIG_FILENAME = "dss_hours_tracker_config.json"
DEFAULT_PROFILE_NAME = "Default"
APP_DIRNAME = "DSSHoursTracker"
CACHE_DIRNAME = "cache"
CACHE_RETENTION_DAYS = 7
HASH_CHECK_INTERVAL_MS = 300000
AUTO_OUTLOOK_SYNC_DELAY_MS = 60000
DEFAULT_HASH_POLL_MINUTES = 5
BUG_REPORT_EMAIL = "lross@jatechpowersystems.com"
MAX_PARALLEL_PARSE_WORKERS = 2
GITHUB_REPO_SLUG = "LochlanRoss/DSS-Viewer"
GITHUB_LATEST_RELEASE_URL = f"https://api.github.com/repos/{GITHUB_REPO_SLUG}/releases/latest"


class OperationCancelled(RuntimeError):
    pass


def discover_app_version() -> str:
    try:
        return importlib_metadata.version("dss-hours-tracker")
    except importlib_metadata.PackageNotFoundError:
        pass
    pyproject_path = Path(__file__).with_name("pyproject.toml")
    try:
        with pyproject_path.open("rb") as handle:
            payload = tomllib.load(handle)
    except (OSError, tomllib.TOMLDecodeError):
        return "0.0.0"
    project = payload.get("project", {})
    return str(project.get("version", "0.0.0")).strip() or "0.0.0"


def normalize_release_version(version_text: str) -> str:
    return version_text.strip().lstrip("vV")


def version_key(version_text: str) -> tuple[tuple[int, object], ...]:
    normalized = normalize_release_version(version_text)
    tokens = re.findall(r"\d+|[A-Za-z]+", normalized)
    if not tokens:
        return ((0, 0),)
    key: list[tuple[int, object]] = []
    for token in tokens:
        if token.isdigit():
            key.append((0, int(token)))
        else:
            key.append((1, token.casefold()))
    return tuple(key)


def is_newer_version(candidate_version: str, current_version: str) -> bool:
    return version_key(candidate_version) > version_key(current_version)


def parse_latest_release_payload(payload: dict) -> dict[str, object]:
    tag_name = str(payload.get("tag_name", "")).strip()
    version = normalize_release_version(tag_name)
    return {
        "tag_name": tag_name,
        "version": version,
        "name": str(payload.get("name", "")).strip(),
        "html_url": str(payload.get("html_url", "")).strip(),
        "published_at": str(payload.get("published_at", "")).strip(),
        "body": str(payload.get("body", "")),
        "asset_names": [
            str(asset.get("name", "")).strip()
            for asset in payload.get("assets", [])
            if isinstance(asset, dict) and str(asset.get("name", "")).strip()
        ],
    }


def fetch_latest_release_info(url: str = GITHUB_LATEST_RELEASE_URL, timeout: int = 10) -> dict[str, object]:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": f"dss-hours-tracker/{discover_app_version()}",
            "Accept": "application/vnd.github+json",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        raise RuntimeError(f"GitHub release check failed with HTTP {exc.code}.") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Could not reach GitHub to check for updates.") from exc
    except (OSError, json.JSONDecodeError, UnicodeDecodeError) as exc:
        raise RuntimeError("Could not read the GitHub release response.") from exc
    if not isinstance(payload, dict):
        raise RuntimeError("Unexpected GitHub release response.")
    return parse_latest_release_payload(payload)


APP_VERSION = discover_app_version()


@dataclass(frozen=True)
class DailyRecord:
    source_path: Path
    source_file: str
    work_date: date
    source_sheet: str
    employee: str
    st: float
    ot: float
    dt: float
    source_ranges: str

    @property
    def total(self) -> float:
        return round(self.st + self.ot + self.dt, 2)


@dataclass(frozen=True)
class WeeklyRecord:
    source_file: str
    week_start: date
    week_end: date
    employee: str
    st: float
    ot: float
    dt: float

    @property
    def total(self) -> float:
        return round(self.st + self.ot + self.dt, 2)


@dataclass(frozen=True)
class WeeklyRollupRow:
    source_file: str
    week_start: date
    week_end: date
    employee: str
    st: float
    ot: float
    dt: float
    row_type: str

    @property
    def total(self) -> float:
        return round(self.st + self.ot + self.dt, 2)


@dataclass(frozen=True)
class WeekTotalRow:
    week_start: date
    week_end: date
    st: float
    ot: float
    dt: float

    @property
    def total(self) -> float:
        return round(self.st + self.ot + self.dt, 2)


@dataclass(frozen=True)
class TrackerData:
    source_paths: list[Path]
    file_hashes: dict[Path, str]
    reused_paths: list[Path]
    reloaded_paths: list[Path]
    cache_status_by_path: dict[Path, str]
    daily_records: list[DailyRecord]
    employee_names: list[str]
    weekly_summary: list[WeeklyRecord]
    weekly_rollup: list[WeeklyRollupRow]
    week_totals: list[WeekTotalRow]
    combined_weekly_summary: list[WeeklyRecord]
    parse_warnings: list["SheetParseWarning"]
    workbook_health: list["WorkbookHealthItem"]


@dataclass(frozen=True)
class FormattingProfile:
    name: str
    st_threshold: float | None
    ot_threshold: float | None
    daily_st_threshold: float | None
    max_hours_per_day: float | None


@dataclass(frozen=True)
class ErrorFinding:
    employee: str
    week_start: date
    week_end: date
    hour_type: str
    threshold: float
    actual_total: float
    delta: float
    trigger_date: date
    trigger_day_st: float
    trigger_day_ot: float
    trigger_day_dt: float
    source_files: str
    reason: str
    breakdown: str


@dataclass(frozen=True)
class NameTypoWarning:
    employee: str
    similar_employee: str
    similarity: float
    locations: list[str]


@dataclass(frozen=True)
class EmailDraftRequest:
    employee: str
    email: str
    week_start: date
    week_end: date
    records: list[DailyRecord]


@dataclass(frozen=True)
class FilterSelection:
    mode: str
    value: str
    values: tuple[str, ...] = ()


@dataclass(frozen=True)
class AppSettings:
    disable_name_typo_notifications: bool = False
    hash_poll_minutes: int = DEFAULT_HASH_POLL_MINUTES
    show_daily_raw_tab: bool = True


@dataclass(frozen=True)
class SheetParseWarning:
    source_file: str
    source_sheet: str
    work_date: str
    issue: str
    details: str


@dataclass(frozen=True)
class WorkbookHealthItem:
    source_file: str
    status: str
    details: str


def get_app_root() -> Path:
    if os.name == "nt":
        base = os.environ.get("LOCALAPPDATA")
        if base:
            return Path(base) / APP_DIRNAME
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / APP_DIRNAME
    return Path.home() / f".{APP_DIRNAME}"


def ensure_app_directories() -> tuple[Path, Path]:
    app_root = get_app_root()
    cache_dir = app_root / CACHE_DIRNAME
    app_root.mkdir(parents=True, exist_ok=True)
    cache_dir.mkdir(parents=True, exist_ok=True)
    return app_root, cache_dir


def parse_sheet_date(sheet_name: str) -> date | None:
    match = DATE_PATTERN.search(sheet_name)
    if not match:
        return None
    return datetime.strptime(match.group(1), "%Y-%m-%d").date()


def parse_sheet_revision(sheet_name: str) -> int:
    match = REVISION_PATTERN.search(sheet_name)
    if not match:
        return 0
    return int(match.group(1))


def extract_pf_identifier(source_name: str) -> str:
    match = PF_PATTERN.search(source_name)
    if match:
        return match.group(1).upper()
    return Path(source_name).stem.strip() or source_name.strip()


def default_formatting_profiles() -> dict[str, FormattingProfile]:
    return {
        DEFAULT_PROFILE_NAME: FormattingProfile(
            name=DEFAULT_PROFILE_NAME,
            st_threshold=40.0,
            ot_threshold=10.0,
            daily_st_threshold=None,
            max_hours_per_day=None,
        )
    }


def read_config_payload(config_path: Path) -> dict:
    if not config_path.exists():
        return {}
    try:
        payload = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, TypeError, ValueError):
        return {}
    return payload if isinstance(payload, dict) else {}


def parse_threshold_value(raw_value: str) -> float | None:
    text = raw_value.strip()
    if not text:
        return None
    return float(text)


def load_formatting_profiles(config_path: Path) -> tuple[dict[str, FormattingProfile], str]:
    defaults = default_formatting_profiles()
    payload = read_config_payload(config_path)
    if not payload:
        return defaults, DEFAULT_PROFILE_NAME

    profiles: dict[str, FormattingProfile] = {}
    for item in payload.get("profiles", []):
        name = str(item.get("name", "")).strip()
        if not name:
            continue
        profiles[name] = FormattingProfile(
            name=name,
            st_threshold=item.get("st_threshold"),
            ot_threshold=item.get("ot_threshold"),
            daily_st_threshold=item.get("daily_st_threshold"),
            max_hours_per_day=item.get("max_hours_per_day"),
        )

    if not profiles:
        profiles = defaults

    current_profile_name = str(payload.get("current_profile", DEFAULT_PROFILE_NAME))
    if current_profile_name not in profiles:
        current_profile_name = sorted(profiles)[0]
    return profiles, current_profile_name


def save_formatting_profiles(
    config_path: Path,
    profiles: dict[str, FormattingProfile],
    current_profile_name: str,
) -> None:
    payload = read_config_payload(config_path)
    payload["current_profile"] = current_profile_name
    payload["profiles"] = [
        {
                "name": profile.name,
                "st_threshold": profile.st_threshold,
                "ot_threshold": profile.ot_threshold,
                "daily_st_threshold": profile.daily_st_threshold,
                "max_hours_per_day": profile.max_hours_per_day,
            }
        for profile in sorted(profiles.values(), key=lambda item: item.name.lower())
    ]
    config_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def load_employee_emails(config_path: Path) -> dict[str, str]:
    payload = read_config_payload(config_path)
    raw_map = payload.get("employee_emails", {})
    if not isinstance(raw_map, dict):
        return {}
    emails: dict[str, str] = {}
    for name, email_value in raw_map.items():
        employee = str(name).strip()
        email = str(email_value).strip()
        if employee:
            emails[employee] = email
    return emails


def save_employee_emails(config_path: Path, employee_emails: dict[str, str]) -> None:
    payload = read_config_payload(config_path)
    payload["employee_emails"] = {
        name: email.strip()
        for name, email in sorted(employee_emails.items(), key=lambda item: item[0].lower())
        if name.strip()
    }
    config_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def load_ignored_name_typos(config_path: Path) -> set[str]:
    payload = read_config_payload(config_path)
    raw_values = payload.get("ignored_name_typos", [])
    if not isinstance(raw_values, list):
        return set()
    return {str(value).strip() for value in raw_values if str(value).strip()}


def save_ignored_name_typos(config_path: Path, ignored_name_typos: set[str]) -> None:
    payload = read_config_payload(config_path)
    payload["ignored_name_typos"] = sorted(value for value in ignored_name_typos if value.strip())
    config_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def load_employee_groups(config_path: Path) -> dict[str, list[str]]:
    payload = read_config_payload(config_path)
    raw_groups = payload.get("employee_groups", {})
    if not isinstance(raw_groups, dict):
        return {}
    groups: dict[str, list[str]] = {}
    for group_name, employees in raw_groups.items():
        name = str(group_name).strip()
        if not name:
            continue
        if isinstance(employees, list):
            groups[name] = sorted({str(employee).strip() for employee in employees if str(employee).strip()})
    return groups


def save_employee_groups(config_path: Path, employee_groups: dict[str, list[str]]) -> None:
    payload = read_config_payload(config_path)
    payload["employee_groups"] = {
        name: sorted({employee.strip() for employee in employees if employee.strip()})
        for name, employees in sorted(employee_groups.items(), key=lambda item: item[0].lower())
        if name.strip()
    }
    config_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def load_employee_notes(config_path: Path) -> dict[str, str]:
    payload = read_config_payload(config_path)
    raw_map = payload.get("employee_notes", {})
    if not isinstance(raw_map, dict):
        return {}
    notes: dict[str, str] = {}
    for name, note_value in raw_map.items():
        employee = str(name).strip()
        note = str(note_value).strip()
        if employee:
            notes[employee] = note
    return notes


def save_employee_notes(config_path: Path, employee_notes: dict[str, str]) -> None:
    payload = read_config_payload(config_path)
    payload["employee_notes"] = {
        name: note.strip()
        for name, note in sorted(employee_notes.items(), key=lambda item: item[0].lower())
        if name.strip()
    }
    config_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def load_job_presets(config_path: Path) -> dict[str, str]:
    payload = read_config_payload(config_path)
    raw_map = payload.get("job_presets", {})
    if not isinstance(raw_map, dict):
        return {}
    presets: dict[str, str] = {}
    for job_name, profile_name in raw_map.items():
        job = str(job_name).strip()
        profile = str(profile_name).strip()
        if job and profile:
            presets[job] = profile
    return presets


def save_job_presets(config_path: Path, job_presets: dict[str, str]) -> None:
    payload = read_config_payload(config_path)
    payload["job_presets"] = {
        job: profile
        for job, profile in sorted(job_presets.items(), key=lambda item: item[0].lower())
        if job.strip() and profile.strip()
    }
    config_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def default_email_body_template() -> str:
    return (
        "<p>Hi {first_name},</p>\n"
        "<p>Here are your DSS hours for the week of {week_start} to {week_end}. "
        "Please use these details to complete your timesheet.</p>\n"
        "{hours_table}\n"
        "<p>Thanks.</p>"
    )


def load_email_templates(config_path: Path) -> tuple[str, str]:
    payload = read_config_payload(config_path)
    subject_template = str(payload.get("email_subject_template", "Timesheet Reminder - Week of {week_start}")).strip()
    if not subject_template:
        subject_template = "Timesheet Reminder - Week of {week_start}"
    body_template = str(payload.get("email_body_template", default_email_body_template()))
    if not body_template.strip():
        body_template = default_email_body_template()
    return subject_template, body_template


def save_email_templates(config_path: Path, subject_template: str, body_template: str) -> None:
    payload = read_config_payload(config_path)
    payload["email_subject_template"] = subject_template
    payload["email_body_template"] = body_template
    config_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def load_app_settings(config_path: Path) -> AppSettings:
    payload = read_config_payload(config_path)
    raw_settings = payload.get("app_settings", {})
    if not isinstance(raw_settings, dict):
        return AppSettings()

    raw_minutes = raw_settings.get("hash_poll_minutes", DEFAULT_HASH_POLL_MINUTES)
    try:
        hash_poll_minutes = max(1, int(raw_minutes))
    except (TypeError, ValueError):
        hash_poll_minutes = DEFAULT_HASH_POLL_MINUTES

    return AppSettings(
        disable_name_typo_notifications=bool(raw_settings.get("disable_name_typo_notifications", False)),
        hash_poll_minutes=hash_poll_minutes,
        show_daily_raw_tab=bool(raw_settings.get("show_daily_raw_tab", True)),
    )


def save_app_settings(config_path: Path, settings: AppSettings) -> None:
    payload = read_config_payload(config_path)
    payload["app_settings"] = {
        "disable_name_typo_notifications": settings.disable_name_typo_notifications,
        "hash_poll_minutes": settings.hash_poll_minutes,
        "show_daily_raw_tab": settings.show_daily_raw_tab,
    }
    config_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def remove_config_keys(config_path: Path, keys: Iterable[str]) -> None:
    payload = read_config_payload(config_path)
    changed = False
    for key in keys:
        if key in payload:
            del payload[key]
            changed = True
    if changed:
        if payload:
            config_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        elif config_path.exists():
            config_path.unlink()


def load_table_layouts(config_path: Path) -> dict[str, dict]:
    payload = read_config_payload(config_path)
    raw_layouts = payload.get("table_layouts", {})
    if not isinstance(raw_layouts, dict):
        return {}

    layouts: dict[str, dict] = {}
    for table_id, layout in raw_layouts.items():
        if not isinstance(layout, dict):
            continue
        visible_columns = layout.get("visible_columns", [])
        column_widths = layout.get("column_widths", {})
        sort_column = str(layout.get("sort_column", "")).strip()
        sort_descending = bool(layout.get("sort_descending", False))
        layouts[str(table_id)] = {
            "visible_columns": [str(column) for column in visible_columns if str(column).strip()],
            "column_widths": {
                str(column): int(width)
                for column, width in column_widths.items()
                if str(column).strip() and isinstance(width, (int, float)) and int(width) > 0
            },
            "sort_column": sort_column,
            "sort_descending": sort_descending,
        }
    return layouts


def save_table_layout(
    config_path: Path,
    table_id: str,
    visible_columns: list[str],
    column_widths: dict[str, int],
    sort_column: str = "",
    sort_descending: bool = False,
) -> None:
    payload = read_config_payload(config_path)
    raw_layouts = payload.get("table_layouts", {})
    layouts = raw_layouts if isinstance(raw_layouts, dict) else {}
    layouts[table_id] = {
        "visible_columns": list(visible_columns),
        "column_widths": {
            column: int(width)
            for column, width in column_widths.items()
            if int(width) > 0
        },
        "sort_column": sort_column,
        "sort_descending": sort_descending,
    }
    payload["table_layouts"] = layouts
    config_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def clear_cache_files(cache_dir: Path) -> int:
    deleted = 0
    if not cache_dir.exists():
        return deleted
    for cache_file in cache_dir.glob("*.json"):
        try:
            cache_file.unlink()
            deleted += 1
        except OSError:
            continue
    return deleted


def source_cache_key(source_path: Path) -> str:
    normalized = normalize_windows_path(str(source_path.resolve()))
    return hashlib.blake2b(normalized.encode("utf-8"), digest_size=16).hexdigest()


def cache_file_path(cache_dir: Path, source_path: Path) -> Path:
    return cache_dir / f"{source_cache_key(source_path)}.json"


def serialize_daily_record(record: DailyRecord) -> dict:
    return {
        "source_path": str(record.source_path),
        "source_file": record.source_file,
        "work_date": record.work_date.isoformat(),
        "source_sheet": record.source_sheet,
        "employee": record.employee,
        "st": record.st,
        "ot": record.ot,
        "dt": record.dt,
        "source_ranges": record.source_ranges,
    }


def deserialize_daily_record(payload: dict) -> DailyRecord:
    return DailyRecord(
        source_path=Path(payload["source_path"]),
        source_file=str(payload["source_file"]),
        work_date=datetime.strptime(str(payload["work_date"]), "%Y-%m-%d").date(),
        source_sheet=str(payload["source_sheet"]),
        employee=str(payload["employee"]),
        st=float(payload["st"]),
        ot=float(payload["ot"]),
        dt=float(payload["dt"]),
        source_ranges=str(payload["source_ranges"]),
    )


def serialize_sheet_parse_warning(warning: SheetParseWarning) -> dict:
    return {
        "source_file": warning.source_file,
        "source_sheet": warning.source_sheet,
        "work_date": warning.work_date,
        "issue": warning.issue,
        "details": warning.details,
    }


def deserialize_sheet_parse_warning(payload: dict) -> SheetParseWarning:
    return SheetParseWarning(
        source_file=str(payload["source_file"]),
        source_sheet=str(payload["source_sheet"]),
        work_date=str(payload["work_date"]),
        issue=str(payload["issue"]),
        details=str(payload["details"]),
    )


def serialize_workbook_health_item(item: WorkbookHealthItem) -> dict:
    return {
        "source_file": item.source_file,
        "status": item.status,
        "details": item.details,
    }


def deserialize_workbook_health_item(payload: dict) -> WorkbookHealthItem:
    return WorkbookHealthItem(
        source_file=str(payload["source_file"]),
        status=str(payload["status"]),
        details=str(payload["details"]),
    )


def purge_stale_cache(cache_dir: Path, now: datetime | None = None) -> None:
    current_time = now or datetime.now()
    cutoff = current_time - timedelta(days=CACHE_RETENTION_DAYS)
    for cache_file in cache_dir.glob("*.json"):
        try:
            modified = datetime.fromtimestamp(cache_file.stat().st_mtime)
        except OSError:
            continue
        if modified < cutoff:
            try:
                cache_file.unlink()
            except OSError:
                pass


def load_cached_daily_records(cache_dir: Path, source_path: Path, file_hash: str) -> list[DailyRecord] | None:
    cache_path = cache_file_path(cache_dir, source_path)
    if not cache_path.exists():
        return None
    try:
        payload = json.loads(cache_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, TypeError, ValueError):
        return None
    if payload.get("file_hash") != file_hash:
        return None
    if normalize_windows_path(str(payload.get("source_path", ""))) != normalize_windows_path(str(source_path)):
        return None
    records_payload = payload.get("records", [])
    if not isinstance(records_payload, list):
        return None
    try:
        return [deserialize_daily_record(item) for item in records_payload]
    except (KeyError, TypeError, ValueError):
        return None


def load_cached_source_analysis(
    cache_dir: Path,
    source_path: Path,
    file_hash: str,
) -> tuple[list[SheetParseWarning], list[WorkbookHealthItem]] | None:
    cache_path = cache_file_path(cache_dir, source_path)
    if not cache_path.exists():
        return None
    try:
        payload = json.loads(cache_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, TypeError, ValueError):
        return None
    if payload.get("file_hash") != file_hash:
        return None
    if normalize_windows_path(str(payload.get("source_path", ""))) != normalize_windows_path(str(source_path)):
        return None
    if "parse_warnings" not in payload or "workbook_health" not in payload:
        return None
    warnings_payload = payload.get("parse_warnings", [])
    health_payload = payload.get("workbook_health", [])
    if not isinstance(warnings_payload, list) or not isinstance(health_payload, list):
        return None
    try:
        return (
            [deserialize_sheet_parse_warning(item) for item in warnings_payload],
            [deserialize_workbook_health_item(item) for item in health_payload],
        )
    except (KeyError, TypeError, ValueError):
        return None


def save_cached_daily_records(
    cache_dir: Path,
    source_path: Path,
    file_hash: str,
    records: list[DailyRecord],
    parse_warnings: list[SheetParseWarning] | None = None,
    workbook_health: list[WorkbookHealthItem] | None = None,
) -> None:
    cache_path = cache_file_path(cache_dir, source_path)
    payload = {
        "source_path": str(source_path),
        "file_hash": file_hash,
        "cached_at": datetime.now().isoformat(),
        "records": [serialize_daily_record(record) for record in records],
        "parse_warnings": [serialize_sheet_parse_warning(warning) for warning in (parse_warnings or [])],
        "workbook_health": [serialize_workbook_health_item(item) for item in (workbook_health or [])],
    }
    cache_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def is_alert_triggered(st: float, ot: float, dt: float, profile: FormattingProfile) -> bool:
    return (
        (profile.st_threshold is not None and st > profile.st_threshold)
        or (profile.ot_threshold is not None and ot > profile.ot_threshold)
    )


def build_error_findings(records: Iterable[DailyRecord], profile: FormattingProfile) -> list[ErrorFinding]:
    employee_week_records: dict[tuple[str, date], list[DailyRecord]] = {}
    for record in records:
        week_start = monday_week_start(record.work_date)
        employee_week_records.setdefault((record.employee, week_start), []).append(record)

    findings: list[ErrorFinding] = []
    for (employee, week_start), week_records in sorted(employee_week_records.items(), key=lambda item: (item[0][1], item[0][0])):
        week_end = week_start + timedelta(days=6)
        daily_totals: list[dict[str, object]] = []
        by_day: dict[date, dict[str, object]] = {}
        for record in sorted(week_records, key=lambda item: (item.work_date, item.source_file, item.source_sheet)):
            bucket = by_day.setdefault(
                record.work_date,
                {
                    "date": record.work_date,
                    "st": 0.0,
                    "ot": 0.0,
                    "dt": 0.0,
                    "sources": set(),
                },
            )
            bucket["st"] = float(bucket["st"]) + record.st
            bucket["ot"] = float(bucket["ot"]) + record.ot
            bucket["dt"] = float(bucket["dt"]) + record.dt
            cast_sources = bucket["sources"]
            if isinstance(cast_sources, set):
                cast_sources.add(record.source_file)
        for work_date in sorted(by_day):
            daily_totals.append(by_day[work_date])

        weekly_st = round(sum(float(day["st"]) for day in daily_totals), 2)
        weekly_ot = round(sum(float(day["ot"]) for day in daily_totals), 2)
        source_files = ", ".join(sorted({source for day in daily_totals for source in day["sources"]}))  # type: ignore[index]
        breakdown = " | ".join(
            (
                f"{day['date'].isoformat()} "
                f"ST {fmt_hours(float(day['st']))} "
                f"OT {fmt_hours(float(day['ot']))} "
                f"DT {fmt_hours(float(day['dt']))}"
            )
            for day in daily_totals
        )

        for hour_type, threshold, actual_total in (
            ("ST", profile.st_threshold, weekly_st),
            ("OT", profile.ot_threshold, weekly_ot),
        ):
            if threshold is None or actual_total <= threshold:
                continue
            cumulative = 0.0
            trigger_day = daily_totals[-1]
            for day in daily_totals:
                cumulative += float(day[hour_type.lower()])
                if cumulative > threshold:
                    trigger_day = day
                    break
            delta = round(actual_total - threshold, 2)
            reason = (
                f"{hour_type} weekly total {fmt_hours(actual_total)} exceeded the limit "
                f"{fmt_hours(threshold)} by {fmt_hours(delta)}."
            )
            findings.append(
                ErrorFinding(
                    employee=employee,
                    week_start=week_start,
                    week_end=week_end,
                    hour_type=hour_type,
                    threshold=threshold,
                    actual_total=actual_total,
                    delta=delta,
                    trigger_date=trigger_day["date"],  # type: ignore[arg-type]
                    trigger_day_st=round(float(trigger_day["st"]), 2),
                    trigger_day_ot=round(float(trigger_day["ot"]), 2),
                    trigger_day_dt=round(float(trigger_day["dt"]), 2),
                    source_files=source_files,
                    reason=reason,
                    breakdown=breakdown,
                )
            )

        for day in daily_totals:
            day_st = round(float(day["st"]), 2)
            day_ot = round(float(day["ot"]), 2)
            day_dt = round(float(day["dt"]), 2)
            day_total = round(day_st + day_ot + day_dt, 2)
            trigger_date = day["date"]  # type: ignore[assignment]
            day_sources = ", ".join(sorted(day["sources"]))  # type: ignore[arg-type]
            day_breakdown = (
                f"{trigger_date.isoformat()} "
                f"ST {fmt_hours(day_st)} OT {fmt_hours(day_ot)} DT {fmt_hours(day_dt)} Total {fmt_hours(day_total)}"
            )

            if profile.daily_st_threshold is not None and day_st > profile.daily_st_threshold:
                delta = round(day_st - profile.daily_st_threshold, 2)
                findings.append(
                    ErrorFinding(
                        employee=employee,
                        week_start=week_start,
                        week_end=week_end,
                        hour_type="Daily ST",
                        threshold=profile.daily_st_threshold,
                        actual_total=day_st,
                        delta=delta,
                        trigger_date=trigger_date,
                        trigger_day_st=day_st,
                        trigger_day_ot=day_ot,
                        trigger_day_dt=day_dt,
                        source_files=day_sources,
                        reason=(
                            f"Daily ST total {fmt_hours(day_st)} exceeded the limit "
                            f"{fmt_hours(profile.daily_st_threshold)} by {fmt_hours(delta)}."
                        ),
                        breakdown=day_breakdown,
                    )
                )

            if profile.max_hours_per_day is not None and day_total > profile.max_hours_per_day:
                delta = round(day_total - profile.max_hours_per_day, 2)
                findings.append(
                    ErrorFinding(
                        employee=employee,
                        week_start=week_start,
                        week_end=week_end,
                        hour_type="Daily Total",
                        threshold=profile.max_hours_per_day,
                        actual_total=day_total,
                        delta=delta,
                        trigger_date=trigger_date,
                        trigger_day_st=day_st,
                        trigger_day_ot=day_ot,
                        trigger_day_dt=day_dt,
                        source_files=day_sources,
                        reason=(
                            f"Daily total {fmt_hours(day_total)} exceeded the max-hours limit "
                            f"{fmt_hours(profile.max_hours_per_day)} by {fmt_hours(delta)}."
                        ),
                        breakdown=day_breakdown,
                    )
                )
    return findings


def filter_employee_names(
    employee_names: Iterable[str],
    filter_selection: FilterSelection,
    employee_groups: dict[str, list[str]],
) -> set[str]:
    mode = filter_selection.mode
    value = filter_selection.value.strip()
    names = {employee for employee in employee_names if employee.strip()}
    if filter_selection.values:
        return {employee for employee in filter_selection.values if employee in names}
    if mode == "employee" and value:
        return {value} if value in names else set()
    if mode == "group" and value:
        return {employee for employee in employee_groups.get(value, []) if employee in names}
    return names


def normalize_person_name(name: str) -> str:
    return " ".join(name.split()).casefold()


def typo_warning_key(employee: str, similar_employee: str) -> str:
    return f"{normalize_person_name(employee)}|{normalize_person_name(similar_employee)}"


def find_potential_name_typos(
    unresolved_names: Iterable[str],
    all_employee_names: Iterable[str],
    daily_records: Iterable[DailyRecord],
    similarity_threshold: float = 0.82,
) -> list[NameTypoWarning]:
    all_names = sorted({name for name in all_employee_names if name.strip()})
    normalized_names = {name: normalize_person_name(name) for name in all_names}
    warnings: list[NameTypoWarning] = []

    for unresolved in sorted({name for name in unresolved_names if name.strip()}):
        unresolved_normalized = normalize_person_name(unresolved)
        best_match = ""
        best_similarity = 0.0
        for candidate in all_names:
            if candidate == unresolved:
                continue
            similarity = difflib.SequenceMatcher(None, unresolved_normalized, normalized_names[candidate]).ratio()
            if similarity > best_similarity:
                best_similarity = similarity
                best_match = candidate

        if not best_match or best_similarity < similarity_threshold:
            continue

        locations = []
        for record in daily_records:
            if record.employee != unresolved:
                continue
            locations.append(f"{record.work_date.isoformat()} | {record.source_sheet} | {record.source_file}")

        warnings.append(
            NameTypoWarning(
                employee=unresolved,
                similar_employee=best_match,
                similarity=best_similarity,
                locations=locations,
            )
        )
    return warnings


def extract_smtp_address(recipient) -> str:
    try:
        address_entry = recipient.AddressEntry
    except Exception:
        address_entry = None

    if address_entry is not None:
        try:
            exchange_user = address_entry.GetExchangeUser()
            if exchange_user and getattr(exchange_user, "PrimarySmtpAddress", ""):
                return str(exchange_user.PrimarySmtpAddress).strip()
        except Exception:
            pass
        try:
            exchange_dist = address_entry.GetExchangeDistributionList()
            if exchange_dist and getattr(exchange_dist, "PrimarySmtpAddress", ""):
                return str(exchange_dist.PrimarySmtpAddress).strip()
        except Exception:
            pass

    address = getattr(recipient, "Address", "")
    return str(address).strip()


def lookup_outlook_emails(
    employee_names: Iterable[str],
    should_cancel: Callable[[], bool] | None = None,
) -> dict[str, str]:
    if pythoncom is None or win32com is None:
        raise RuntimeError("Outlook lookup requires pywin32 and desktop Outlook.")

    names = [name for name in employee_names if name.strip()]
    if not names:
        return {}
    check_cancel = should_cancel or (lambda: False)

    pythoncom.CoInitialize()
    try:
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
        except Exception as exc:
            raise RuntimeError("Could not open Outlook to query email addresses.") from exc

        results: dict[str, str] = {}
        for employee in names:
            if check_cancel():
                raise OperationCancelled("Cancelled Outlook email sync.")
            try:
                recipient = namespace.CreateRecipient(employee)
                recipient.Resolve()
            except Exception:
                continue
            if not getattr(recipient, "Resolved", False):
                continue
            email = extract_smtp_address(recipient)
            if email:
                results[employee] = email
        return results
    finally:
        pythoncom.CoUninitialize()


def format_week_label(week_start: date, week_end: date) -> str:
    return f"{week_start.isoformat()} to {week_end.isoformat()}"


def collect_week_ranges(records: Iterable[DailyRecord]) -> list[tuple[date, date]]:
    week_starts = sorted({monday_week_start(record.work_date) for record in records})
    return [(week_start, week_start + timedelta(days=6)) for week_start in week_starts]


def records_for_week(records: Iterable[DailyRecord], week_start: date) -> list[DailyRecord]:
    week_end = week_start + timedelta(days=6)
    return [record for record in records if week_start <= record.work_date <= week_end]


def extract_pf_number(source_name: str) -> str | None:
    match = PF_PATTERN.search(source_name)
    if not match:
        return None
    return match.group(1).upper()


def pf_numbers_for_records(records: Iterable[DailyRecord]) -> str:
    pf_numbers = sorted({pf_number for record in records if (pf_number := extract_pf_number(record.source_file))})
    return ", ".join(pf_numbers)


def build_email_draft_requests(
    records: Iterable[DailyRecord],
    employee_emails: dict[str, str],
    week_start: date,
) -> list[EmailDraftRequest]:
    grouped: dict[str, list[DailyRecord]] = {}
    week_records = records_for_week(records, week_start)
    for record in week_records:
        grouped.setdefault(record.employee, []).append(record)

    requests: list[EmailDraftRequest] = []
    week_end = week_start + timedelta(days=6)
    for employee in sorted(grouped):
        requests.append(
            EmailDraftRequest(
                employee=employee,
                email=employee_emails.get(employee, "").strip(),
                week_start=week_start,
                week_end=week_end,
                records=sorted(grouped[employee], key=lambda item: (item.work_date, item.source_file, item.source_sheet)),
            )
        )
    return requests


def format_email_subject(
    template: str,
    employee: str,
    week_start: date,
    week_end: date,
    records: Iterable[DailyRecord] | None = None,
) -> str:
    subject_template = template.strip() or "Timesheet Reminder - Week of {week_start}"
    first_name = employee.strip().split()[0] if employee.strip() else employee
    pf_numbers = pf_numbers_for_records(records or [])
    if pf_numbers and "{pf_numbers}" not in subject_template:
        subject_template = f"{subject_template} - {pf_numbers}"
    return subject_template.format(
        employee=employee,
        first_name=first_name,
        week_start=week_start.isoformat(),
        week_end=week_end.isoformat(),
        pf_numbers=pf_numbers,
    )


def build_hours_table_html(records: list[DailyRecord]) -> str:
    total_st = round(sum(record.st for record in records), 2)
    total_ot = round(sum(record.ot for record in records), 2)
    total_dt = round(sum(record.dt for record in records), 2)
    total_hours = round(total_st + total_ot + total_dt, 2)

    rows = []
    for record in records:
        rows.append(
            "<tr>"
            f"<td>{html.escape(record.work_date.isoformat())}</td>"
            f"<td>{html.escape(record.work_date.strftime('%A'))}</td>"
            f"<td>{html.escape(record.source_file)}</td>"
            f"<td>{html.escape(record.source_sheet)}</td>"
            f"<td>{html.escape(fmt_hours(record.st))}</td>"
            f"<td>{html.escape(fmt_hours(record.ot))}</td>"
            f"<td>{html.escape(fmt_hours(record.dt))}</td>"
            f"<td>{html.escape(fmt_hours(record.total))}</td>"
            "</tr>"
        )
    rows_html = "".join(rows)

    return (
        "<table border='1' cellspacing='0' cellpadding='4' style='border-collapse:collapse;'>"
        "<thead><tr>"
        "<th>Date</th><th>Day</th><th>Source File</th><th>Source Sheet</th><th>ST</th><th>OT</th><th>DT</th><th>Total</th>"
        "</tr></thead>"
        f"<tbody>{rows_html}</tbody>"
        "<tfoot><tr>"
        "<th colspan='4'>Week Total</th>"
        f"<th>{html.escape(fmt_hours(total_st))}</th>"
        f"<th>{html.escape(fmt_hours(total_ot))}</th>"
        f"<th>{html.escape(fmt_hours(total_dt))}</th>"
        f"<th>{html.escape(fmt_hours(total_hours))}</th>"
        "</tr></tfoot>"
        "</table>"
    )


def build_email_html(
    employee: str,
    week_start: date,
    week_end: date,
    records: list[DailyRecord],
    body_template: str,
) -> str:
    first_name = employee.strip().split()[0] if employee.strip() else employee
    template = body_template.strip() or default_email_body_template()
    return template.format(
        employee=html.escape(employee),
        first_name=html.escape(first_name),
        week_start=html.escape(week_start.isoformat()),
        week_end=html.escape(week_end.isoformat()),
        hours_table=build_hours_table_html(records),
    )


def build_bug_report_html(
    current_profile_name: str,
    app_root: Path,
    snapshot_path: Path,
    loaded_sources: list[Path],
    cache_status_by_path: dict[Path, str],
) -> str:
    source_items = "".join(
        f"<li>{html.escape(str(path))} ({html.escape(cache_status_by_path.get(path, 'Unknown'))})</li>"
        for path in loaded_sources
    ) or "<li>None loaded</li>"
    return (
        "<p>Hi Lochlan,</p>"
        "<p>Please find a bug report for DSS Hours Tracker below.</p>"
        "<p><strong>Summary:</strong><br>[Describe the bug briefly]</p>"
        "<p><strong>What I was trying to do:</strong><br>[Describe the task]</p>"
        "<p><strong>What happened:</strong><br>[Describe the actual result or error]</p>"
        "<p><strong>What I expected to happen:</strong><br>[Describe the expected result]</p>"
        "<p><strong>Steps to reproduce:</strong><br>"
        "1. [Step one]<br>"
        "2. [Step two]<br>"
        "3. [Step three]</p>"
        "<p><strong>When it happened:</strong><br>[Date / time]</p>"
        "<p><strong>Relevant screenshots or notes:</strong><br>[Add anything helpful here]</p>"
        "<hr>"
        "<p><strong>Attached diagnostic snapshot:</strong><br>"
        f"{html.escape(str(snapshot_path))}</p>"
        "<p><strong>Current profile:</strong><br>"
        f"{html.escape(current_profile_name)}</p>"
        "<p><strong>App data folder:</strong><br>"
        f"{html.escape(str(app_root))}</p>"
        "<p><strong>Loaded DSS files:</strong></p>"
        f"<ul>{source_items}</ul>"
    )


def create_bug_report_draft(
    recipient_email: str,
    subject: str,
    html_body: str,
    attachment_path: Path | None = None,
) -> None:
    if pythoncom is None or win32com is None:
        raise RuntimeError("Bug report draft creation requires desktop Outlook with pywin32 available.")

    pythoncom.CoInitialize()
    try:
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as exc:
            raise RuntimeError("Could not open Outlook to create the bug report draft.") from exc

        mail_item = outlook.CreateItem(0)
        mail_item.To = recipient_email
        mail_item.Subject = subject
        mail_item.HTMLBody = html_body
        if attachment_path is not None and attachment_path.exists():
            mail_item.Attachments.Add(str(attachment_path))
        mail_item.Save()
    finally:
        pythoncom.CoUninitialize()


def create_outlook_drafts(
    requests: list[EmailDraftRequest],
    subject_template: str,
    body_template: str,
) -> tuple[int, list[str]]:
    if pythoncom is None or win32com is None:
        raise RuntimeError("Outlook draft creation requires desktop Outlook with pywin32 available.")

    created = 0
    skipped: list[str] = []
    pythoncom.CoInitialize()
    try:
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as exc:
            raise RuntimeError("Could not open Outlook to create draft emails.") from exc

        for request in requests:
            if not request.email:
                skipped.append(request.employee)
                continue
            mail_item = outlook.CreateItem(0)
            mail_item.To = request.email
            mail_item.Subject = format_email_subject(
                subject_template,
                request.employee,
                request.week_start,
                request.week_end,
                request.records,
            )
            mail_item.HTMLBody = build_email_html(
                request.employee,
                request.week_start,
                request.week_end,
                request.records,
                body_template,
            )
            mail_item.Save()
            created += 1
    finally:
        pythoncom.CoUninitialize()
    return created, skipped


def select_preferred_dated_sheets(sheet_names: Iterable[str]) -> list[tuple[date, str]]:
    best_by_date: dict[date, tuple[int, str]] = {}
    for sheet_name in sheet_names:
        sheet_date = parse_sheet_date(sheet_name)
        if not sheet_date:
            continue
        revision = parse_sheet_revision(sheet_name)
        current = best_by_date.get(sheet_date)
        if current is None or revision > current[0] or (revision == current[0] and sheet_name.strip() > current[1].strip()):
            best_by_date[sheet_date] = (revision, sheet_name)
    return [(sheet_date, best_by_date[sheet_date][1]) for sheet_date in sorted(best_by_date)]


def is_text_name(value: object) -> bool:
    return isinstance(value, str) and bool(value.strip())


def to_number(value: object) -> float:
    if value in (None, "") or isinstance(value, bool):
        return 0.0
    if isinstance(value, (int, float, Decimal)):
        return round(float(value), 2)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return 0.0
        try:
            return round(float(Decimal(text)), 2)
        except (InvalidOperation, ValueError):
            return 0.0
    return 0.0


def detect_employee_name(ws, start_row: int, column_indexes: Iterable[int]) -> str | None:
    for col_idx in reversed(tuple(column_indexes)):
        for row_idx in range(start_row, start_row + 3):
            value = ws.cell(row=row_idx, column=col_idx).value
            if is_text_name(value):
                return value.strip()
    return None


def sum_hour_row(ws, row_idx: int, hour_columns: Iterable[int]) -> float:
    return round(sum(to_number(ws.cell(row=row_idx, column=col_idx).value) for col_idx in hour_columns), 2)


def build_source_ranges(label: str, start_row: int, name_cols: tuple[int, ...], hour_cols: tuple[int, ...]) -> str:
    return (
        f"{label} name {get_column_letter(name_cols[0])}{start_row}:{get_column_letter(name_cols[-1])}{start_row + 2}; "
        f"{label} hours {get_column_letter(hour_cols[0])}{start_row}:{get_column_letter(hour_cols[-1])}{start_row + 2}"
    )


def find_name_cell(ws, start_row: int, column_indexes: Iterable[int]) -> tuple[str, str] | tuple[None, None]:
    for col_idx in reversed(tuple(column_indexes)):
        for row_idx in range(start_row, start_row + 3):
            value = ws.cell(row=row_idx, column=col_idx).value
            if is_text_name(value):
                return value.strip(), f"{get_column_letter(col_idx)}{row_idx}"
    return None, None


def process_workbook_bytes(
    source_path: Path,
    workbook_bytes: bytes,
    progress_callback: Callable[[float, str], None] | None = None,
    should_cancel: Callable[[], bool] | None = None,
    preview_callback: Callable[[list[DailyRecord], list[SheetParseWarning], list[WorkbookHealthItem], str], None] | None = None,
) -> tuple[list[DailyRecord], list[SheetParseWarning], list[WorkbookHealthItem]]:
    emit_progress = progress_callback or (lambda _fraction, _message: None)
    check_cancel = should_cancel or (lambda: False)

    def raise_if_cancelled() -> None:
        if check_cancel():
            raise OperationCancelled(f"Cancelled {source_path.name}")

    raise_if_cancelled()
    workbook = load_workbook(io.BytesIO(workbook_bytes), data_only=True, read_only=True)
    records: list[DailyRecord] = []
    warnings: list[SheetParseWarning] = []
    health: list[WorkbookHealthItem] = []

    try:
        raise_if_cancelled()
        emit_progress(0.1, f"Opened {source_path.name}")
        selected_sheets = sorted(
            select_preferred_dated_sheets(workbook.sheetnames),
            key=lambda item: (item[0], item[1].strip().casefold()),
            reverse=True,
        )
        emit_progress(0.2, f"Selected sheets for {source_path.name}")
        dated_sheets = [sheet_name for sheet_name in workbook.sheetnames if parse_sheet_date(sheet_name) is not None]
        if not dated_sheets:
            health.append(
                WorkbookHealthItem(
                    source_file=source_path.name,
                    status="Warning",
                    details="No dated DSS sheets matching YYYY-MM-DD were found.",
                )
            )
        if "onedrive" in str(source_path).lower() or "sharepoint" in str(source_path).lower():
            health.append(
                WorkbookHealthItem(
                    source_file=source_path.name,
                    status="Info",
                    details="Workbook is in a OneDrive / SharePoint path and may be temporarily locked during sync or save.",
                )
            )

        by_date: dict[date, list[str]] = {}
        for sheet_name in dated_sheets:
            sheet_date = parse_sheet_date(sheet_name)
            if sheet_date is not None:
                by_date.setdefault(sheet_date, []).append(sheet_name)
        preferred_sheet_by_date = dict(selected_sheets)
        for sheet_date, sheet_names in sorted(by_date.items()):
            if len(sheet_names) > 1:
                health.append(
                    WorkbookHealthItem(
                        source_file=source_path.name,
                        status="Info",
                        details=(
                            f"{sheet_date.isoformat()} has {len(sheet_names)} revision candidate sheets. "
                            f"Using '{preferred_sheet_by_date.get(sheet_date, '')}'."
                        ),
                    )
                )

        total_sheets = max(len(selected_sheets), 1)
        preview_emitted = False
        recent_week_starts = sorted({monday_week_start(sheet_date) for sheet_date, _sheet_name in selected_sheets}, reverse=True)
        preview_cutoff = recent_week_starts[1] if len(recent_week_starts) >= 2 else (recent_week_starts[0] if recent_week_starts else None)
        preview_sheet_count = sum(
            1
            for sheet_date, _sheet_name in selected_sheets
            if preview_cutoff is not None and monday_week_start(sheet_date) >= preview_cutoff
        )
        for sheet_index, (sheet_date, sheet_name) in enumerate(selected_sheets, start=1):
            raise_if_cancelled()
            ws = workbook[sheet_name]
            seen_names: dict[str, int] = {}
            for start_row in BLOCK_START_ROWS:
                for label, name_cols, hour_cols in (
                    ("Left", LEFT_NAME_COLS, LEFT_HOUR_COLS),
                    ("Right", RIGHT_NAME_COLS, RIGHT_HOUR_COLS),
                ):
                    raise_if_cancelled()
                    employee, cell_ref = find_name_cell(ws, start_row, name_cols)
                    st = sum_hour_row(ws, start_row, hour_cols)
                    ot = sum_hour_row(ws, start_row + 1, hour_cols)
                    dt = sum_hour_row(ws, start_row + 2, hour_cols)
                    total = round(st + ot + dt, 2)
                    if not employee:
                        if total > 0:
                            warnings.append(
                                SheetParseWarning(
                                    source_file=source_path.name,
                                    source_sheet=sheet_name,
                                    work_date=sheet_date.isoformat(),
                                    issue="Hours Without Name",
                                    details=f"{label} block at row {start_row} has {fmt_hours(total)} hours but no employee name.",
                                )
                            )
                        continue

                    seen_names[employee] = seen_names.get(employee, 0) + 1
                    if total > 24:
                        warnings.append(
                            SheetParseWarning(
                                source_file=source_path.name,
                                source_sheet=sheet_name,
                                work_date=sheet_date.isoformat(),
                                issue="Suspicious Total",
                                details=(
                                    f"{employee} in {label} block at row {start_row} totals {fmt_hours(total)} hours "
                                    f"(name cell {cell_ref})."
                                ),
                            )
                        )
                    records.append(
                        DailyRecord(
                            source_path=source_path,
                            source_file=source_path.name,
                            work_date=sheet_date,
                            source_sheet=sheet_name,
                            employee=employee,
                            st=st,
                            ot=ot,
                            dt=dt,
                            source_ranges=build_source_ranges(label, start_row, tuple(name_cols), tuple(hour_cols)),
                        )
                    )

            for employee, count in sorted(seen_names.items()):
                if count > 1:
                    warnings.append(
                        SheetParseWarning(
                            source_file=source_path.name,
                            source_sheet=sheet_name,
                            work_date=sheet_date.isoformat(),
                            issue="Duplicate Employee",
                            details=f"{employee} appears {count} times on the same sheet.",
                        )
                    )
            emit_progress(
                0.2 + (0.75 * sheet_index / total_sheets),
                f"Processed {sheet_name}",
            )
            if (
                preview_callback is not None
                and not preview_emitted
                and preview_sheet_count > 0
                and sheet_index >= preview_sheet_count
            ):
                preview_callback(
                    list(records),
                    list(warnings),
                    list(health),
                    "Showing the most recent one to two weeks while older DSS sheets continue loading.",
                )
                preview_emitted = True
    finally:
        workbook.close()

    if not health:
        health.append(
            WorkbookHealthItem(
                source_file=source_path.name,
                status="OK",
                details="No workbook health issues detected from the current workbook structure.",
            )
        )
    emit_progress(1.0, f"Finished {source_path.name}")
    return records, warnings, health


def analyze_workbook_bytes(
    source_path: Path,
    workbook_bytes: bytes,
    progress_callback: Callable[[float, str], None] | None = None,
) -> tuple[list[SheetParseWarning], list[WorkbookHealthItem]]:
    _records, warnings, health = process_workbook_bytes(
        source_path,
        workbook_bytes,
        progress_callback=progress_callback,
    )
    return warnings, health


def read_source_bytes(source_path: Path) -> bytes:
    try:
        return source_path.read_bytes()
    except PermissionError as exc:
        workbook_bytes = load_open_excel_workbook_copy_bytes(source_path)
        if workbook_bytes is not None:
            return workbook_bytes
        raise PermissionError(build_permission_denied_message(source_path)) from exc
    except FileNotFoundError as exc:
        raise FileNotFoundError(f"Workbook not found:\n{source_path}") from exc

def load_source_workbook(source_path: Path):
    return load_workbook(io.BytesIO(read_source_bytes(source_path)), data_only=True, read_only=True)


def compute_bytes_hash(workbook_bytes: bytes, chunk_size: int = 1024 * 1024) -> str:
    digest = hashlib.blake2b(digest_size=16)
    for offset in range(0, len(workbook_bytes), chunk_size):
        digest.update(workbook_bytes[offset:offset + chunk_size])
    return digest.hexdigest()


def normalize_workbook_member_bytes(member_name: str, member_bytes: bytes) -> bytes:
    normalized_name = member_name.replace("\\", "/")
    if normalized_name == "xl/workbook.xml":
        return WORKBOOK_CALC_ID_PATTERN.sub(b"", member_bytes)
    return member_bytes


def compute_workbook_content_hash(workbook_bytes: bytes) -> str:
    try:
        with zipfile.ZipFile(io.BytesIO(workbook_bytes)) as archive:
            digest = hashlib.blake2b(digest_size=16)
            member_names = sorted(
                name
                for name in archive.namelist()
                if not name.endswith("/")
                and name not in {"docProps/core.xml", "docProps/app.xml", "docProps/custom.xml", "xl/calcChain.xml"}
            )
            if not member_names:
                return compute_bytes_hash(workbook_bytes)
            for member_name in member_names:
                digest.update(member_name.encode("utf-8"))
                digest.update(b"\0")
                digest.update(normalize_workbook_member_bytes(member_name, archive.read(member_name)))
                digest.update(b"\0")
            return digest.hexdigest()
    except (OSError, ValueError, zipfile.BadZipFile, KeyError):
        return compute_bytes_hash(workbook_bytes)


def load_open_excel_workbook_copy_bytes(source_path: Path) -> bytes | None:
    if pythoncom is None or win32com is None:
        return None

    temp_path: str | None = None
    pythoncom.CoInitialize()
    try:
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            return None

        workbook = find_open_excel_workbook(excel, source_path)
        if workbook is None:
            return None

        fd, temp_path = tempfile.mkstemp(prefix="dss_hours_tracker_", suffix=".xlsx")
        os.close(fd)
        workbook.SaveCopyAs(temp_path)
        return Path(temp_path).read_bytes()
    finally:
        if temp_path:
            try:
                Path(temp_path).unlink(missing_ok=True)
            except OSError:
                pass
        pythoncom.CoUninitialize()


def normalize_windows_path(value: str) -> str:
    return value.replace("/", "\\").strip().rstrip("\\").casefold()


def find_open_excel_workbook(excel_app, source_path: Path):
    target = normalize_windows_path(str(source_path))
    target_name = source_path.name.casefold()
    name_matches = []

    for workbook in excel_app.Workbooks:
        full_name = normalize_windows_path(str(workbook.FullName))
        if full_name == target:
            return workbook
        if Path(str(workbook.Name)).name.casefold() == target_name:
            name_matches.append(workbook)

    if len(name_matches) == 1:
        return name_matches[0]
    return None


def build_permission_denied_message(source_path: Path) -> str:
    lower_path = str(source_path).lower()
    if "onedrive" in lower_path or "sharepoint" in lower_path:
        return (
            "Windows denied access to this OneDrive / SharePoint workbook.\n\n"
            f"Path:\n{source_path}\n\n"
            "Things to try:\n"
            "- Keep the workbook open in desktop Excel, then press 'Update View' again.\n"
            "- In File Explorer, right-click the file and choose 'Always keep on this device'.\n"
            "- Wait for OneDrive to finish syncing, then try again.\n"
            "- If Excel just saved the workbook, wait a moment for OneDrive to release the file lock.\n"
            "- If needed, copy the workbook to a normal local folder and open that copy."
        )
    return (
        "Windows denied access to this workbook.\n\n"
        f"Path:\n{source_path}\n\n"
        "Make sure the file is not open in another app and that you have local read access."
    )


def parse_daily_records(workbook_path: Path) -> list[DailyRecord]:
    return parse_daily_records_from_bytes(workbook_path, read_source_bytes(workbook_path))


def parse_daily_records_from_bytes(
    workbook_path: Path,
    workbook_bytes: bytes,
    progress_callback: Callable[[float, str], None] | None = None,
    should_cancel: Callable[[], bool] | None = None,
) -> list[DailyRecord]:
    records, _warnings, _health = process_workbook_bytes(
        workbook_path,
        workbook_bytes,
        progress_callback=progress_callback,
        should_cancel=should_cancel,
    )
    return records


def monday_week_start(value: date) -> date:
    return value - timedelta(days=value.weekday())


def aggregate_weekly(records: Iterable[DailyRecord], combine_sources: bool = False) -> list[WeeklyRecord]:
    grouped: dict[tuple[str, date, str], dict[str, float]] = {}
    for record in records:
        week_start = monday_week_start(record.work_date)
        source_file = "All DSSs" if combine_sources else record.source_file
        bucket = grouped.setdefault((source_file, week_start, record.employee), {"ST": 0.0, "OT": 0.0, "DT": 0.0})
        bucket["ST"] += record.st
        bucket["OT"] += record.ot
        bucket["DT"] += record.dt

    results = []
    for (source_file, week_start, employee), totals in sorted(grouped.items(), key=lambda item: (item[0][1], item[0][0], item[0][2])):
        results.append(
            WeeklyRecord(
                source_file=source_file,
                week_start=week_start,
                week_end=week_start + timedelta(days=6),
                employee=employee,
                st=round(totals["ST"], 2),
                ot=round(totals["OT"], 2),
                dt=round(totals["DT"], 2),
            )
        )
    return results


def build_weekly_rollup(weekly_records: list[WeeklyRecord]) -> list[WeeklyRollupRow]:
    grouped: dict[tuple[str, date], list[WeeklyRecord]] = {}
    for record in weekly_records:
        grouped.setdefault((record.source_file, record.week_start), []).append(record)

    rows: list[WeeklyRollupRow] = []
    for source_file, week_start in sorted(grouped):
        week_entries = sorted(grouped[(source_file, week_start)], key=lambda item: item.employee)
        week_end = week_start + timedelta(days=6)
        crew_st = 0.0
        crew_ot = 0.0
        crew_dt = 0.0
        for record in week_entries:
            rows.append(
                WeeklyRollupRow(
                    source_file=source_file,
                    week_start=week_start,
                    week_end=week_end,
                    employee=record.employee,
                    st=record.st,
                    ot=record.ot,
                    dt=record.dt,
                    row_type="Employee",
                )
            )
            crew_st += record.st
            crew_ot += record.ot
            crew_dt += record.dt
        rows.append(
            WeeklyRollupRow(
                source_file=source_file,
                week_start=week_start,
                week_end=week_end,
                employee="Whole Crew",
                st=round(crew_st, 2),
                ot=round(crew_ot, 2),
                dt=round(crew_dt, 2),
                row_type="Crew Total",
            )
        )
    return rows


def build_week_totals(weekly_records: list[WeeklyRecord]) -> list[WeekTotalRow]:
    grouped: dict[date, dict[str, float]] = {}
    for record in weekly_records:
        bucket = grouped.setdefault(record.week_start, {"ST": 0.0, "OT": 0.0, "DT": 0.0})
        bucket["ST"] += record.st
        bucket["OT"] += record.ot
        bucket["DT"] += record.dt

    rows: list[WeekTotalRow] = []
    for week_start in sorted(grouped):
        totals = grouped[week_start]
        rows.append(
            WeekTotalRow(
                week_start=week_start,
                week_end=week_start + timedelta(days=6),
                st=round(totals["ST"], 2),
                ot=round(totals["OT"], 2),
                dt=round(totals["DT"], 2),
            )
        )
    return rows


def build_tracker_data(source_paths: list[Path], file_hashes: dict[Path, str], daily_records: list[DailyRecord]) -> TrackerData:
    return build_tracker_data_with_status(source_paths, file_hashes, [], source_paths, daily_records)


def build_tracker_data_with_status(
    source_paths: list[Path],
    file_hashes: dict[Path, str],
    reused_paths: list[Path],
    reloaded_paths: list[Path],
    daily_records: list[DailyRecord],
    cache_status_by_path: dict[Path, str] | None = None,
    parse_warnings: list[SheetParseWarning] | None = None,
    workbook_health: list[WorkbookHealthItem] | None = None,
) -> TrackerData:
    if not daily_records:
        raise ValueError("No daily DSS records were found in sheets named with YYYY-MM-DD.")
    weekly_summary = aggregate_weekly(daily_records, combine_sources=False)
    combined_weekly_summary = aggregate_weekly(daily_records, combine_sources=True)
    return TrackerData(
        source_paths=source_paths,
        file_hashes=file_hashes,
        reused_paths=reused_paths,
        reloaded_paths=reloaded_paths,
        cache_status_by_path=cache_status_by_path or {},
        daily_records=sorted(daily_records, key=lambda item: (item.work_date, item.source_sheet, item.employee)),
        employee_names=sorted({record.employee for record in daily_records}),
        weekly_summary=weekly_summary,
        weekly_rollup=build_weekly_rollup(weekly_summary),
        week_totals=build_week_totals(combined_weekly_summary),
        combined_weekly_summary=combined_weekly_summary,
        parse_warnings=parse_warnings or [],
        workbook_health=workbook_health or [],
    )


def load_tracker_data(
    source_paths: Path | Iterable[Path],
    previous_data: TrackerData | None = None,
    progress_callback: Callable[[float, str], None] | None = None,
    partial_callback: Callable[[TrackerData, str], None] | None = None,
    cache_dir: Path | None = None,
    should_cancel: Callable[[], bool] | None = None,
) -> TrackerData:
    if isinstance(source_paths, Path):
        normalized_paths = [source_paths]
    else:
        normalized_paths = sorted({Path(path) for path in source_paths}, key=lambda item: item.name.lower())

    if cache_dir is None:
        _app_root, cache_dir = ensure_app_directories()
    purge_stale_cache(cache_dir)

    previous_hashes = previous_data.file_hashes if previous_data else {}
    previous_records_by_path: dict[Path, list[DailyRecord]] = {}
    previous_warnings_by_source_file: dict[str, list[SheetParseWarning]] = {}
    previous_health_by_source_file: dict[str, list[WorkbookHealthItem]] = {}
    if previous_data is not None:
        for source_path in previous_data.source_paths:
            previous_records_by_path[source_path] = [
                record for record in previous_data.daily_records if record.source_path == source_path
            ]
            previous_warnings_by_source_file[source_path.name] = [
                warning for warning in previous_data.parse_warnings if warning.source_file == source_path.name
            ]
            previous_health_by_source_file[source_path.name] = [
                item for item in previous_data.workbook_health if item.source_file == source_path.name
            ]

    file_hashes: dict[Path, str] = {}
    reused_paths: list[Path] = []
    reloaded_paths: list[Path] = []
    cache_status_by_path: dict[Path, str] = {}
    daily_records: list[DailyRecord] = []
    parse_warnings: list[SheetParseWarning] = []
    workbook_health: list[WorkbookHealthItem] = []
    total_files = max(len(normalized_paths), 1)
    check_cancel = should_cancel or (lambda: False)
    state_lock = threading.Lock()
    file_progress = {source_path: 0.0 for source_path in normalized_paths}

    def raise_if_cancelled(message: str) -> None:
        if check_cancel():
            raise OperationCancelled(message)

    def emit_overall_progress(source_path: Path, file_fraction: float, message: str) -> None:
        if progress_callback is None:
            return
        clamped_fraction = min(max(file_fraction, 0.0), 1.0)
        with state_lock:
            file_progress[source_path] = clamped_fraction
            overall = sum(file_progress.values()) / total_files
        progress_callback(overall, message)

    def emit_partial_data(
        source_path: Path,
        partial_records_for_file: list[DailyRecord],
        partial_warnings_for_file: list[SheetParseWarning],
        partial_health_for_file: list[WorkbookHealthItem],
        message: str,
    ) -> None:
        if partial_callback is None:
            return
        with state_lock:
            snapshot_records = [*daily_records, *partial_records_for_file]
            if not snapshot_records:
                return
            snapshot_status_by_path = dict(cache_status_by_path)
            snapshot_status_by_path[source_path] = snapshot_status_by_path.get(source_path, "Miss")
            tracker_snapshot = build_tracker_data_with_status(
                normalized_paths,
                dict(file_hashes),
                list(reused_paths),
                list(reloaded_paths),
                snapshot_records,
                cache_status_by_path=snapshot_status_by_path,
                parse_warnings=[*parse_warnings, *partial_warnings_for_file],
                workbook_health=[*workbook_health, *partial_health_for_file],
            )
        partial_callback(tracker_snapshot, message)

    def parse_miss_file(
        source_path: Path,
        workbook_bytes: bytes,
    ) -> tuple[Path, list[DailyRecord], list[SheetParseWarning], list[WorkbookHealthItem]]:
        parsed_records, file_warnings, file_health = process_workbook_bytes(
            source_path,
            workbook_bytes,
            progress_callback=lambda fraction, message, current_path=source_path: emit_overall_progress(
                current_path,
                0.2 + (0.8 * min(max(fraction, 0.0), 1.0)),
                message,
            ),
            should_cancel=check_cancel,
            preview_callback=lambda partial_records, partial_warnings, partial_health, message, current_path=source_path: emit_partial_data(
                current_path,
                partial_records,
                partial_warnings,
                partial_health,
                message,
            ),
        )
        return source_path, parsed_records, file_warnings, file_health

    pending_parse_inputs: list[tuple[Path, bytes]] = []

    for source_path in normalized_paths:
        raise_if_cancelled(f"Cancelled loading {source_path.name}")
        emit_overall_progress(source_path, 0.0, f"Starting {source_path.name}")
        workbook_bytes = read_source_bytes(source_path)
        emit_overall_progress(source_path, 0.05, f"Read bytes for {source_path.name}")
        raise_if_cancelled(f"Cancelled loading {source_path.name}")
        file_hash = compute_workbook_content_hash(workbook_bytes)
        file_hashes[source_path] = file_hash
        emit_overall_progress(source_path, 0.1, f"Hashed {source_path.name}")
        if previous_hashes.get(source_path) == file_hash:
            with state_lock:
                reused_paths.append(source_path)
                cache_status_by_path[source_path] = "Memory Hit"
                daily_records.extend(previous_records_by_path.get(source_path, []))
                parse_warnings.extend(previous_warnings_by_source_file.get(source_path.name, []))
                workbook_health.extend(previous_health_by_source_file.get(source_path.name, []))
            emit_overall_progress(source_path, 1.0, f"Unchanged {source_path.name}")
        else:
            raise_if_cancelled(f"Cancelled loading {source_path.name}")
            cached_records = load_cached_daily_records(cache_dir, source_path, file_hash)
            if cached_records is not None:
                with state_lock:
                    reused_paths.append(source_path)
                    cache_status_by_path[source_path] = "Disk Hit"
                    daily_records.extend(cached_records)
                cached_analysis = load_cached_source_analysis(cache_dir, source_path, file_hash)
                if cached_analysis is not None:
                    cached_warnings, cached_health = cached_analysis
                    with state_lock:
                        parse_warnings.extend(cached_warnings)
                        workbook_health.extend(cached_health)
                else:
                    emit_overall_progress(source_path, 0.2, f"Inspecting {source_path.name}")
                    file_warnings, file_health = analyze_workbook_bytes(
                        source_path,
                        workbook_bytes,
                        progress_callback=lambda fraction, message, current_path=source_path: emit_overall_progress(
                            current_path,
                            0.2 + (0.1 * min(max(fraction, 0.0), 1.0)),
                            message,
                        ),
                    )
                    with state_lock:
                        parse_warnings.extend(file_warnings)
                        workbook_health.extend(file_health)
                    save_cached_daily_records(
                        cache_dir,
                        source_path,
                        file_hash,
                        cached_records,
                        parse_warnings=file_warnings,
                        workbook_health=file_health,
                    )
                emit_overall_progress(source_path, 1.0, f"Loaded cached data for {source_path.name}")
            else:
                with state_lock:
                    reloaded_paths.append(source_path)
                    cache_status_by_path[source_path] = "Miss"
                emit_overall_progress(source_path, 0.15, f"Queued {source_path.name} for parsing")
                pending_parse_inputs.append((source_path, workbook_bytes))

    if pending_parse_inputs:
        max_workers = min(MAX_PARALLEL_PARSE_WORKERS, len(pending_parse_inputs))
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_map = {
                executor.submit(parse_miss_file, source_path, workbook_bytes): source_path
                for source_path, workbook_bytes in pending_parse_inputs
            }
            for future in concurrent.futures.as_completed(future_map):
                raise_if_cancelled("Cancelled DSS load.")
                source_path = future_map[future]
                parsed_source_path, parsed_records, file_warnings, file_health = future.result()
                save_cached_daily_records(
                    cache_dir,
                    parsed_source_path,
                    file_hashes[parsed_source_path],
                    parsed_records,
                    parse_warnings=file_warnings,
                    workbook_health=file_health,
                )
                with state_lock:
                    parse_warnings.extend(file_warnings)
                    workbook_health.extend(file_health)
                    daily_records.extend(parsed_records)
                emit_overall_progress(source_path, 1.0, f"Finished {source_path.name}")
    return build_tracker_data_with_status(
        normalized_paths,
        file_hashes,
        reused_paths,
        reloaded_paths,
        daily_records,
        cache_status_by_path=cache_status_by_path,
        parse_warnings=parse_warnings,
        workbook_health=workbook_health,
    )


def fmt_hours(value: float) -> str:
    text = f"{value:.2f}"
    return text.rstrip("0").rstrip(".")


def expanded_hours(st: float, ot: float, dt: float) -> float:
    return round(st + (ot * 1.5) + (dt * 2.0), 2)


class SortableTreeview(ttk.Treeview):
    def __init__(self, master, columns: list[str], headings: list[str]):
        super().__init__(master, columns=columns, show="headings")
        self._rows: list[tuple[tuple[str, ...], tuple[str, ...]]] = []
        self._sort_state: dict[str, bool] = {}
        self._active_sort_column: str | None = None
        self._active_sort_descending = False
        self.on_sort_changed: Callable[[], None] | None = None
        for column, heading in zip(columns, headings):
            self.heading(column, text=heading, command=lambda col=column: self.sort_by(col))
            self.column(column, anchor="w", width=120, stretch=True)

    def set_rows(self, rows: list[tuple[str, ...]], tags: list[tuple[str, ...]] | None = None) -> None:
        self._rows = []
        tags = tags or [tuple() for _ in rows]
        for row, row_tags in zip(rows, tags):
            self._rows.append((row, row_tags))
        if self._active_sort_column and self._active_sort_column in self["columns"]:
            self._sort_rows(self._active_sort_column, self._active_sort_descending)
        self._render_rows()

    def sort_by(self, column: str, descending: bool | None = None) -> None:
        columns = list(self["columns"])
        index = columns.index(column)
        if descending is None:
            if self._active_sort_column == column:
                reverse = not self._active_sort_descending
            else:
                reverse = False
        else:
            reverse = descending

        self._sort_rows(column, reverse, index=index)
        self._active_sort_column = column
        self._active_sort_descending = reverse
        self._sort_state[column] = not reverse
        self._render_rows()
        if self.on_sort_changed is not None:
            self.on_sort_changed()

    def set_sort(self, column: str, descending: bool) -> None:
        if column not in self["columns"]:
            return
        self._active_sort_column = column
        self._active_sort_descending = descending
        self._sort_state[column] = not descending
        if self._rows:
            self._sort_rows(column, descending)
            self._render_rows()

    def current_sort(self) -> tuple[str, bool] | None:
        if not self._active_sort_column:
            return None
        return self._active_sort_column, self._active_sort_descending

    def clear_sort(self) -> None:
        self._active_sort_column = None
        self._active_sort_descending = False
        self._sort_state.clear()

    def _sort_rows(self, column: str, reverse: bool, index: int | None = None) -> None:
        columns = list(self["columns"])
        column_index = columns.index(column) if index is None else index

        def sort_key(item: tuple[tuple[str, ...], tuple[str, ...]]):
            value = item[0][column_index]
            try:
                return (0, float(value))
            except ValueError:
                return (1, value.lower())

        self._rows.sort(key=sort_key, reverse=reverse)

    def _render_rows(self) -> None:
        self.delete(*self.get_children())
        for row, row_tags in self._rows:
            self.insert("", "end", values=row, tags=row_tags)


class DataTable(ttk.Frame):
    def __init__(
        self,
        master,
        columns: list[str],
        headings: list[str],
        table_id: str | None = None,
        config_path: Path | None = None,
        default_sort_column: str | None = None,
        default_sort_descending: bool = False,
    ):
        super().__init__(master)
        self._table_id = table_id
        self._config_path = config_path
        self._all_columns = list(columns)
        self._headings_by_column = dict(zip(columns, headings))
        self._column_visibility = {column: True for column in columns}
        self._default_sort_column = default_sort_column if default_sort_column in self._all_columns else None
        self._default_sort_descending = default_sort_descending

        toolbar = ttk.Frame(self)
        toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Button(toolbar, text="Columns", command=self.open_column_picker).pack(side="right")

        self.tree = SortableTreeview(self, columns, headings)
        y_scroll = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.tree.grid(row=1, column=0, sticky="nsew")
        y_scroll.grid(row=1, column=1, sticky="ns")
        x_scroll.grid(row=2, column=0, sticky="ew")
        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)

        self.tree.tag_configure("alert", background="#ffc7ce", foreground="#9c0006")
        self.tree.tag_configure("crew_total", background="#e2f0d9", foreground="#1f1f1f")
        self.tree.bind("<ButtonRelease-1>", self._on_tree_button_release, add="+")
        self.tree.on_sort_changed = self._save_layout
        bind_vertical_mousewheel(self.tree, self.tree, units_per_notch=4)
        bind_horizontal_mousewheel(self.tree, self.tree, units_per_notch=4)

        self._load_saved_layout()

    def set_rows(self, rows: list[tuple[str, ...]], tags: list[tuple[str, ...]] | None = None) -> None:
        self.tree.set_rows(rows, tags)

    def open_column_picker(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("Show / Hide Columns")
        dialog.transient(self.winfo_toplevel())
        dialog.resizable(False, False)

        vars_by_column: dict[str, tk.BooleanVar] = {}
        content = ttk.Frame(dialog, padding=12)
        content.pack(fill="both", expand=True)

        for row_index, column in enumerate(self._all_columns):
            var = tk.BooleanVar(value=self._column_visibility[column])
            vars_by_column[column] = var
            ttk.Checkbutton(
                content,
                text=self._headings_by_column[column],
                variable=var,
            ).grid(row=row_index, column=0, sticky="w")

        buttons = ttk.Frame(content)
        buttons.grid(row=len(self._all_columns), column=0, sticky="ew", pady=(12, 0))
        ttk.Button(buttons, text="Show All", command=lambda: self._set_all_column_vars(vars_by_column, True)).pack(side="left")
        ttk.Button(buttons, text="Hide All", command=lambda: self._set_all_column_vars(vars_by_column, False)).pack(side="left", padx=(8, 0))
        ttk.Button(
            buttons,
            text="Apply",
            command=lambda: self._apply_column_visibility(vars_by_column, dialog),
        ).pack(side="right")

    @staticmethod
    def _set_all_column_vars(vars_by_column: dict[str, tk.BooleanVar], value: bool) -> None:
        for var in vars_by_column.values():
            var.set(value)

    def _apply_column_visibility(self, vars_by_column: dict[str, tk.BooleanVar], dialog: tk.Toplevel) -> None:
        selected_columns = [column for column in self._all_columns if vars_by_column[column].get()]
        if not selected_columns:
            messagebox.showerror("Columns", "At least one column must remain visible.")
            return
        for column in self._all_columns:
            self._column_visibility[column] = vars_by_column[column].get()
        self.tree.configure(displaycolumns=selected_columns)
        self._save_layout()
        dialog.destroy()

    def _load_saved_layout(self) -> None:
        if not self._config_path or not self._table_id:
            self._apply_default_sort()
            return
        layouts = load_table_layouts(self._config_path)
        layout = layouts.get(self._table_id)
        if not isinstance(layout, dict):
            self._apply_default_sort()
            return

        requested_columns = [
            column for column in layout.get("visible_columns", [])
            if column in self._all_columns
        ]
        if requested_columns:
            self.tree.configure(displaycolumns=requested_columns)
            for column in self._all_columns:
                self._column_visibility[column] = column in requested_columns

        for column, width in layout.get("column_widths", {}).items():
            if column in self._all_columns and width > 0:
                self.tree.column(column, width=width)
        sort_column = str(layout.get("sort_column", "")).strip()
        if sort_column in self._all_columns:
            self.tree.set_sort(sort_column, bool(layout.get("sort_descending", False)))
        else:
            self._apply_default_sort()

    def _apply_default_sort(self) -> None:
        if self._default_sort_column:
            self.tree.set_sort(self._default_sort_column, self._default_sort_descending)

    def _save_layout(self) -> None:
        if not self._config_path or not self._table_id:
            return
        display_columns_value = self.tree.cget("displaycolumns")
        if display_columns_value == "#all":
            visible_columns = list(self._all_columns)
        else:
            display_columns = list(display_columns_value)
            visible_columns = [column for column in display_columns if column in self._all_columns]
        column_widths = {
            column: int(self.tree.column(column, "width"))
            for column in self._all_columns
        }
        current_sort = self.tree.current_sort()
        sort_column = current_sort[0] if current_sort is not None else ""
        sort_descending = current_sort[1] if current_sort is not None else False
        save_table_layout(
            self._config_path,
            self._table_id,
            visible_columns,
            column_widths,
            sort_column=sort_column,
            sort_descending=sort_descending,
        )

    def _on_tree_button_release(self, _event=None) -> None:
        self.after_idle(self._save_layout)

    def reset_layout(self) -> None:
        for column in self._all_columns:
            self._column_visibility[column] = True
            self.tree.column(column, width=120)
        self.tree.configure(displaycolumns="#all")
        self.tree.clear_sort()
        self._apply_default_sort()
        self.tree._render_rows()

    def displayed_columns_and_rows(self) -> tuple[list[str], list[tuple[str, ...]]]:
        display_columns_value = self.tree.cget("displaycolumns")
        if display_columns_value == "#all":
            columns = list(self._all_columns)
        else:
            columns = [column for column in list(display_columns_value) if column in self._all_columns]
        headings = [self._headings_by_column[column] for column in columns]
        rows: list[tuple[str, ...]] = []
        for item_id in self.tree.get_children():
            values = self.tree.item(item_id, "values")
            rows.append(tuple(str(value) for value in values))
        return headings, rows


class ToolTip:
    def __init__(self, widget: tk.Widget, text: str):
        self.widget = widget
        self.text = text
        self.tip_window: tk.Toplevel | None = None
        self._after_id: str | None = None
        existing = getattr(widget, "_tooltip_refs", [])
        existing.append(self)
        setattr(widget, "_tooltip_refs", existing)
        widget.bind("<Enter>", self._schedule_show, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")
        widget.bind("<Destroy>", self._hide, add="+")

    def _schedule_show(self, _event=None) -> None:
        self._cancel_scheduled()
        self._after_id = self.widget.after(500, self._show)

    def _cancel_scheduled(self) -> None:
        if self._after_id is not None:
            self.widget.after_cancel(self._after_id)
            self._after_id = None

    def _show(self) -> None:
        self._after_id = None
        if self.tip_window is not None or not self.text.strip():
            return
        x = self.widget.winfo_rootx() + 18
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        self.tip_window = tk.Toplevel(self.widget)
        self.tip_window.wm_overrideredirect(True)
        self.tip_window.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            self.tip_window,
            text=self.text,
            justify="left",
            background="#fff8dc",
            relief="solid",
            borderwidth=1,
            wraplength=320,
            padx=8,
            pady=6,
        )
        label.pack()

    def _hide(self, _event=None) -> None:
        self._cancel_scheduled()
        if self.tip_window is not None:
            self.tip_window.destroy()
            self.tip_window = None


def bind_vertical_mousewheel(widget: tk.Widget, scroll_target, units_per_notch: int = 3) -> None:
    def on_mousewheel(event) -> str:
        delta = int(getattr(event, "delta", 0))
        if delta == 0:
            return "break"
        steps = max(1, abs(delta) // 120) * units_per_notch
        direction = -1 if delta > 0 else 1
        scroll_target.yview_scroll(direction * steps, "units")
        return "break"

    def on_button4(_event) -> str:
        scroll_target.yview_scroll(-units_per_notch, "units")
        return "break"

    def on_button5(_event) -> str:
        scroll_target.yview_scroll(units_per_notch, "units")
        return "break"

    widget.bind("<MouseWheel>", on_mousewheel, add="+")
    widget.bind("<Button-4>", on_button4, add="+")
    widget.bind("<Button-5>", on_button5, add="+")


def bind_horizontal_mousewheel(widget: tk.Widget, scroll_target, units_per_notch: int = 3) -> None:
    def on_shift_mousewheel(event) -> str:
        delta = int(getattr(event, "delta", 0))
        if delta == 0:
            return "break"
        steps = max(1, abs(delta) // 120) * units_per_notch
        direction = -1 if delta > 0 else 1
        scroll_target.xview_scroll(direction * steps, "units")
        return "break"

    widget.bind("<Shift-MouseWheel>", on_shift_mousewheel, add="+")


class EmployeeListEditor(ttk.Frame):
    def __init__(
        self,
        master,
        on_email_changed: Callable[[str, str], None] | None = None,
        on_request_edit_email: Callable[[str], None] | None = None,
    ):
        super().__init__(master)
        self.names: list[str] = []
        self.email_map: dict[str, str] = {}
        self.on_email_changed = on_email_changed
        self.on_request_edit_email = on_request_edit_email

        top = ttk.Frame(self)
        top.pack(fill="x", pady=(0, 8))

        self.entry = ttk.Entry(top)
        self.entry.pack(side="left", fill="x", expand=True)
        ttk.Button(top, text="Add Name", command=self.add_name).pack(side="left", padx=(8, 0))
        ttk.Button(top, text="Remove Selected", command=self.remove_selected).pack(side="left", padx=(8, 0))

        self.listbox = tk.Listbox(self, activestyle="none")
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        self.listbox.bind("<<ListboxSelect>>", self._on_selection_changed)
        self.listbox.bind("<Double-Button-1>", self._on_double_click)
        bind_vertical_mousewheel(self.listbox, self.listbox, units_per_notch=4)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="left", fill="y")

        editor = ttk.Frame(self)
        editor.pack(fill="x", pady=(8, 0))
        ttk.Label(editor, text="Selected Employee").grid(row=0, column=0, sticky="w")
        self.selected_name_var = tk.StringVar(value="")
        ttk.Label(editor, textvariable=self.selected_name_var).grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Label(editor, text="Email").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.email_var = tk.StringVar(value="")
        ttk.Entry(editor, textvariable=self.email_var).grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))
        ttk.Button(editor, text="Save Email", command=self.save_email).grid(row=1, column=2, padx=(8, 0), pady=(8, 0))
        editor.columnconfigure(1, weight=1)

        note = ttk.Label(
            self,
            text="Parsed employee names are grouped by exact spelling. You can also store employee email addresses here for draft emails.",
            wraplength=500,
            justify="left",
        )
        note.pack(fill="x", pady=(8, 0))

    def set_names(self, names: list[str], email_map: dict[str, str]) -> None:
        self.names = list(names)
        self.email_map = dict(email_map)
        self.refresh()

    def refresh(self) -> None:
        self.listbox.delete(0, tk.END)
        for name in self.names:
            self.listbox.insert(tk.END, name)
        self._on_selection_changed()

    def add_name(self) -> None:
        value = self.entry.get().strip()
        if not value or value in self.names:
            return
        self.names.append(value)
        self.names.sort()
        self.entry.delete(0, tk.END)
        self.refresh()

    def remove_selected(self) -> None:
        selected = self.listbox.curselection()
        if not selected:
            return
        for index in reversed(selected):
            employee = self.names[index]
            self.email_map.pop(employee, None)
            del self.names[index]
        self.refresh()

    def _on_selection_changed(self, _event=None) -> None:
        selected = self.listbox.curselection()
        if not selected:
            self.selected_name_var.set("")
            self.email_var.set("")
            return
        employee = self.names[selected[0]]
        self.selected_name_var.set(employee)
        self.email_var.set(self.email_map.get(employee, ""))

    def save_email(self) -> None:
        employee = self.selected_name_var.get().strip()
        if not employee:
            return
        email = self.email_var.get().strip()
        self.email_map[employee] = email
        if self.on_email_changed is not None:
            self.on_email_changed(employee, email)

    def _on_double_click(self, _event=None) -> None:
        employee = self.selected_name_var.get().strip()
        if employee and self.on_request_edit_email is not None:
            self.on_request_edit_email(employee)


class EmailDraftsFrame(ttk.Frame):
    def __init__(
        self,
        master,
        create_drafts_callback: Callable[[], None],
        sync_emails_callback: Callable[[], None],
        on_request_edit_email: Callable[[str], None],
        save_templates_callback: Callable[[], None],
    ):
        super().__init__(master, padding=12)
        self.create_drafts_callback = create_drafts_callback
        self.sync_emails_callback = sync_emails_callback
        self.on_request_edit_email = on_request_edit_email
        self.save_templates_callback = save_templates_callback

        self.columnconfigure(1, weight=1)
        self.rowconfigure(4, weight=1)

        self.week_options: list[tuple[date, date]] = []
        self.week_var = tk.StringVar()

        ttk.Label(self, text="Week").grid(row=0, column=0, sticky="w")
        self.week_combo = ttk.Combobox(self, textvariable=self.week_var, state="readonly")
        self.week_combo.grid(row=0, column=1, sticky="ew", pady=(0, 8))

        ttk.Label(self, text="Subject Template").grid(row=1, column=0, sticky="nw")
        self.subject_template_text = tk.Text(self, wrap="word", height=2)
        self.subject_template_text.grid(row=1, column=1, sticky="ew", pady=(0, 8))
        bind_vertical_mousewheel(self.subject_template_text, self.subject_template_text, units_per_notch=4)

        ttk.Button(self, text="Create Outlook Drafts", command=self.create_drafts_callback).grid(
            row=1, column=2, padx=(8, 0), pady=(0, 8)
        )
        ttk.Button(self, text="Sync Outlook Emails", command=self.sync_emails_callback).grid(
            row=1, column=3, padx=(8, 0), pady=(0, 8)
        )
        ttk.Button(self, text="Save Templates", command=self.save_templates_callback).grid(
            row=1, column=4, padx=(8, 0), pady=(0, 8)
        )

        self.summary_label = ttk.Label(
            self,
            text="Load DSS data to preview employee email drafts.",
            wraplength=700,
            justify="left",
        )
        self.summary_label.grid(row=2, column=0, columnspan=3, sticky="w", pady=(0, 8))

        ttk.Label(self, text="Body Template (HTML)").grid(row=3, column=0, sticky="nw")
        self.body_template_text = tk.Text(self, wrap="word", height=8)
        self.body_template_text.grid(row=3, column=1, columnspan=4, sticky="nsew", pady=(0, 8))
        bind_vertical_mousewheel(self.body_template_text, self.body_template_text, units_per_notch=4)

        self.preview_table = DataTable(
            self,
            columns=["employee", "email", "days", "st", "ot", "dt", "total", "expanded"],
            headings=["Employee", "Email", "Rows", "ST", "OT", "DT", "Total", "Expanded Hours"],
        )
        self.preview_table.grid(row=4, column=0, columnspan=5, sticky="nsew")
        self.preview_table.tree.bind("<Double-Button-1>", self._on_preview_double_click)

        note = (
            "Use {employee}, {first_name}, {week_start}, {week_end}, {pf_numbers}, and {hours_table} in the templates. "
            "Drafts are saved in Outlook and are not sent automatically."
        )
        ttk.Label(self, text=note, wraplength=700, justify="left").grid(row=5, column=0, columnspan=5, sticky="w", pady=(8, 0))

    def set_week_options(self, week_options: list[tuple[date, date]]) -> None:
        self.week_options = week_options
        labels = [format_week_label(week_start, week_end) for week_start, week_end in week_options]
        self.week_combo.configure(values=labels)
        if labels:
            self.week_var.set(labels[-1])
        else:
            self.week_var.set("")

    def selected_week_start(self) -> date | None:
        selected_label = self.week_var.get().strip()
        for week_start, week_end in self.week_options:
            if format_week_label(week_start, week_end) == selected_label:
                return week_start
        return self.week_options[-1][0] if self.week_options else None

    def _on_preview_double_click(self, _event=None) -> None:
        selected = self.preview_table.tree.selection()
        if not selected:
            return
        employee = self.preview_table.tree.item(selected[0], "values")[0]
        if employee:
            self.on_request_edit_email(employee)

    def set_templates(self, subject_template: str, body_template: str) -> None:
        self.subject_template_text.delete("1.0", tk.END)
        self.subject_template_text.insert("1.0", subject_template)
        self.body_template_text.delete("1.0", tk.END)
        self.body_template_text.insert("1.0", body_template)

    def get_subject_template(self) -> str:
        return self.subject_template_text.get("1.0", tk.END).strip()

    def get_body_template(self) -> str:
        return self.body_template_text.get("1.0", tk.END).strip()

    def selected_employees(self) -> set[str]:
        employees: set[str] = set()
        for item_id in self.preview_table.tree.selection():
            values = self.preview_table.tree.item(item_id, "values")
            if values:
                employees.add(str(values[0]))
        return employees


class EmployeeGroupsFrame(ttk.Frame):
    def __init__(
        self,
        master,
        on_groups_changed: Callable[[dict[str, list[str]]], None],
        on_request_edit_email: Callable[[str], None],
        sync_emails_callback: Callable[[], None],
    ):
        super().__init__(master, padding=12)
        self.on_groups_changed = on_groups_changed
        self.on_request_edit_email = on_request_edit_email
        self.sync_emails_callback = sync_emails_callback
        self.employee_names: list[str] = []
        self.employee_groups: dict[str, list[str]] = {}
        self.email_map: dict[str, str] = {}

        self.columnconfigure(1, weight=1)
        self.rowconfigure(1, weight=1)

        left = ttk.Frame(self)
        left.grid(row=0, column=0, rowspan=3, sticky="nsw", padx=(0, 12))
        ttk.Label(left, text="Groups").pack(anchor="w")
        self.group_listbox = tk.Listbox(left, exportselection=False, height=12)
        self.group_listbox.pack(fill="both", expand=True, pady=(4, 8))
        self.group_listbox.bind("<<ListboxSelect>>", self._on_group_selected)
        bind_vertical_mousewheel(self.group_listbox, self.group_listbox, units_per_notch=4)
        ttk.Button(left, text="New Group", command=self.create_group).pack(fill="x")
        ttk.Button(left, text="Delete Group", command=self.delete_group).pack(fill="x", pady=(8, 0))
        ttk.Button(left, text="Sync Outlook Emails", command=self.sync_emails_callback).pack(fill="x", pady=(8, 0))

        right = ttk.Frame(self)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        ttk.Label(right, text="Employees in selected group").grid(row=0, column=0, sticky="w")
        self.employee_listbox = tk.Listbox(right, selectmode="extended", exportselection=False)
        self.employee_listbox.grid(row=1, column=0, sticky="nsew", pady=(4, 8))
        self.employee_listbox.bind("<Double-Button-1>", self._on_employee_double_click)
        bind_vertical_mousewheel(self.employee_listbox, self.employee_listbox, units_per_notch=4)
        employee_scroll = ttk.Scrollbar(right, orient="vertical", command=self.employee_listbox.yview)
        self.employee_listbox.configure(yscrollcommand=employee_scroll.set)
        employee_scroll.grid(row=1, column=1, sticky="ns", pady=(4, 8))

        ttk.Button(right, text="Save Group Members", command=self.save_group_members).grid(row=2, column=0, sticky="w")
        self.summary_label = ttk.Label(
            self,
            text="Create groups to filter the tables by crews, trades, or project splits.",
            wraplength=700,
            justify="left",
        )
        self.summary_label.grid(row=3, column=0, columnspan=2, sticky="w", pady=(12, 0))

    def set_data(
        self,
        employee_names: list[str],
        employee_groups: dict[str, list[str]],
        email_map: dict[str, str],
    ) -> None:
        current_group = self.selected_group_name()
        self.employee_names = list(employee_names)
        self.employee_groups = {name: list(members) for name, members in employee_groups.items()}
        self.email_map = dict(email_map)

        self.group_listbox.delete(0, tk.END)
        for group_name in sorted(self.employee_groups):
            self.group_listbox.insert(tk.END, group_name)

        self.employee_listbox.delete(0, tk.END)
        for employee in self.employee_names:
            email = self.email_map.get(employee, "").strip()
            label = f"{employee} | {email}" if email else f"{employee} | (missing)"
            self.employee_listbox.insert(tk.END, label)

        if current_group and current_group in self.employee_groups:
            group_names = sorted(self.employee_groups)
            self.group_listbox.selection_set(group_names.index(current_group))
        elif self.group_listbox.size() > 0:
            self.group_listbox.selection_set(0)
        self._on_group_selected()

    def selected_group_name(self) -> str | None:
        selection = self.group_listbox.curselection()
        if not selection:
            return None
        return self.group_listbox.get(selection[0])

    def create_group(self) -> None:
        name = simpledialog.askstring("New Employee Group", "Group name:", parent=self)
        if not name:
            return
        name = name.strip()
        if not name:
            return
        if name in self.employee_groups:
            messagebox.showerror("Employee Groups", "A group with that name already exists.")
            return
        self.employee_groups[name] = []
        self._notify_change()
        self.set_data(self.employee_names, self.employee_groups)

    def delete_group(self) -> None:
        group_name = self.selected_group_name()
        if not group_name:
            return
        if not messagebox.askyesno("Delete Group", f"Delete group '{group_name}'?"):
            return
        del self.employee_groups[group_name]
        self._notify_change()
        self.set_data(self.employee_names, self.employee_groups)

    def save_group_members(self) -> None:
        group_name = self.selected_group_name()
        if not group_name:
            return
        selected_indices = self.employee_listbox.curselection()
        self.employee_groups[group_name] = [self.employee_names[index] for index in selected_indices]
        self._notify_change()
        self.summary_label.configure(
            text=f"Saved {len(self.employee_groups[group_name])} employee(s) in group '{group_name}'."
        )

    def _on_group_selected(self, _event=None) -> None:
        for index in range(self.employee_listbox.size()):
            self.employee_listbox.selection_clear(index)
        group_name = self.selected_group_name()
        if not group_name:
            return
        members = set(self.employee_groups.get(group_name, []))
        for index, employee in enumerate(self.employee_names):
            if employee in members:
                self.employee_listbox.selection_set(index)
        self.summary_label.configure(
            text=f"Group '{group_name}' contains {len(members)} employee(s)."
        )

    def _notify_change(self) -> None:
        self.on_groups_changed({name: list(members) for name, members in self.employee_groups.items()})

    def _on_employee_double_click(self, _event=None) -> None:
        selected_indices = self.employee_listbox.curselection()
        if not selected_indices:
            return
        employee = self.employee_names[selected_indices[0]]
        self.on_request_edit_email(employee)


class EmployeeNotesEditor(ttk.Frame):
    def __init__(self, master, on_notes_changed: Callable[[str, str], None]):
        super().__init__(master, padding=12)
        self.on_notes_changed = on_notes_changed
        self.names: list[str] = []
        self.notes_map: dict[str, str] = {}

        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        self.listbox = tk.Listbox(self, exportselection=False)
        self.listbox.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self.listbox.bind("<<ListboxSelect>>", self._on_selection_changed)
        bind_vertical_mousewheel(self.listbox, self.listbox, units_per_notch=4)

        editor = ttk.Frame(self)
        editor.grid(row=0, column=1, sticky="nsew")
        editor.columnconfigure(0, weight=1)
        editor.rowconfigure(1, weight=1)

        self.selected_name_var = tk.StringVar(value="")
        ttk.Label(editor, textvariable=self.selected_name_var).grid(row=0, column=0, sticky="w", pady=(0, 8))
        self.note_text = tk.Text(editor, wrap="word", height=10)
        self.note_text.grid(row=1, column=0, sticky="nsew")
        bind_vertical_mousewheel(self.note_text, self.note_text, units_per_notch=4)
        ttk.Button(editor, text="Save Note", command=self.save_note).grid(row=2, column=0, sticky="w", pady=(8, 0))

    def set_data(self, names: list[str], notes_map: dict[str, str]) -> None:
        current = self.selected_employee()
        self.names = list(names)
        self.notes_map = dict(notes_map)
        self.listbox.delete(0, tk.END)
        for name in self.names:
            note = self.notes_map.get(name, "").strip()
            label = f"{name} *" if note else name
            self.listbox.insert(tk.END, label)
        if current and current in self.names:
            self.listbox.selection_set(self.names.index(current))
        elif self.names:
            self.listbox.selection_set(0)
        self._on_selection_changed()

    def selected_employee(self) -> str | None:
        selected = self.listbox.curselection()
        if not selected:
            return None
        return self.names[selected[0]]

    def _on_selection_changed(self, _event=None) -> None:
        employee = self.selected_employee()
        self.selected_name_var.set(employee or "")
        self.note_text.delete("1.0", tk.END)
        if employee:
            self.note_text.insert("1.0", self.notes_map.get(employee, ""))

    def save_note(self) -> None:
        employee = self.selected_employee()
        if not employee:
            return
        note = self.note_text.get("1.0", tk.END).strip()
        self.notes_map[employee] = note
        self.on_notes_changed(employee, note)
        self.set_data(self.names, self.notes_map)


class DssHoursTrackerApp(tk.Tk):
    def __init__(self, initial_source: Path | Iterable[Path] | None = None):
        super().__init__()
        self.title("DSS Hours Tracker")
        self.geometry("1200x760")
        self.minsize(760, 520)
        self.app_root, self.cache_dir = ensure_app_directories()
        self.config_path = self.app_root / CONFIG_FILENAME
        self.formatting_profiles, self.current_profile_name = load_formatting_profiles(self.config_path)
        self.employee_emails = load_employee_emails(self.config_path)
        self.employee_notes = load_employee_notes(self.config_path)
        self.email_subject_template, self.email_body_template = load_email_templates(self.config_path)
        self.employee_groups = load_employee_groups(self.config_path)
        self.ignored_name_typos = load_ignored_name_typos(self.config_path)
        self.job_presets = load_job_presets(self.config_path)
        self.app_settings = load_app_settings(self.config_path)
        self.filter_button_var = tk.StringVar(value="All Employees")
        self.pf_filter_button_var = tk.StringVar(value="All PFs")
        self._employee_filter_state: dict[str, bool] = {}
        self._employee_filter_vars: dict[str, tk.BooleanVar] = {}
        self._pf_filter_state: dict[str, bool] = {}
        self._pf_filter_vars: dict[str, tk.BooleanVar] = {}
        self.filter_popup: tk.Toplevel | None = None
        self.filter_popup_content: ttk.Frame | None = None
        self.pf_filter_popup: tk.Toplevel | None = None
        self.pf_filter_popup_content: ttk.Frame | None = None
        self.current_data: TrackerData | None = None
        self._has_partial_preview = False
        self._load_request_id = 0
        self._is_loading = False
        self.progress_var = tk.DoubleVar(value=0.0)
        self._hash_monitor_token = 0
        self._hash_alerted_paths: set[Path] = set()
        self._outlook_sync_in_progress = False
        self._outlook_auto_sync_done = False
        self._update_check_in_progress = False
        self.update_status_var = tk.StringVar(value=f"Installed version: {APP_VERSION}")
        self.hash_poll_interval_ms = self.app_settings.hash_poll_minutes * 60 * 1000
        self._cancel_event: threading.Event | None = None
        self._active_operation_name = ""

        self._build_layout()
        self.after(AUTO_OUTLOOK_SYNC_DELAY_MS, self._auto_sync_outlook_emails)
        if initial_source:
            self.load_source(initial_source)

    def _build_layout(self) -> None:
        container = ttk.Frame(self, padding=12)
        container.pack(fill="both", expand=True)

        top = ttk.Frame(container)
        top.pack(fill="x", pady=(0, 12))

        self.open_button = ttk.Button(top, text="Open DSS Workbook(s)", command=self.choose_sources)
        self.open_button.pack(side="left")
        self.add_button = ttk.Button(top, text="Add DSS", command=self.add_sources)
        self.add_button.pack(side="left", padx=(8, 0))
        self.remove_button = ttk.Button(top, text="Remove DSS(s)", command=self.remove_sources, state="disabled")
        self.remove_button.pack(side="left", padx=(8, 0))
        self.reload_button = ttk.Button(top, text="Update View", command=self.reload_source, state="disabled")
        self.reload_button.pack(side="left", padx=(8, 0))
        self.export_button = ttk.Button(top, text="Export Current View", command=self.export_current_view)
        self.export_button.pack(side="left", padx=(8, 0))
        ttk.Label(top, text="Filter").pack(side="left", padx=(12, 4))
        self.filter_button = ttk.Button(top, textvariable=self.filter_button_var, width=28, command=self._toggle_filter_popup)
        self.filter_button.pack(side="left")
        ttk.Label(top, text="PF").pack(side="left", padx=(8, 4))
        self.pf_filter_button = ttk.Button(
            top,
            textvariable=self.pf_filter_button_var,
            width=18,
            command=self._toggle_pf_filter_popup,
            state="disabled",
        )
        self.pf_filter_button.pack(side="left")
        self.source_label = ttk.Label(top, text="No workbook loaded", anchor="w")
        self.source_label.pack(side="left", fill="x", expand=True, padx=(12, 0))
        self.loading_label = ttk.Label(top, text="", anchor="e")
        self.loading_label.pack(side="right")

        stats = ttk.Frame(container)
        stats.pack(fill="x", pady=(0, 12))
        self.stats_label = ttk.Label(stats, text="Load a DSS workbook to view daily and weekly labour summaries.")
        self.stats_label.pack(side="left", fill="x", expand=True)
        self.cancel_button = ttk.Button(stats, text="Cancel", command=self._cancel_current_action, state="disabled")
        self.cancel_button.pack(side="right")
        self.progress_bar = ttk.Progressbar(stats, variable=self.progress_var, maximum=100, mode="determinate", length=240)
        self.progress_bar.pack(side="right", padx=(12, 8))

        self.group_notebook = ttk.Notebook(container)
        self.group_notebook.pack(fill="both", expand=True)

        self.data_group = ttk.Frame(self.group_notebook, padding=6)
        self.data_group.columnconfigure(0, weight=1)
        self.data_group.rowconfigure(0, weight=1)
        self.data_notebook = ttk.Notebook(self.data_group)
        self.data_notebook.grid(row=0, column=0, sticky="nsew")

        self.summaries_group = ttk.Frame(self.group_notebook, padding=6)
        self.summaries_group.columnconfigure(0, weight=1)
        self.summaries_group.rowconfigure(0, weight=1)
        self.summaries_notebook = ttk.Notebook(self.summaries_group)
        self.summaries_notebook.grid(row=0, column=0, sticky="nsew")

        self.reports_group = ttk.Frame(self.group_notebook, padding=6)
        self.reports_group.columnconfigure(0, weight=1)
        self.reports_group.rowconfigure(0, weight=1)
        self.reports_notebook = ttk.Notebook(self.reports_group)
        self.reports_notebook.grid(row=0, column=0, sticky="nsew")

        self.settings_group = ttk.Frame(self.group_notebook, padding=6)
        self.settings_group.columnconfigure(0, weight=1)
        self.settings_group.rowconfigure(0, weight=1)
        self.settings_notebook = ttk.Notebook(self.settings_group)
        self.settings_notebook.grid(row=0, column=0, sticky="nsew")

        self.daily_table = DataTable(
            self.data_notebook,
            columns=["source_file", "date", "sheet", "employee", "st", "ot", "dt", "total", "expanded", "ranges"],
            headings=["Source File", "Date", "Source Sheet", "Employee", "ST", "OT", "DT", "Total", "Expanded Hours", "Source Ranges Used"],
            table_id="daily_raw",
            config_path=self.config_path,
            default_sort_column="sheet",
            default_sort_descending=True,
        )
        self.employee_editor = EmployeeListEditor(
            self.settings_notebook,
            on_email_changed=self._update_employee_email,
            on_request_edit_email=self._prompt_edit_employee_email,
        )
        self.weekly_rollup_table = DataTable(
            self.summaries_notebook,
            columns=["source_file", "week_start", "week_end", "employee", "st", "ot", "dt", "total", "expanded", "row_type"],
            headings=["Source File", "Week Start", "Week End", "Employee", "ST", "OT", "DT", "Total", "Expanded Hours", "Row Type"],
            table_id="weekly_rollup",
            config_path=self.config_path,
            default_sort_column="week_start",
            default_sort_descending=True,
        )
        self.combined_weekly_summary_table = DataTable(
            self.summaries_notebook,
            columns=["week_start", "week_end", "employee", "st", "ot", "dt", "total", "expanded"],
            headings=["Week Start", "Week End", "Employee", "ST", "OT", "DT", "Total", "Expanded Hours"],
            table_id="combined_summary",
            config_path=self.config_path,
            default_sort_column="week_start",
            default_sort_descending=True,
        )
        self.week_totals_table = DataTable(
            self.data_notebook,
            columns=["week_start", "week_end", "st", "ot", "dt", "total", "expanded"],
            headings=["Week Start", "Week End", "Whole Crew ST", "Whole Crew OT", "Whole Crew DT", "Whole Crew Total", "Expanded Hours"],
            table_id="week_totals",
            config_path=self.config_path,
            default_sort_column="week_start",
            default_sort_descending=True,
        )
        self.error_report_table = DataTable(
            self.reports_notebook,
            columns=[
                "employee",
                "week_start",
                "week_end",
                "rule",
                "trigger_date",
                "actual",
                "limit",
                "delta",
                "day_st",
                "day_ot",
                "day_dt",
                "sources",
                "reason",
                "breakdown",
            ],
            headings=[
                "Employee",
                "Week Start",
                "Week End",
                "Rule",
                "Trigger Date",
                "Actual",
                "Limit",
                "Delta",
                "Trigger Day ST",
                "Trigger Day OT",
                "Trigger Day DT",
                "Source Files",
                "Reason",
                "Daily Breakdown",
            ],
            table_id="error_report",
            config_path=self.config_path,
            default_sort_column="week_start",
            default_sort_descending=True,
        )
        self.parse_warnings_table = DataTable(
            self.reports_notebook,
            columns=["source_file", "sheet", "date", "issue", "details"],
            headings=["Source File", "Sheet", "Date", "Issue", "Details"],
            table_id="parse_warnings",
            config_path=self.config_path,
            default_sort_column="sheet",
            default_sort_descending=True,
        )
        self.workbook_health_table = DataTable(
            self.reports_notebook,
            columns=["source_file", "status", "details"],
            headings=["Source File", "Status", "Details"],
            table_id="workbook_health",
            config_path=self.config_path,
        )
        self.audit_data_trail_table = DataTable(
            self.reports_notebook,
            columns=["source_file", "date", "sheet", "employee", "st", "ot", "dt", "total", "expanded", "source_ranges", "audit"],
            headings=["Source File", "Date", "Sheet", "Employee", "ST", "OT", "DT", "Total", "Expanded Hours", "Source Ranges", "Audit"],
            table_id="audit_data_trail",
            config_path=self.config_path,
            default_sort_column="sheet",
            default_sort_descending=True,
        )
        self.email_drafts_frame = EmailDraftsFrame(
            self.reports_notebook,
            create_drafts_callback=self.create_email_drafts,
            sync_emails_callback=self.sync_outlook_emails,
            on_request_edit_email=self._prompt_edit_employee_email,
            save_templates_callback=self._save_email_templates,
        )
        self.email_drafts_frame.week_combo.bind("<<ComboboxSelected>>", self._on_email_week_changed)
        self.email_drafts_frame.set_templates(self.email_subject_template, self.email_body_template)
        self.groups_frame = EmployeeGroupsFrame(
            self.settings_notebook,
            on_groups_changed=self._update_employee_groups,
            on_request_edit_email=self._prompt_edit_employee_email,
            sync_emails_callback=self.sync_outlook_emails,
        )
        self.notes_frame = EmployeeNotesEditor(self.settings_notebook, on_notes_changed=self._update_employee_note)
        self.rules_frame = ttk.Frame(self.settings_notebook, padding=12)
        self._build_rules_tab()
        self.config_frame = ttk.Frame(self.settings_notebook, padding=12)
        self._build_config_tab()

        self.group_notebook.add(self.data_group, text="Data")
        self.group_notebook.add(self.summaries_group, text="Summaries")
        self.group_notebook.add(self.reports_group, text="Reports")
        self.group_notebook.add(self.settings_group, text="Settings")

        self._refresh_data_tabs()
        self.summaries_notebook.add(self.weekly_rollup_table, text="Weekly Rollup")
        self.summaries_notebook.add(self.combined_weekly_summary_table, text="Combined Summary")
        self.reports_notebook.add(self.error_report_table, text="Error Report")
        self.reports_notebook.add(self.parse_warnings_table, text="Sheet Parse Warnings")
        self.reports_notebook.add(self.workbook_health_table, text="Workbook Health")
        self.reports_notebook.add(self.audit_data_trail_table, text="Audit Data Trail")
        self.reports_notebook.add(self.email_drafts_frame, text="Email Drafts")
        self.settings_notebook.add(self.config_frame, text="Configuration")
        self.settings_notebook.add(self.employee_editor, text="Employee List")
        self.settings_notebook.add(self.notes_frame, text="Employee Notes")
        self.settings_notebook.add(self.groups_frame, text="Employee Groups")
        self.settings_notebook.add(self.rules_frame, text="Formatting Rules")
        self._refresh_filter_options()

    def _build_rules_tab(self) -> None:
        self.rules_frame.columnconfigure(1, weight=1)

        self.profile_var = tk.StringVar(value=self.current_profile_name)
        self.job_preset_var = tk.StringVar()
        self.st_threshold_var = tk.StringVar()
        self.ot_threshold_var = tk.StringVar()
        self.daily_st_threshold_var = tk.StringVar()
        self.max_hours_per_day_var = tk.StringVar()

        ttk.Label(self.rules_frame, text="Job Preset").grid(row=0, column=0, sticky="w", pady=(0, 8))
        self.job_preset_combo = ttk.Combobox(
            self.rules_frame,
            textvariable=self.job_preset_var,
            values=self._job_preset_names(),
        )
        self.job_preset_combo.grid(row=0, column=1, sticky="ew", pady=(0, 8))
        ttk.Button(self.rules_frame, text="Apply Job Preset", command=self._apply_job_preset).grid(row=0, column=2, padx=(8, 0), pady=(0, 8))
        ttk.Button(self.rules_frame, text="Save Job Preset", command=self._save_job_preset).grid(row=0, column=3, padx=(8, 0), pady=(0, 8))
        ttk.Button(self.rules_frame, text="Delete Job Preset", command=self._delete_job_preset).grid(row=0, column=4, padx=(8, 0), pady=(0, 8))

        ttk.Label(self.rules_frame, text="Profile").grid(row=1, column=0, sticky="w", pady=(0, 8))
        self.profile_combo = ttk.Combobox(
            self.rules_frame,
            textvariable=self.profile_var,
            state="readonly",
            values=self._profile_names(),
        )
        self.profile_combo.grid(row=1, column=1, sticky="ew", pady=(0, 8))
        self.profile_combo.bind("<<ComboboxSelected>>", self._on_profile_selected)

        ttk.Button(self.rules_frame, text="New Profile", command=self._create_profile).grid(row=1, column=2, padx=(8, 0), pady=(0, 8))
        ttk.Button(self.rules_frame, text="Delete Profile", command=self._delete_profile).grid(row=1, column=3, padx=(8, 0), pady=(0, 8))

        daily_st_label = ttk.Label(self.rules_frame, text="Daily ST Alert")
        daily_st_label.grid(row=2, column=0, sticky="w", pady=(0, 8))
        daily_st_entry = ttk.Entry(self.rules_frame, textvariable=self.daily_st_threshold_var)
        daily_st_entry.grid(row=2, column=1, sticky="ew", pady=(0, 8))
        daily_st_help = "How many regular time hours per week an employee can work (Usually 8 or 10)."
        ToolTip(daily_st_label, daily_st_help)
        ToolTip(daily_st_entry, daily_st_help)

        weekly_st_label = ttk.Label(self.rules_frame, text="Weekly ST Alert")
        weekly_st_label.grid(row=3, column=0, sticky="w", pady=(0, 8))
        weekly_st_entry = ttk.Entry(self.rules_frame, textvariable=self.st_threshold_var)
        weekly_st_entry.grid(row=3, column=1, sticky="ew", pady=(0, 8))
        weekly_st_help = "How many regular time hours per week an employee can work."
        ToolTip(weekly_st_label, weekly_st_help)
        ToolTip(weekly_st_entry, weekly_st_help)

        weekly_ot_label = ttk.Label(self.rules_frame, text="Weekly OT Alert")
        weekly_ot_label.grid(row=4, column=0, sticky="w", pady=(0, 8))
        weekly_ot_entry = ttk.Entry(self.rules_frame, textvariable=self.ot_threshold_var)
        weekly_ot_entry.grid(row=4, column=1, sticky="ew", pady=(0, 8))
        weekly_ot_help = "Some sites have limits to how much OT you can work before you automatically start making DT."
        ToolTip(weekly_ot_label, weekly_ot_help)
        ToolTip(weekly_ot_entry, weekly_ot_help)

        max_hours_label = ttk.Label(self.rules_frame, text="Max Hours Per Day")
        max_hours_label.grid(row=5, column=0, sticky="w", pady=(0, 8))
        max_hours_entry = ttk.Entry(self.rules_frame, textvariable=self.max_hours_per_day_var)
        max_hours_entry.grid(row=5, column=1, sticky="ew", pady=(0, 8))
        max_hours_help = "Max hours before fatigue management per day"
        ToolTip(max_hours_label, max_hours_help)
        ToolTip(max_hours_entry, max_hours_help)

        ttk.Button(self.rules_frame, text="Apply Rules", command=self._apply_rule_changes).grid(row=6, column=1, sticky="w", pady=(8, 0))

        note = (
            "Enter weekly alert thresholds for this profile. Leave a field blank to disable that alert. "
            "These rules apply to employee rows in Weekly Rollup and Combined Summary."
        )
        ttk.Label(self.rules_frame, text=note, wraplength=700, justify="left").grid(
            row=7, column=0, columnspan=5, sticky="w", pady=(12, 0)
        )

        self._populate_rule_editor()

    def _build_config_tab(self) -> None:
        self.config_frame.columnconfigure(1, weight=1)

        self.disable_name_typos_var = tk.BooleanVar(value=self.app_settings.disable_name_typo_notifications)
        self.show_daily_raw_var = tk.BooleanVar(value=self.app_settings.show_daily_raw_tab)
        self.hash_poll_minutes_var = tk.StringVar(value=str(self.app_settings.hash_poll_minutes))

        ttk.Checkbutton(
            self.config_frame,
            text="Disable name typo notifications",
            variable=self.disable_name_typos_var,
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

        ttk.Label(self.config_frame, text="Check source DSS(s) frequency (minutes)").grid(row=1, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self.config_frame, textvariable=self.hash_poll_minutes_var, width=12).grid(
            row=1, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Checkbutton(
            self.config_frame,
            text="Show Daily Raw tab",
            variable=self.show_daily_raw_var,
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 8))

        ttk.Button(self.config_frame, text="Apply Settings", command=self._apply_app_settings).grid(
            row=3, column=0, sticky="w", pady=(8, 0)
        )

        maintenance = ttk.LabelFrame(self.config_frame, text="Maintenance", padding=8)
        maintenance.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(16, 0))
        ttk.Button(maintenance, text="Reset All Settings to Default", command=self._reset_all_settings).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Button(maintenance, text="Clear Cached DSSs", command=self._clear_cached_dsss).grid(
            row=1, column=0, sticky="w", pady=(8, 0)
        )
        ttk.Button(maintenance, text="Clear Stored Emails", command=self._clear_stored_emails).grid(
            row=1, column=1, sticky="w", padx=(12, 0), pady=(8, 0)
        )
        ttk.Button(maintenance, text="Clear All Stored Data", command=self._clear_all_stored_data).grid(
            row=2, column=0, sticky="w", pady=(8, 0)
        )

        diagnostics = ttk.LabelFrame(self.config_frame, text="Diagnostics", padding=8)
        diagnostics.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(16, 0))
        ttk.Button(diagnostics, text="Show App Data Folder", command=self._show_app_data_folder).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Button(diagnostics, text="Export Diagnostic Snapshot", command=self._export_diagnostic_snapshot).grid(
            row=0, column=1, sticky="w", padx=(12, 0)
        )
        ttk.Button(diagnostics, text="Test Outlook Connection", command=self._test_outlook_connection).grid(
            row=1, column=0, sticky="w", pady=(8, 0)
        )
        ttk.Button(diagnostics, text="Show Loaded DSS Status", command=self._show_loaded_dss_status).grid(
            row=1, column=1, sticky="w", padx=(12, 0), pady=(8, 0)
        )
        ttk.Button(diagnostics, text="Submit Bug Report", command=self._submit_bug_report).grid(
            row=2, column=0, sticky="w", pady=(8, 0)
        )
        ttk.Button(diagnostics, text="Check for Updates", command=self._check_for_updates).grid(
            row=2, column=1, sticky="w", padx=(12, 0), pady=(8, 0)
        )
        ttk.Label(diagnostics, textvariable=self.update_status_var, wraplength=700, justify="left").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(8, 0)
        )

        note = (
            "These settings control background notifications, how often the app checks loaded DSS files for changes, "
            "and whether the Daily Raw page is visible in the Data group."
        )
        ttk.Label(self.config_frame, text=note, wraplength=700, justify="left").grid(
            row=6, column=0, columnspan=2, sticky="w", pady=(12, 0)
        )

    def _profile_names(self) -> list[str]:
        return sorted(self.formatting_profiles)

    def _job_preset_names(self) -> list[str]:
        return sorted(self.job_presets)

    def _active_profile(self) -> FormattingProfile:
        return self.formatting_profiles[self.current_profile_name]

    def _populate_rule_editor(self) -> None:
        profile = self._active_profile()
        self.profile_var.set(profile.name)
        self.profile_combo.configure(values=self._profile_names())
        self.job_preset_combo.configure(values=self._job_preset_names())
        self.st_threshold_var.set("" if profile.st_threshold is None else fmt_hours(profile.st_threshold))
        self.ot_threshold_var.set("" if profile.ot_threshold is None else fmt_hours(profile.ot_threshold))
        self.daily_st_threshold_var.set("" if profile.daily_st_threshold is None else fmt_hours(profile.daily_st_threshold))
        self.max_hours_per_day_var.set("" if profile.max_hours_per_day is None else fmt_hours(profile.max_hours_per_day))

    def _on_profile_selected(self, _event=None) -> None:
        selected = self.profile_var.get()
        if selected in self.formatting_profiles:
            self.current_profile_name = selected
            self._populate_rule_editor()
            self._persist_profiles()
            self._refresh_alert_rendering()

    def _apply_rule_changes(self) -> None:
        try:
            profile = FormattingProfile(
                name=self.current_profile_name,
                st_threshold=parse_threshold_value(self.st_threshold_var.get()),
                ot_threshold=parse_threshold_value(self.ot_threshold_var.get()),
                daily_st_threshold=parse_threshold_value(self.daily_st_threshold_var.get()),
                max_hours_per_day=parse_threshold_value(self.max_hours_per_day_var.get()),
            )
        except ValueError:
            messagebox.showerror("Formatting Rules", "Thresholds must be numbers or blank.")
            return

        self.formatting_profiles[profile.name] = profile
        self._persist_profiles()
        self._refresh_alert_rendering()
        messagebox.showinfo("Formatting Rules", f"Saved rules for profile '{profile.name}'.")

    def _create_profile(self) -> None:
        name = simpledialog.askstring("New Profile", "Profile name:", parent=self)
        if not name:
            return
        name = name.strip()
        if not name:
            return
        if name in self.formatting_profiles:
            messagebox.showerror("Formatting Rules", "A profile with that name already exists.")
            return

        base = self._active_profile()
        self.formatting_profiles[name] = FormattingProfile(
            name=name,
            st_threshold=base.st_threshold,
            ot_threshold=base.ot_threshold,
            daily_st_threshold=base.daily_st_threshold,
            max_hours_per_day=base.max_hours_per_day,
        )
        self.current_profile_name = name
        self._persist_profiles()
        self._populate_rule_editor()

    def _delete_profile(self) -> None:
        if len(self.formatting_profiles) == 1:
            messagebox.showerror("Formatting Rules", "At least one profile must remain.")
            return
        profile_name = self.current_profile_name
        if not messagebox.askyesno("Delete Profile", f"Delete profile '{profile_name}'?"):
            return
        del self.formatting_profiles[profile_name]
        self.current_profile_name = self._profile_names()[0]
        self._persist_profiles()
        self._populate_rule_editor()
        self._refresh_alert_rendering()

    def _persist_profiles(self) -> None:
        save_formatting_profiles(self.config_path, self.formatting_profiles, self.current_profile_name)

    def _persist_employee_emails(self) -> None:
        save_employee_emails(self.config_path, self.employee_emails)

    def _persist_employee_groups(self) -> None:
        save_employee_groups(self.config_path, self.employee_groups)

    def _persist_employee_notes(self) -> None:
        save_employee_notes(self.config_path, self.employee_notes)

    def _persist_app_settings(self) -> None:
        save_app_settings(self.config_path, self.app_settings)

    def _persist_job_presets(self) -> None:
        save_job_presets(self.config_path, self.job_presets)

    def _all_layout_tables(self) -> list[DataTable]:
        return [
            self.daily_table,
            self.weekly_rollup_table,
            self.combined_weekly_summary_table,
            self.week_totals_table,
            self.error_report_table,
            self.parse_warnings_table,
            self.workbook_health_table,
            self.audit_data_trail_table,
        ]

    def _reload_defaults_into_ui(self) -> None:
        self.disable_name_typos_var.set(self.app_settings.disable_name_typo_notifications)
        self.show_daily_raw_var.set(self.app_settings.show_daily_raw_tab)
        self.hash_poll_minutes_var.set(str(self.app_settings.hash_poll_minutes))
        self.email_drafts_frame.set_templates(self.email_subject_template, self.email_body_template)
        self._populate_rule_editor()
        self._refresh_data_tabs()
        self._refresh_filter_options()
        self.groups_frame.set_data(
            self.current_data.employee_names if self.current_data is not None else [],
            self.employee_groups,
            self.employee_emails,
        )
        employee_names = self.current_data.employee_names if self.current_data is not None else []
        self.employee_editor.set_names(employee_names, self.employee_emails)
        self.notes_frame.set_data(employee_names, self.employee_notes)
        self.job_preset_combo.configure(values=self._job_preset_names())
        for table in self._all_layout_tables():
            table.reset_layout()
        if self.current_data is not None:
            self._render_data(self.current_data)

    def _save_email_templates(self) -> None:
        self.email_subject_template = self.email_drafts_frame.get_subject_template()
        self.email_body_template = self.email_drafts_frame.get_body_template()
        save_email_templates(self.config_path, self.email_subject_template, self.email_body_template)
        messagebox.showinfo("Email Templates", "Saved email templates.")

    def _apply_app_settings(self) -> None:
        try:
            hash_poll_minutes = max(1, int(self.hash_poll_minutes_var.get().strip()))
        except ValueError:
            messagebox.showerror("Configuration", "Hash poll frequency must be a whole number of minutes.")
            return

        self.app_settings = AppSettings(
            disable_name_typo_notifications=bool(self.disable_name_typos_var.get()),
            hash_poll_minutes=hash_poll_minutes,
            show_daily_raw_tab=bool(self.show_daily_raw_var.get()),
        )
        self.hash_poll_interval_ms = self.app_settings.hash_poll_minutes * 60 * 1000
        self._persist_app_settings()
        self._refresh_data_tabs()
        self._schedule_hash_monitor()
        messagebox.showinfo("Configuration", "Saved application settings.")

    def _reset_all_settings(self) -> None:
        if not messagebox.askyesno(
            "Reset Settings",
            "Reset saved settings, formatting rules, email templates, and table layouts to defaults?\n\n"
            "This does not delete stored employee emails, employee groups, or cached DSS data.",
        ):
            return

        remove_config_keys(
            self.config_path,
            [
                "app_settings",
                "profiles",
                "current_profile",
                "email_subject_template",
                "email_body_template",
                "table_layouts",
                "job_presets",
            ],
        )
        self.app_settings = AppSettings()
        self.hash_poll_interval_ms = self.app_settings.hash_poll_minutes * 60 * 1000
        self.formatting_profiles, self.current_profile_name = load_formatting_profiles(self.config_path)
        self.email_subject_template, self.email_body_template = load_email_templates(self.config_path)
        self.job_presets = {}
        self._reload_defaults_into_ui()
        self._schedule_hash_monitor()
        messagebox.showinfo("Configuration", "Settings were reset to defaults.")

    def _clear_cached_dsss(self) -> None:
        if not messagebox.askyesno("Clear Cached DSSs", "Delete all cached parsed DSS files?"):
            return
        deleted = clear_cache_files(self.cache_dir)
        messagebox.showinfo("Configuration", f"Deleted {deleted} cached DSS file(s).")

    def _clear_stored_emails(self) -> None:
        if not messagebox.askyesno("Clear Stored Emails", "Delete all saved employee email addresses?"):
            return
        self.employee_emails = {}
        remove_config_keys(self.config_path, ["employee_emails"])
        if self.current_data is not None:
            self._render_data(self.current_data)
        else:
            self.employee_editor.set_names([], self.employee_emails)
            self.groups_frame.set_data([], self.employee_groups, self.employee_emails)
        messagebox.showinfo("Configuration", "Stored employee emails were cleared.")

    def _clear_all_stored_data(self) -> None:
        if not messagebox.askyesno(
            "Clear All Stored Data",
            "Delete all stored tracker data?\n\n"
            "This clears cached DSS data, employee emails, employee groups, formatting rules, "
            "templates, saved layouts, and app settings.",
        ):
            return

        clear_cache_files(self.cache_dir)
        if self.config_path.exists():
            self.config_path.unlink()

        self.employee_emails = {}
        self.employee_groups = {}
        self.employee_notes = {}
        self.job_presets = {}
        self.app_settings = AppSettings()
        self.hash_poll_interval_ms = self.app_settings.hash_poll_minutes * 60 * 1000
        self.formatting_profiles, self.current_profile_name = load_formatting_profiles(self.config_path)
        self.email_subject_template, self.email_body_template = load_email_templates(self.config_path)
        self._reload_defaults_into_ui()
        self._schedule_hash_monitor()
        messagebox.showinfo("Configuration", "All stored tracker data was cleared.")

    def _show_app_data_folder(self) -> None:
        try:
            if os.name == "nt":
                os.startfile(self.app_root)  # type: ignore[attr-defined]
            else:
                messagebox.showinfo("App Data Folder", str(self.app_root))
        except Exception as exc:
            messagebox.showerror("App Data Folder", f"Could not open the app data folder.\n\n{exc}")

    def _export_diagnostic_snapshot(self) -> None:
        try:
            export_path = self._write_diagnostic_snapshot()
        except OSError as exc:
            messagebox.showerror("Diagnostic Snapshot", f"Could not write the diagnostic snapshot.\n\n{exc}")
            return
        messagebox.showinfo("Diagnostic Snapshot", f"Saved diagnostic snapshot:\n{export_path}")

    def _write_diagnostic_snapshot(self) -> Path:
        snapshot = {
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "app_root": str(self.app_root),
            "cache_dir": str(self.cache_dir),
            "config_path": str(self.config_path),
            "python_version": sys.version,
            "app_version": APP_VERSION,
            "os_name": os.name,
            "app_settings": {
                "disable_name_typo_notifications": self.app_settings.disable_name_typo_notifications,
                "hash_poll_minutes": self.app_settings.hash_poll_minutes,
                "show_daily_raw_tab": self.app_settings.show_daily_raw_tab,
            },
            "formatting_profiles": sorted(self.formatting_profiles),
            "current_profile_name": self.current_profile_name,
            "employee_email_count": len([email for email in self.employee_emails.values() if email.strip()]),
            "employee_group_count": len(self.employee_groups),
            "cache_file_count": len(list(self.cache_dir.glob("*.json"))),
            "loaded_dss": [],
        }

        if self.current_data is not None:
            for source_path in self.current_data.source_paths:
                snapshot["loaded_dss"].append(
                    {
                        "path": str(source_path),
                        "hash": self.current_data.file_hashes.get(source_path, ""),
                        "reused": source_path in self.current_data.reused_paths,
                        "reloaded": source_path in self.current_data.reloaded_paths,
                        "hash_alerted": source_path in self._hash_alerted_paths,
                    }
                )
            snapshot["current_data"] = {
                "daily_record_count": len(self.current_data.daily_records),
                "employee_count": len(self.current_data.employee_names),
                "week_count": len(self.current_data.week_totals),
                "source_count": len(self.current_data.source_paths),
            }
        else:
            snapshot["current_data"] = None

        export_path = self.app_root / f"diagnostic_snapshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        export_path.write_text(json.dumps(snapshot, indent=2), encoding="utf-8")
        return export_path

    def _test_outlook_connection(self) -> None:
        if pythoncom is None or win32com is None:
            messagebox.showerror("Outlook Connection", "Outlook integration requires pywin32 and desktop Outlook.")
            return

        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            current_user = getattr(namespace, "CurrentUser", None)
            user_name = str(getattr(current_user, "Name", "")).strip() or "Unknown user"
        except Exception as exc:
            messagebox.showerror("Outlook Connection", f"Could not connect to Outlook.\n\n{exc}")
            return
        finally:
            pythoncom.CoUninitialize()

        messagebox.showinfo("Outlook Connection", f"Connected to desktop Outlook successfully.\nUser: {user_name}")

    def _show_loaded_dss_status(self) -> None:
        if self.current_data is None or not self.current_data.source_paths:
            messagebox.showinfo("Loaded DSS Status", "No DSS workbooks are currently loaded.")
            return

        lines = [
            f"Loaded DSS files: {len(self.current_data.source_paths)}",
            f"Daily records: {len(self.current_data.daily_records)}",
            f"Employees: {len(self.current_data.employee_names)}",
            f"Weeks: {len(self.current_data.week_totals)}",
            "",
        ]
        for source_path in self.current_data.source_paths:
            status_bits = []
            if source_path in self.current_data.reloaded_paths:
                status_bits.append("reloaded")
            if source_path in self.current_data.reused_paths:
                status_bits.append("reused")
            if source_path in self._hash_alerted_paths:
                status_bits.append("changed since load")
            if not status_bits:
                status_bits.append("loaded")
            lines.append(str(source_path))
            lines.append(f"Hash: {self.current_data.file_hashes.get(source_path, '')}")
            lines.append(f"Cache: {self.current_data.cache_status_by_path.get(source_path, 'Unknown')}")
            lines.append(f"Status: {', '.join(status_bits)}")
            lines.append("")

        messagebox.showinfo("Loaded DSS Status", "\n".join(lines).rstrip())

    def _check_for_updates(self) -> None:
        if self._update_check_in_progress:
            return
        self._update_check_in_progress = True
        self.update_status_var.set(f"Checking GitHub releases from version {APP_VERSION}...")
        worker = threading.Thread(target=self._check_for_updates_worker, daemon=True)
        worker.start()

    def _check_for_updates_worker(self) -> None:
        try:
            release_info = fetch_latest_release_info()
        except Exception as exc:
            self.after(0, lambda: self._handle_update_check_error(exc))
            return
        self.after(0, lambda: self._handle_update_check_result(release_info))

    def _handle_update_check_error(self, exc: Exception) -> None:
        self._update_check_in_progress = False
        self.update_status_var.set(f"Installed version: {APP_VERSION}")
        messagebox.showerror("Check for Updates", f"Could not check for updates.\n\n{exc}")

    def _handle_update_check_result(self, release_info: dict[str, object]) -> None:
        self._update_check_in_progress = False
        latest_version = str(release_info.get("version", "")).strip()
        latest_tag = str(release_info.get("tag_name", "")).strip()
        html_url = str(release_info.get("html_url", "")).strip()
        published_at = str(release_info.get("published_at", "")).strip()
        asset_names = release_info.get("asset_names", [])
        assets_text = ", ".join(asset_names) if isinstance(asset_names, list) and asset_names else "No assets listed"

        if latest_version and is_newer_version(latest_version, APP_VERSION):
            self.update_status_var.set(f"Update available: {latest_tag or latest_version} (installed: {APP_VERSION})")
            messagebox.showinfo(
                "Check for Updates",
                "A newer version is available.\n\n"
                f"Installed: {APP_VERSION}\n"
                f"Latest: {latest_tag or latest_version}\n"
                f"Published: {published_at or 'Unknown'}\n"
                f"Assets: {assets_text}\n\n"
                f"Release page:\n{html_url}",
            )
            return

        self.update_status_var.set(f"Installed version: {APP_VERSION} (up to date)")
        messagebox.showinfo(
            "Check for Updates",
            "You are up to date.\n\n"
            f"Installed: {APP_VERSION}\n"
            f"Latest release: {latest_tag or latest_version or 'Unknown'}",
        )

    def _submit_bug_report(self) -> None:
        try:
            snapshot_path = self._write_diagnostic_snapshot()
            loaded_sources = self.current_data.source_paths if self.current_data is not None else []
            cache_status_by_path = self.current_data.cache_status_by_path if self.current_data is not None else {}
            subject = f"DSS Hours Tracker Bug Report - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            html_body = build_bug_report_html(
                self.current_profile_name,
                self.app_root,
                snapshot_path,
                loaded_sources,
                cache_status_by_path,
            )
            create_bug_report_draft(
                BUG_REPORT_EMAIL,
                subject,
                html_body,
                attachment_path=snapshot_path,
            )
        except Exception as exc:
            messagebox.showerror("Submit Bug Report", f"Could not create the bug report draft.\n\n{exc}")
            return

        messagebox.showinfo(
            "Submit Bug Report",
            "Created an Outlook draft bug report addressed to "
            f"{BUG_REPORT_EMAIL} and attached a diagnostic snapshot.",
        )

    def _refresh_data_tabs(self) -> None:
        current_tabs = set(self.data_notebook.tabs())
        daily_raw_tab = str(self.daily_table)
        week_totals_tab = str(self.week_totals_table)

        if self.app_settings.show_daily_raw_tab:
            if daily_raw_tab not in current_tabs:
                if current_tabs:
                    self.data_notebook.insert(0, self.daily_table, text="Daily Raw")
                else:
                    self.data_notebook.add(self.daily_table, text="Daily Raw")
        elif daily_raw_tab in current_tabs:
            self.data_notebook.forget(self.daily_table)

        current_tabs = set(self.data_notebook.tabs())
        if week_totals_tab not in current_tabs:
            self.data_notebook.add(self.week_totals_table, text="Week Totals")

    def _update_employee_email(self, employee: str, email: str) -> None:
        self.employee_emails[employee] = email.strip()
        self._persist_employee_emails()
        if self.current_data is not None:
            self._refresh_email_preview(self.current_data)

    def _prompt_edit_employee_email(self, employee: str) -> None:
        current_email = self.employee_emails.get(employee, "")
        new_email = simpledialog.askstring(
            "Edit Employee Email",
            f"Email for {employee}:",
            initialvalue=current_email,
            parent=self,
        )
        if new_email is None:
            return
        self._update_employee_email(employee, new_email.strip())
        if self.current_data is not None:
            self._render_data(self.current_data)

    def _update_employee_groups(self, employee_groups: dict[str, list[str]]) -> None:
        self.employee_groups = {name: sorted(members) for name, members in employee_groups.items()}
        self._persist_employee_groups()
        self._refresh_filter_options()
        if self.current_data is not None:
            self._render_data(self.current_data)

    def _update_employee_note(self, employee: str, note: str) -> None:
        self.employee_notes[employee] = note.strip()
        self._persist_employee_notes()

    def _save_job_preset(self) -> None:
        job_name = self.job_preset_var.get().strip()
        if not job_name:
            messagebox.showerror("Formatting Rules", "Enter a job preset name first.")
            return
        self.job_presets[job_name] = self.current_profile_name
        self._persist_job_presets()
        self.job_preset_combo.configure(values=self._job_preset_names())
        messagebox.showinfo("Formatting Rules", f"Saved job preset '{job_name}' -> profile '{self.current_profile_name}'.")

    def _apply_job_preset(self) -> None:
        job_name = self.job_preset_var.get().strip()
        profile_name = self.job_presets.get(job_name, "")
        if not profile_name or profile_name not in self.formatting_profiles:
            messagebox.showerror("Formatting Rules", "That job preset is missing or points to a deleted profile.")
            return
        self.current_profile_name = profile_name
        self._populate_rule_editor()
        self._persist_profiles()
        self._refresh_alert_rendering()

    def _delete_job_preset(self) -> None:
        job_name = self.job_preset_var.get().strip()
        if not job_name or job_name not in self.job_presets:
            return
        if not messagebox.askyesno("Delete Job Preset", f"Delete job preset '{job_name}'?"):
            return
        del self.job_presets[job_name]
        self._persist_job_presets()
        self.job_preset_var.set("")
        self.job_preset_combo.configure(values=self._job_preset_names())

    def export_current_view(self) -> None:
        table = self._current_export_table()
        if table is None:
            messagebox.showerror("Export Current View", "The current page does not support table export.")
            return
        headings, rows = table.displayed_columns_and_rows()
        if not rows:
            messagebox.showerror("Export Current View", "There is no data to export in the current view.")
            return
        target = filedialog.asksaveasfilename(
            title="Export Current View",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("Excel Workbook", "*.xlsx")],
        )
        if not target:
            return
        export_path = Path(target)
        try:
            if export_path.suffix.lower() == ".xlsx":
                workbook = Workbook()
                worksheet = workbook.active
                worksheet.title = "Export"
                worksheet.append(headings)
                for row in rows:
                    worksheet.append(list(row))
                workbook.save(export_path)
            else:
                def csv_escape(value: str) -> str:
                    return '"' + value.replace('"', '""') + '"'

                lines = [",".join(csv_escape(value) for value in headings)]
                for row in rows:
                    lines.append(",".join(csv_escape(value) for value in row))
                export_path.write_text("\n".join(lines), encoding="utf-8")
        except Exception as exc:
            messagebox.showerror("Export Current View", f"Could not export the current view.\n\n{exc}")
            return
        messagebox.showinfo("Export Current View", f"Exported current view to:\n{export_path}")

    def _current_export_table(self) -> DataTable | None:
        current_group = self.group_notebook.select()
        if current_group == str(self.data_group):
            current_page = self.data_notebook.select()
            mapping = {
                str(self.daily_table): self.daily_table,
                str(self.week_totals_table): self.week_totals_table,
            }
            return mapping.get(current_page)
        if current_group == str(self.summaries_group):
            current_page = self.summaries_notebook.select()
            mapping = {
                str(self.weekly_rollup_table): self.weekly_rollup_table,
                str(self.combined_weekly_summary_table): self.combined_weekly_summary_table,
            }
            return mapping.get(current_page)
        if current_group == str(self.reports_group):
            current_page = self.reports_notebook.select()
            mapping = {
                str(self.error_report_table): self.error_report_table,
                str(self.parse_warnings_table): self.parse_warnings_table,
                str(self.workbook_health_table): self.workbook_health_table,
                str(self.audit_data_trail_table): self.audit_data_trail_table,
                str(self.email_drafts_frame): self.email_drafts_frame.preview_table,
            }
            return mapping.get(current_page)
        return None

    def _refresh_alert_rendering(self) -> None:
        if self.current_data is not None:
            self._render_data(self.current_data)

    def _refresh_outlook_sync_button(self) -> None:
        return

    def _begin_cancellable_action(self, operation_name: str, cancel_event: threading.Event) -> None:
        self._cancel_event = cancel_event
        self._active_operation_name = operation_name
        self.cancel_button.configure(state="normal")
        if not self._is_loading:
            self.loading_label.configure(text=operation_name)
            self.stats_label.configure(text=f"{operation_name} in progress...")

    def _end_cancellable_action(self, cancel_event: threading.Event | None = None) -> None:
        if cancel_event is not None and self._cancel_event is not cancel_event:
            return
        self._cancel_event = None
        self._active_operation_name = ""
        self.cancel_button.configure(state="disabled")
        if not self._is_loading:
            self.loading_label.configure(text="")
            self._refresh_stats_summary()

    def _cancel_current_action(self) -> None:
        if self._cancel_event is None:
            return
        self._cancel_event.set()
        self.loading_label.configure(text="Cancelling...")
        if self._active_operation_name:
            self.stats_label.configure(text=f"Cancelling {self._active_operation_name}...")

    def _refresh_stats_summary(self) -> None:
        if self.current_data is None:
            self.stats_label.configure(text="Load a DSS workbook to view daily and weekly labour summaries.")
            return
        memory_hits = sum(1 for status in self.current_data.cache_status_by_path.values() if status == "Memory Hit")
        disk_hits = sum(1 for status in self.current_data.cache_status_by_path.values() if status == "Disk Hit")
        misses = sum(1 for status in self.current_data.cache_status_by_path.values() if status == "Miss")
        self.stats_label.configure(
            text=(
                f"{len(self.current_data.daily_records)} daily records, "
                f"{len(self.current_data.employee_names)} employees, "
                f"{len(self.current_data.week_totals)} weeks, "
                f"{len(self.current_data.source_paths)} DSS file(s), "
                f"rules: {self.current_profile_name}, "
                f"cache: {memory_hits} memory hit / {disk_hits} disk hit / {misses} miss"
            )
        )

    def _current_filter_selection(self) -> FilterSelection:
        employee_names = self.current_data.employee_names if self.current_data is not None else []
        selected_names = tuple(
            employee for employee in employee_names
            if self._employee_filter_state.get(employee, True)
        )
        if not employee_names or len(selected_names) == len(employee_names):
            return FilterSelection(mode="all", value="All Employees")
        return FilterSelection(mode="employee_multi", value=", ".join(selected_names), values=selected_names)

    def _pf_identifier_by_source_path(self) -> dict[Path, str]:
        if self.current_data is None:
            return {}
        return {
            source_path: extract_pf_identifier(source_path.name)
            for source_path in self.current_data.source_paths
        }

    def _current_pf_selection(self) -> set[str]:
        pf_values = set(self._pf_identifier_by_source_path().values())
        selected = {
            pf_value
            for pf_value in pf_values
            if self._pf_filter_state.get(pf_value, True)
        }
        return selected if selected else pf_values

    def _refresh_filter_options(self) -> None:
        employee_names = self.current_data.employee_names if self.current_data is not None else []
        previous_state = dict(self._employee_filter_state)
        if not previous_state:
            self._employee_filter_state = {employee: True for employee in employee_names}
        else:
            self._employee_filter_state = {
                employee: previous_state.get(employee, True)
                for employee in employee_names
            }

        pf_values = sorted(set(self._pf_identifier_by_source_path().values()))
        previous_pf_state = dict(self._pf_filter_state)
        if not previous_pf_state:
            self._pf_filter_state = {pf_value: True for pf_value in pf_values}
        else:
            self._pf_filter_state = {
                pf_value: previous_pf_state.get(pf_value, True)
                for pf_value in pf_values
            }
        if len(pf_values) <= 1:
            self._pf_filter_state = {pf_value: True for pf_value in pf_values}

        self._employee_filter_vars = {}
        self._pf_filter_vars = {}
        self._rebuild_filter_popup_content()
        self._rebuild_pf_filter_popup_content()
        self._update_filter_button_label()
        self._update_pf_filter_button_label()
        self.pf_filter_button.configure(state="normal" if len(pf_values) > 1 else "disabled")

    def _toggle_filter_popup(self) -> None:
        if self.filter_popup is not None and self.filter_popup.winfo_exists():
            self._close_filter_popup()
        else:
            self._open_filter_popup()

    def _toggle_pf_filter_popup(self) -> None:
        if self.pf_filter_popup is not None and self.pf_filter_popup.winfo_exists():
            self._close_pf_filter_popup()
        else:
            self._open_pf_filter_popup()

    def _open_filter_popup(self) -> None:
        if self.filter_popup is not None and self.filter_popup.winfo_exists():
            return
        self.filter_popup = tk.Toplevel(self)
        self.filter_popup.withdraw()
        self.filter_popup.wm_overrideredirect(True)
        self.filter_popup.transient(self)
        self.filter_popup.bind("<FocusOut>", self._on_filter_popup_focus_out, add="+")
        self.filter_popup.bind("<Escape>", lambda _event: self._close_filter_popup(), add="+")
        self.filter_popup_content = ttk.Frame(self.filter_popup, padding=6, relief="solid", borderwidth=1)
        self.filter_popup_content.pack(fill="both", expand=True)
        self._rebuild_filter_popup_content()
        x = self.filter_button.winfo_rootx()
        y = self.filter_button.winfo_rooty() + self.filter_button.winfo_height() + 2
        self.filter_popup.wm_geometry(f"+{x}+{y}")
        self.filter_popup.deiconify()
        self.filter_popup.lift()
        self.filter_popup.focus_force()

    def _close_filter_popup(self) -> None:
        if self.filter_popup is not None and self.filter_popup.winfo_exists():
            self.filter_popup.destroy()
        self.filter_popup = None
        self.filter_popup_content = None

    def _open_pf_filter_popup(self) -> None:
        if self.pf_filter_popup is not None and self.pf_filter_popup.winfo_exists():
            return
        if self.pf_filter_button.instate(("disabled",)):
            return
        self.pf_filter_popup = tk.Toplevel(self)
        self.pf_filter_popup.withdraw()
        self.pf_filter_popup.wm_overrideredirect(True)
        self.pf_filter_popup.transient(self)
        self.pf_filter_popup.bind("<FocusOut>", self._on_pf_filter_popup_focus_out, add="+")
        self.pf_filter_popup.bind("<Escape>", lambda _event: self._close_pf_filter_popup(), add="+")
        self.pf_filter_popup_content = ttk.Frame(self.pf_filter_popup, padding=6, relief="solid", borderwidth=1)
        self.pf_filter_popup_content.pack(fill="both", expand=True)
        self._rebuild_pf_filter_popup_content()
        x = self.pf_filter_button.winfo_rootx()
        y = self.pf_filter_button.winfo_rooty() + self.pf_filter_button.winfo_height() + 2
        self.pf_filter_popup.wm_geometry(f"+{x}+{y}")
        self.pf_filter_popup.deiconify()
        self.pf_filter_popup.lift()
        self.pf_filter_popup.focus_force()

    def _close_pf_filter_popup(self) -> None:
        if self.pf_filter_popup is not None and self.pf_filter_popup.winfo_exists():
            self.pf_filter_popup.destroy()
        self.pf_filter_popup = None
        self.pf_filter_popup_content = None

    def _rebuild_filter_popup_content(self) -> None:
        if self.filter_popup_content is None or not self.filter_popup_content.winfo_exists():
            return
        for child in self.filter_popup_content.winfo_children():
            child.destroy()

        ttk.Button(self.filter_popup_content, text="All Employees", command=self._select_all_employees).pack(
            fill="x", anchor="w"
        )
        ttk.Button(self.filter_popup_content, text="Uncheck All", command=self._clear_employee_selection).pack(
            fill="x", anchor="w", pady=(4, 0)
        )

        employee_names = self.current_data.employee_names if self.current_data is not None else []
        if employee_names:
            ttk.Separator(self.filter_popup_content, orient="horizontal").pack(fill="x", pady=6)

        checklist = ttk.Frame(self.filter_popup_content)
        checklist.pack(fill="both", expand=True)
        for employee in employee_names:
            var = tk.BooleanVar(value=self._employee_filter_state.get(employee, True))
            self._employee_filter_vars[employee] = var
            ttk.Checkbutton(
                checklist,
                text=employee,
                variable=var,
                command=self._on_employee_filter_menu_changed,
            ).pack(anchor="w", fill="x")

    def _rebuild_pf_filter_popup_content(self) -> None:
        if self.pf_filter_popup_content is None or not self.pf_filter_popup_content.winfo_exists():
            return
        for child in self.pf_filter_popup_content.winfo_children():
            child.destroy()

        ttk.Button(self.pf_filter_popup_content, text="All PFs", command=self._select_all_pfs).pack(
            fill="x", anchor="w"
        )
        if not self.pf_filter_button.instate(("disabled",)):
            ttk.Button(self.pf_filter_popup_content, text="Uncheck All", command=self._clear_pf_selection).pack(
                fill="x", anchor="w", pady=(4, 0)
            )

        pf_values = sorted(self._pf_filter_state)
        if pf_values:
            ttk.Separator(self.pf_filter_popup_content, orient="horizontal").pack(fill="x", pady=6)

        checklist = ttk.Frame(self.pf_filter_popup_content)
        checklist.pack(fill="both", expand=True)
        disable_changes = len(pf_values) <= 1
        for pf_value in pf_values:
            var = tk.BooleanVar(value=self._pf_filter_state.get(pf_value, True))
            self._pf_filter_vars[pf_value] = var
            ttk.Checkbutton(
                checklist,
                text=pf_value,
                variable=var,
                command=self._on_pf_filter_menu_changed,
                state="disabled" if disable_changes else "normal",
            ).pack(anchor="w", fill="x")

    def _on_filter_popup_focus_out(self, _event=None) -> None:
        self.after(1, self._maybe_close_filter_popup)

    def _maybe_close_filter_popup(self) -> None:
        if self.filter_popup is None or not self.filter_popup.winfo_exists():
            return
        focused_widget = self.focus_displayof()
        if focused_widget is None:
            self._close_filter_popup()
            return
        current = focused_widget
        while current is not None:
            if current == self.filter_popup or current == self.filter_button:
                return
            current = current.master
        self._close_filter_popup()

    def _on_pf_filter_popup_focus_out(self, _event=None) -> None:
        self.after(1, self._maybe_close_pf_filter_popup)

    def _maybe_close_pf_filter_popup(self) -> None:
        if self.pf_filter_popup is None or not self.pf_filter_popup.winfo_exists():
            return
        focused_widget = self.focus_displayof()
        if focused_widget is None:
            self._close_pf_filter_popup()
            return
        current = focused_widget
        while current is not None:
            if current == self.pf_filter_popup or current == self.pf_filter_button:
                return
            current = current.master
        self._close_pf_filter_popup()

    def _select_all_employees(self) -> None:
        for employee in self._employee_filter_state:
            self._employee_filter_state[employee] = True
        for var in self._employee_filter_vars.values():
            var.set(True)
        self._update_filter_button_label()
        if self.current_data is not None:
            self._render_data(self.current_data)

    def _clear_employee_selection(self) -> None:
        for employee in self._employee_filter_state:
            self._employee_filter_state[employee] = False
        for var in self._employee_filter_vars.values():
            var.set(False)
        self._update_filter_button_label()
        if self.current_data is not None:
            self._render_data(self.current_data)

    def _select_all_pfs(self) -> None:
        for pf_value in self._pf_filter_state:
            self._pf_filter_state[pf_value] = True
        for var in self._pf_filter_vars.values():
            var.set(True)
        self._update_pf_filter_button_label()
        if self.current_data is not None:
            self._render_data(self.current_data)

    def _clear_pf_selection(self) -> None:
        if len(self._pf_filter_state) <= 1:
            self._select_all_pfs()
            return
        for pf_value in self._pf_filter_state:
            self._pf_filter_state[pf_value] = False
        for var in self._pf_filter_vars.values():
            var.set(False)
        self._update_pf_filter_button_label()
        if self.current_data is not None:
            self._render_data(self.current_data)

    def _on_employee_filter_menu_changed(self) -> None:
        for employee, var in self._employee_filter_vars.items():
            self._employee_filter_state[employee] = bool(var.get())
        self._update_filter_button_label()
        if self.current_data is not None:
            self._render_data(self.current_data)

    def _on_pf_filter_menu_changed(self) -> None:
        if len(self._pf_filter_state) <= 1:
            self._select_all_pfs()
            return
        for pf_value, var in self._pf_filter_vars.items():
            self._pf_filter_state[pf_value] = bool(var.get())
        self._update_pf_filter_button_label()
        if self.current_data is not None:
            self._render_data(self.current_data)

    def _update_filter_button_label(self) -> None:
        employee_names = self.current_data.employee_names if self.current_data is not None else []
        if not employee_names:
            self.filter_button_var.set("All Employees")
            return
        selected = [employee for employee in employee_names if self._employee_filter_state.get(employee, True)]
        if len(selected) == len(employee_names):
            self.filter_button_var.set("All Employees")
        elif not selected:
            self.filter_button_var.set("No Employees")
        elif len(selected) == 1:
            self.filter_button_var.set(selected[0])
        else:
            self.filter_button_var.set(f"{len(selected)} Employees")

    def _update_pf_filter_button_label(self) -> None:
        pf_values = sorted(self._pf_filter_state)
        if not pf_values:
            self.pf_filter_button_var.set("All PFs")
            return
        selected = [pf_value for pf_value in pf_values if self._pf_filter_state.get(pf_value, True)]
        if len(selected) == len(pf_values):
            self.pf_filter_button_var.set("All PFs")
        elif not selected:
            self.pf_filter_button_var.set("No PFs")
        elif len(selected) == 1:
            self.pf_filter_button_var.set(selected[0])
        else:
            self.pf_filter_button_var.set(f"{len(selected)} PFs")

    def _auto_sync_outlook_emails(self) -> None:
        if self._outlook_auto_sync_done:
            return
        self._outlook_auto_sync_done = True
        if self.current_data is not None:
            self.sync_outlook_emails(manual=False)

    def sync_outlook_emails(self, manual: bool = True) -> None:
        if self.current_data is None or self._outlook_sync_in_progress or self._cancel_event is not None:
            return

        missing_names = [
            employee
            for employee in self.current_data.employee_names
            if not self.employee_emails.get(employee, "").strip()
        ]
        if not missing_names:
            if manual:
                messagebox.showinfo("Outlook Email Sync", "All loaded employees already have email addresses saved.")
            return

        self._outlook_sync_in_progress = True
        cancel_event = threading.Event()
        self._begin_cancellable_action("Outlook email sync", cancel_event)
        self._refresh_outlook_sync_button()
        worker = threading.Thread(
            target=self._sync_outlook_emails_worker,
            args=(missing_names, manual, cancel_event),
            daemon=True,
        )
        worker.start()

    def _sync_outlook_emails_worker(self, employee_names: list[str], manual: bool, cancel_event: threading.Event) -> None:
        try:
            results = lookup_outlook_emails(employee_names, should_cancel=cancel_event.is_set)
        except OperationCancelled:
            self.after(0, lambda: self._handle_outlook_sync_cancelled(cancel_event, manual))
            return
        except Exception as exc:
            self.after(0, lambda: self._handle_outlook_sync_error(exc, manual, cancel_event))
            return
        self.after(0, lambda: self._handle_outlook_sync_success(results, employee_names, manual, cancel_event))

    def _handle_outlook_sync_cancelled(self, cancel_event: threading.Event, manual: bool) -> None:
        self._outlook_sync_in_progress = False
        self._end_cancellable_action(cancel_event)
        self._refresh_outlook_sync_button()
        if manual:
            messagebox.showinfo("Outlook Email Sync", "Outlook email sync was cancelled.")

    def _handle_outlook_sync_error(self, exc: Exception, manual: bool, cancel_event: threading.Event | None = None) -> None:
        self._outlook_sync_in_progress = False
        self._end_cancellable_action(cancel_event)
        self._refresh_outlook_sync_button()
        if manual and not isinstance(exc, OperationCancelled):
            messagebox.showerror("Outlook Email Sync", f"Could not query Outlook emails.\n\n{exc}")

    def _handle_outlook_sync_success(
        self,
        results: dict[str, str],
        employee_names: list[str],
        manual: bool,
        cancel_event: threading.Event | None = None,
    ) -> None:
        self._outlook_sync_in_progress = False
        self._end_cancellable_action(cancel_event)
        updated = 0
        for employee, email in results.items():
            if email and not self.employee_emails.get(employee, "").strip():
                self.employee_emails[employee] = email
                updated += 1
        if updated:
            self._persist_employee_emails()
            if self.current_data is not None:
                self._render_data(self.current_data)
        self._refresh_outlook_sync_button()

        typo_warnings: list[NameTypoWarning] = []
        if self.current_data is not None:
            unresolved_names = [employee for employee in employee_names if not results.get(employee, "").strip()]
            typo_warnings = [
                warning
                for warning in find_potential_name_typos(
                    unresolved_names,
                    self.current_data.employee_names,
                    self.current_data.daily_records,
                )
                if typo_warning_key(warning.employee, warning.similar_employee) not in self.ignored_name_typos
            ]

        if manual:
            missing_after = sum(1 for employee in employee_names if not self.employee_emails.get(employee, "").strip())
            messagebox.showinfo(
                "Outlook Email Sync",
                f"Matched emails: {updated}\nStill missing: {missing_after}",
            )
        if typo_warnings and not self.app_settings.disable_name_typo_notifications:
            self._show_typo_warning_dialog(typo_warnings)

    def _persist_ignored_name_typos(self) -> None:
        save_ignored_name_typos(self.config_path, self.ignored_name_typos)

    def _show_typo_warning_dialog(self, typo_warnings: list[NameTypoWarning]) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("Potential Name Typos")
        dialog.transient(self)
        dialog.geometry("780x420")

        ttk.Label(
            dialog,
            text="Possible name typo(s) were found for unresolved Outlook matches. Select any entry and choose Ignore Selected to stop warning on that inconsistency going forward.",
            wraplength=740,
            justify="left",
        ).pack(fill="x", padx=12, pady=(12, 8))

        content = ttk.Frame(dialog, padding=(12, 0, 12, 12))
        content.pack(fill="both", expand=True)
        content.columnconfigure(2, weight=1)
        content.rowconfigure(0, weight=1)

        listbox = tk.Listbox(content, selectmode="extended", exportselection=False)
        listbox.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        scrollbar = ttk.Scrollbar(content, orient="vertical", command=listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        listbox.configure(yscrollcommand=scrollbar.set)
        bind_vertical_mousewheel(listbox, listbox, units_per_notch=4)

        details = tk.Text(content, wrap="word", height=12)
        details.grid(row=0, column=2, sticky="nsew")
        bind_vertical_mousewheel(details, details, units_per_notch=4)

        def refresh_details(_event=None) -> None:
            selection = listbox.curselection()
            details.configure(state="normal")
            details.delete("1.0", tk.END)
            if selection:
                warning = typo_warnings[selection[0]]
                lines = [
                    f"{warning.employee} may match {warning.similar_employee}",
                    f"Similarity: {warning.similarity * 100:.0f}%",
                    "",
                    "Locations:",
                    *(warning.locations or ["(No locations found)"]),
                ]
                details.insert("1.0", "\n".join(lines))
            details.configure(state="disabled")

        for warning in typo_warnings:
            listbox.insert(
                tk.END,
                f"{warning.employee} -> {warning.similar_employee} ({warning.similarity * 100:.0f}% similar)",
            )
        if typo_warnings:
            listbox.selection_set(0)
        listbox.bind("<<ListboxSelect>>", refresh_details)
        refresh_details()

        buttons = ttk.Frame(dialog, padding=(12, 0, 12, 12))
        buttons.pack(fill="x")

        def ignore_selected() -> None:
            selected = listbox.curselection()
            if not selected:
                return
            for index in selected:
                warning = typo_warnings[index]
                self.ignored_name_typos.add(typo_warning_key(warning.employee, warning.similar_employee))
            self._persist_ignored_name_typos()
            dialog.destroy()

        ttk.Button(buttons, text="Ignore Selected", command=ignore_selected).pack(side="left")
        ttk.Button(buttons, text="Close", command=dialog.destroy).pack(side="right")

    def choose_sources(self) -> None:
        selected = filedialog.askopenfilenames(
            title="Select DSS workbook(s)",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if selected:
            self.load_source([Path(path) for path in selected])

    def add_sources(self) -> None:
        if self._cancel_event is not None:
            return
        selected = filedialog.askopenfilenames(
            title="Add DSS workbook(s)",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not selected:
            return
        existing = list(self.current_data.source_paths) if self.current_data is not None else []
        combined = existing + [Path(path) for path in selected]
        self.load_source(combined)

    def remove_sources(self) -> None:
        if self.current_data is None or self._cancel_event is not None:
            return
        source_paths = list(self.current_data.source_paths)
        if not source_paths:
            return

        dialog = tk.Toplevel(self)
        dialog.title("Remove DSS Workbook(s)")
        dialog.transient(self)
        dialog.geometry("720x320")

        ttk.Label(dialog, text="Select the loaded DSS workbook(s) to remove.").pack(anchor="w", padx=12, pady=(12, 8))

        list_frame = ttk.Frame(dialog, padding=(12, 0, 12, 0))
        list_frame.pack(fill="both", expand=True)
        listbox = tk.Listbox(list_frame, selectmode="extended", exportselection=False)
        listbox.pack(side="left", fill="both", expand=True)
        bind_vertical_mousewheel(listbox, listbox, units_per_notch=4)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="left", fill="y")

        for source_path in source_paths:
            listbox.insert(tk.END, str(source_path))

        buttons = ttk.Frame(dialog, padding=12)
        buttons.pack(fill="x")

        def remove_selected() -> None:
            selected_indices = listbox.curselection()
            if not selected_indices:
                dialog.destroy()
                return
            keep_paths = [path for index, path in enumerate(source_paths) if index not in selected_indices]
            dialog.destroy()
            if keep_paths:
                self.load_source(keep_paths)
            else:
                self.current_data = None
                self._hash_alerted_paths.clear()
                self.source_label.configure(text="No workbook loaded")
                self._employee_filter_state = {}
                self._refresh_filter_options()
                self._refresh_stats_summary()
                self._clear_all_views()
                self._set_loading_state(False)

        ttk.Button(buttons, text="Remove Selected", command=remove_selected).pack(side="right")
        ttk.Button(buttons, text="Cancel", command=dialog.destroy).pack(side="right", padx=(0, 8))

    def reload_source(self) -> None:
        if not self.current_data or self._is_loading:
            return
        self.load_source(self.current_data.source_paths, show_success=True)

    def load_source(self, source_paths: Path | Iterable[Path], show_success: bool = False) -> None:
        if self._cancel_event is not None:
            return
        normalized_paths = self._normalize_source_paths(source_paths)
        if not normalized_paths:
            return

        self._load_request_id += 1
        request_id = self._load_request_id
        cancel_event = threading.Event()
        self._has_partial_preview = False
        self._begin_cancellable_action("DSS load", cancel_event)
        self._set_loading_state(True, normalized_paths)

        worker = threading.Thread(
            target=self._load_source_worker,
            args=(request_id, normalized_paths, show_success, cancel_event),
            daemon=True,
        )
        worker.start()

    @staticmethod
    def _normalize_source_paths(source_paths: Path | Iterable[Path]) -> list[Path]:
        if isinstance(source_paths, Path):
            return [source_paths.expanduser().resolve()]
        return [Path(path).expanduser().resolve() for path in source_paths]

    def _set_loading_state(self, is_loading: bool, source_paths: list[Path] | None = None) -> None:
        self._is_loading = is_loading
        self.open_button.configure(state="disabled" if is_loading else "normal")
        self.add_button.configure(state="disabled" if is_loading else "normal")
        self.remove_button.configure(state="disabled" if is_loading or self.current_data is None else "normal")
        self.reload_button.configure(
            state="disabled" if is_loading or self.current_data is None else "normal"
        )
        if is_loading:
            if source_paths and len(source_paths) == 1:
                target = source_paths[0].name
            elif source_paths:
                target = f"{len(source_paths)} DSS files"
            else:
                target = "selected DSS files"
            self.loading_label.configure(text=f"Loading {target}...")
            self.stats_label.configure(text="Reading workbook data in the background...")
            self.progress_var.set(0.0)
        else:
            self.loading_label.configure(text="")
            self.progress_var.set(0.0)
        self._refresh_outlook_sync_button()

    def _clear_all_views(self) -> None:
        self.daily_table.set_rows([])
        self.weekly_rollup_table.set_rows([])
        self.combined_weekly_summary_table.set_rows([])
        self.week_totals_table.set_rows([])
        self.error_report_table.set_rows([])
        self.parse_warnings_table.set_rows([])
        self.workbook_health_table.set_rows([])
        self.audit_data_trail_table.set_rows([])
        self.email_drafts_frame.preview_table.set_rows([])
        self.email_drafts_frame.summary_label.configure(text="Load DSS data to preview employee email drafts.")
        self.email_drafts_frame.set_week_options([])
        self.employee_editor.set_names([], self.employee_emails)
        self.notes_frame.set_data([], self.employee_notes)
        self.groups_frame.set_data([], self.employee_groups, self.employee_emails)

    def _load_source_worker(
        self,
        request_id: int,
        source_paths: list[Path],
        show_success: bool,
        cancel_event: threading.Event,
    ) -> None:
        def report_progress(progress_fraction: float, message: str) -> None:
            self.after(0, lambda: self._update_progress(request_id, progress_fraction, message))

        def report_partial(tracker_data: TrackerData, message: str) -> None:
            self.after(0, lambda: self._handle_partial_load_update(request_id, tracker_data, message))

        try:
            tracker_data = load_tracker_data(
                source_paths,
                previous_data=self.current_data,
                progress_callback=report_progress,
                partial_callback=report_partial,
                cache_dir=self.cache_dir,
                should_cancel=cancel_event.is_set,
            )
        except OperationCancelled:
            self.after(0, lambda: self._handle_load_cancelled(request_id, cancel_event))
            return
        except Exception as exc:
            self.after(0, lambda: self._handle_load_error(request_id, exc, cancel_event))
            return
        self.after(0, lambda: self._handle_load_success(request_id, tracker_data, show_success, cancel_event))

    def _update_progress(self, request_id: int, progress_fraction: float, message: str) -> None:
        if request_id != self._load_request_id:
            return
        percentage = round(min(max(progress_fraction, 0.0), 1.0) * 100, 1)
        self.progress_var.set(percentage)
        self.loading_label.configure(text=f"{percentage:.1f}%")
        self.stats_label.configure(text=message)

    def _handle_load_cancelled(self, request_id: int, cancel_event: threading.Event | None = None) -> None:
        if request_id != self._load_request_id:
            return
        self._set_loading_state(False)
        self._end_cancellable_action(cancel_event)
        if self._has_partial_preview and self.current_data is not None:
            self.stats_label.configure(text="Load cancelled. Showing partially loaded recent DSS data.")
        else:
            self.stats_label.configure(text="Load cancelled.")

    def _handle_load_error(self, request_id: int, exc: Exception, cancel_event: threading.Event | None = None) -> None:
        if request_id != self._load_request_id:
            return
        self._set_loading_state(False)
        self._end_cancellable_action(cancel_event)
        messagebox.showerror("DSS Hours Tracker", f"Failed to open workbook.\n\n{exc}")

    def _handle_partial_load_update(self, request_id: int, tracker_data: TrackerData, message: str) -> None:
        if request_id != self._load_request_id:
            return
        self._has_partial_preview = True
        self.current_data = tracker_data
        self._refresh_filter_options()
        self._refresh_outlook_sync_button()
        if len(tracker_data.source_paths) == 1:
            source_text = str(tracker_data.source_paths[0])
        else:
            source_text = f"{len(tracker_data.source_paths)} DSS workbooks loading"
        self.source_label.configure(text=source_text)
        self._render_data(tracker_data)
        self.stats_label.configure(text=message)

    def _handle_load_success(
        self,
        request_id: int,
        tracker_data: TrackerData,
        show_success: bool,
        cancel_event: threading.Event | None = None,
    ) -> None:
        if request_id != self._load_request_id:
            return

        self._has_partial_preview = False
        self.current_data = tracker_data
        self._hash_alerted_paths.clear()
        self._set_loading_state(False)
        self._end_cancellable_action(cancel_event)
        self._refresh_filter_options()
        self._refresh_outlook_sync_button()
        if len(tracker_data.source_paths) == 1:
            source_text = str(tracker_data.source_paths[0])
        else:
            source_text = f"{len(tracker_data.source_paths)} DSS workbooks loaded"
        self.source_label.configure(text=source_text)
        self._refresh_stats_summary()
        self._render_data(tracker_data)
        self._schedule_hash_monitor()
        if show_success:
            summary_lines = [
                f"Reloaded: {len(tracker_data.reloaded_paths)}",
                f"Unchanged: {len(tracker_data.reused_paths)}",
            ]
            if tracker_data.reloaded_paths:
                summary_lines.append("")
                summary_lines.append("Reprocessed files:")
                summary_lines.extend(str(path) for path in tracker_data.reloaded_paths)
            messagebox.showinfo(
                "DSS Hours Tracker",
                "\n".join(summary_lines),
            )

    def _render_data(self, tracker_data: TrackerData) -> None:
        filter_selection = self._current_filter_selection()
        selected_pfs = self._current_pf_selection()
        filtered_daily_records = [
            record
            for record in tracker_data.daily_records
            if extract_pf_identifier(record.source_file) in selected_pfs
        ]
        filtered_employee_names = sorted({record.employee for record in filtered_daily_records})
        allowed_employees = filter_employee_names(filtered_employee_names, filter_selection, self.employee_groups)
        filtered_daily_records = [record for record in filtered_daily_records if record.employee in allowed_employees]
        filtered_weekly_rollup = [
            record for record in tracker_data.weekly_rollup
            if (
                extract_pf_identifier(record.source_file) in selected_pfs
                and (record.row_type == "Crew Total" or record.employee in allowed_employees)
            )
        ]
        filtered_combined_summary = [
            record for record in tracker_data.combined_weekly_summary if record.employee in allowed_employees
        ]
        filtered_week_totals = build_week_totals(filtered_combined_summary) if filtered_combined_summary else []
        filtered_source_files = {record.source_file for record in filtered_daily_records}

        self.daily_table.set_rows(
            [
                (
                    record.source_file,
                    record.work_date.isoformat(),
                    record.source_sheet,
                    record.employee,
                    fmt_hours(record.st),
                    fmt_hours(record.ot),
                    fmt_hours(record.dt),
                    fmt_hours(record.total),
                    fmt_hours(expanded_hours(record.st, record.ot, record.dt)),
                    record.source_ranges,
                )
                for record in filtered_daily_records
            ]
        )

        self.employee_editor.set_names(tracker_data.employee_names, self.employee_emails)
        self.notes_frame.set_data(tracker_data.employee_names, self.employee_notes)
        self.groups_frame.set_data(tracker_data.employee_names, self.employee_groups, self.employee_emails)

        weekly_rollup_rows: list[tuple[str, ...]] = []
        weekly_rollup_tags: list[tuple[str, ...]] = []
        for record in filtered_weekly_rollup:
            weekly_rollup_rows.append(
                (
                    record.source_file,
                    record.week_start.isoformat(),
                    record.week_end.isoformat(),
                    record.employee,
                    fmt_hours(record.st),
                    fmt_hours(record.ot),
                    fmt_hours(record.dt),
                    fmt_hours(record.total),
                    fmt_hours(expanded_hours(record.st, record.ot, record.dt)),
                    record.row_type,
                )
            )
            if record.row_type == "Crew Total":
                weekly_rollup_tags.append(("crew_total",))
            else:
                weekly_rollup_tags.append(self._threshold_tags(record.st, record.ot, record.dt))
        self.weekly_rollup_table.set_rows(weekly_rollup_rows, weekly_rollup_tags)

        combined_summary_rows: list[tuple[str, ...]] = []
        combined_summary_tags: list[tuple[str, ...]] = []
        for record in filtered_combined_summary:
            combined_summary_rows.append(
                (
                    record.week_start.isoformat(),
                    record.week_end.isoformat(),
                    record.employee,
                    fmt_hours(record.st),
                    fmt_hours(record.ot),
                    fmt_hours(record.dt),
                    fmt_hours(record.total),
                    fmt_hours(expanded_hours(record.st, record.ot, record.dt)),
                )
            )
            combined_summary_tags.append(self._threshold_tags(record.st, record.ot, record.dt))
        self.combined_weekly_summary_table.set_rows(combined_summary_rows, combined_summary_tags)

        self.week_totals_table.set_rows(
            [
                (
                    record.week_start.isoformat(),
                    record.week_end.isoformat(),
                    fmt_hours(record.st),
                    fmt_hours(record.ot),
                    fmt_hours(record.dt),
                    fmt_hours(record.total),
                    fmt_hours(expanded_hours(record.st, record.ot, record.dt)),
                )
                for record in filtered_week_totals
            ]
        )
        error_findings = build_error_findings(filtered_daily_records, self._active_profile())
        self.error_report_table.set_rows(
            [
                (
                    finding.employee,
                    finding.week_start.isoformat(),
                    finding.week_end.isoformat(),
                    f"{finding.hour_type} > {fmt_hours(finding.threshold)}",
                    finding.trigger_date.isoformat(),
                    fmt_hours(finding.actual_total),
                    fmt_hours(finding.threshold),
                    fmt_hours(finding.delta),
                    fmt_hours(finding.trigger_day_st),
                    fmt_hours(finding.trigger_day_ot),
                    fmt_hours(finding.trigger_day_dt),
                    finding.source_files,
                    finding.reason,
                    finding.breakdown,
                )
                for finding in error_findings
            ],
            tags=[("alert",) for _ in error_findings],
        )
        self.parse_warnings_table.set_rows(
            [
                (
                    warning.source_file,
                    warning.source_sheet,
                    warning.work_date,
                    warning.issue,
                    warning.details,
                )
                for warning in tracker_data.parse_warnings
                if warning.source_file in filtered_source_files or not filtered_daily_records
            ],
            tags=[
                ("alert",)
                for warning in tracker_data.parse_warnings
                if warning.source_file in filtered_source_files or not filtered_daily_records
            ],
        )
        self.workbook_health_table.set_rows(
            [
                (
                    item.source_file,
                    item.status,
                    item.details,
                )
                for item in tracker_data.workbook_health
            ],
            tags=[("alert",) if item.status.lower() == "warning" else tuple() for item in tracker_data.workbook_health],
        )
        self.audit_data_trail_table.set_rows(
            [
                (
                    record.source_file,
                    record.work_date.isoformat(),
                    record.source_sheet,
                    record.employee,
                    fmt_hours(record.st),
                    fmt_hours(record.ot),
                    fmt_hours(record.dt),
                    fmt_hours(record.total),
                    fmt_hours(expanded_hours(record.st, record.ot, record.dt)),
                    record.source_ranges,
                    f"Derived from parsed block totals for {record.employee} on {record.source_sheet}.",
                )
                for record in filtered_daily_records
            ]
        )
        self._refresh_email_preview(tracker_data, allowed_employees)

    def _refresh_email_preview(self, tracker_data: TrackerData, allowed_employees: set[str] | None = None) -> None:
        filtered_daily_records = [
            record for record in tracker_data.daily_records
            if allowed_employees is None or record.employee in allowed_employees
        ]
        week_options = collect_week_ranges(filtered_daily_records)
        previous_selection = self.email_drafts_frame.week_var.get()
        self.email_drafts_frame.set_week_options(week_options)
        if previous_selection in self.email_drafts_frame.week_combo.cget("values"):
            self.email_drafts_frame.week_var.set(previous_selection)

        week_start = self.email_drafts_frame.selected_week_start()
        if week_start is None:
            self.email_drafts_frame.summary_label.configure(text="Load DSS data to preview employee email drafts.")
            self.email_drafts_frame.preview_table.set_rows([])
            return

        requests = build_email_draft_requests(filtered_daily_records, self.employee_emails, week_start)
        preview_rows: list[tuple[str, ...]] = []
        for request in requests:
            st_total = round(sum(record.st for record in request.records), 2)
            ot_total = round(sum(record.ot for record in request.records), 2)
            dt_total = round(sum(record.dt for record in request.records), 2)
            preview_rows.append(
                (
                    request.employee,
                    request.email or "(missing)",
                    str(len(request.records)),
                    fmt_hours(st_total),
                    fmt_hours(ot_total),
                    fmt_hours(dt_total),
                    fmt_hours(st_total + ot_total + dt_total),
                    fmt_hours(expanded_hours(st_total, ot_total, dt_total)),
                )
            )
        self.email_drafts_frame.preview_table.set_rows(preview_rows)
        missing = sum(1 for request in requests if not request.email)
        week_end = week_start + timedelta(days=6)
        self.email_drafts_frame.summary_label.configure(
            text=(
                f"Previewing {len(requests)} employee draft(s) for {format_week_label(week_start, week_end)}. "
                f"Missing email addresses: {missing}."
            )
        )

    def create_email_drafts(self) -> None:
        if self.current_data is None:
            messagebox.showerror("Email Drafts", "Load DSS data before creating draft emails.")
            return

        week_start = self.email_drafts_frame.selected_week_start()
        if week_start is None:
            messagebox.showerror("Email Drafts", "No week is available for draft creation.")
            return

        allowed_employees = filter_employee_names(
            self.current_data.employee_names,
            self._current_filter_selection(),
            self.employee_groups,
        )
        filtered_daily_records = [
            record for record in self.current_data.daily_records if record.employee in allowed_employees
        ]
        requests = build_email_draft_requests(filtered_daily_records, self.employee_emails, week_start)
        selected_employees = self.email_drafts_frame.selected_employees()
        if selected_employees:
            requests = [request for request in requests if request.employee in selected_employees]
        if not requests:
            messagebox.showerror("Email Drafts", "No employee records were found for the selected week.")
            return

        try:
            created, skipped = create_outlook_drafts(
                requests,
                self.email_drafts_frame.get_subject_template(),
                self.email_drafts_frame.get_body_template(),
            )
        except Exception as exc:
            messagebox.showerror("Email Drafts", f"Could not create Outlook drafts.\n\n{exc}")
            return

        summary_lines = [f"Created drafts: {created}", f"Skipped (missing email): {len(skipped)}"]
        if skipped:
            summary_lines.append("")
            summary_lines.append("Missing email addresses:")
            summary_lines.extend(skipped)
        messagebox.showinfo("Email Drafts", "\n".join(summary_lines))

    def _schedule_hash_monitor(self) -> None:
        self._hash_monitor_token += 1
        token = self._hash_monitor_token
        self.after(self.hash_poll_interval_ms, lambda: self._run_hash_monitor(token))

    def _run_hash_monitor(self, token: int) -> None:
        if token != self._hash_monitor_token:
            return
        if self.current_data is None or self._is_loading or not self.current_data.source_paths:
            self._schedule_hash_monitor()
            return

        source_paths = list(self.current_data.source_paths)
        expected_hashes = dict(self.current_data.file_hashes)
        worker = threading.Thread(
            target=self._hash_monitor_worker,
            args=(token, source_paths, expected_hashes),
            daemon=True,
        )
        worker.start()

    def _hash_monitor_worker(self, token: int, source_paths: list[Path], expected_hashes: dict[Path, str]) -> None:
        changed_paths: list[Path] = []
        for source_path in source_paths:
            try:
                current_hash = compute_workbook_content_hash(read_source_bytes(source_path))
            except Exception:
                continue
            if expected_hashes.get(source_path) != current_hash:
                changed_paths.append(source_path)
        self.after(0, lambda: self._handle_hash_monitor_result(token, changed_paths))

    def _handle_hash_monitor_result(self, token: int, changed_paths: list[Path]) -> None:
        if token != self._hash_monitor_token:
            return
        unseen_changes = [path for path in changed_paths if path not in self._hash_alerted_paths]
        if unseen_changes:
            self._hash_alerted_paths.update(unseen_changes)
            changed_names = ", ".join(path.name for path in unseen_changes)
            self.loading_label.configure(text="Changed")
            self.stats_label.configure(
                text=f"DSS changed since last view: {changed_names}. Press Update View to refresh."
            )
            messagebox.showinfo(
                "DSS Changed",
                "One or more loaded DSS workbooks changed since the last inspected state.\n\n"
                "Press 'Update View' to refresh the summaries.",
            )
        self._schedule_hash_monitor()

    def _on_email_week_changed(self, _event=None) -> None:
        if self.current_data is not None:
            self._refresh_email_preview(self.current_data)

    def _threshold_tags(self, st: float, ot: float, dt: float) -> tuple[str, ...]:
        if is_alert_triggered(st, ot, dt, self._active_profile()):
            return ("alert",)
        return tuple()


def main() -> int:
    parser = argparse.ArgumentParser(description="Desktop viewer for DSS labour-hour summaries.")
    parser.add_argument("source", nargs="*", help="Optional DSS workbook(s) to open on launch.")
    args = parser.parse_args()

    initial_source = [Path(path).expanduser().resolve() for path in args.source] if args.source else None
    app = DssHoursTrackerApp(initial_source=initial_source)
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

