from __future__ import annotations

"""
Fast unit tests (no RUN_SLOW_TESTS required).

Workbook / load_tracker_data / cache integration tests live in
`test_dss_hours_tracker_integration.py` and run when RUN_SLOW_TESTS=1.
"""

import os
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

from dss_hours_tracker import (
    AppSettings,
    DEFAULT_UI_THEME,
    binding_sequence_from_keypress_event,
    az2_revision_matches_sheet_name,
    build_bug_report_html,
    checksum_for_asset_name,
    choose_release_checksum_asset,
    choose_release_installer_asset,
    build_permission_denied_message,
    is_newer_version,
    normalize_release_version,
    parse_checksum_manifest,
    parse_latest_release_payload,
    compute_bytes_hash,
    combine_sheet_hashes,
    load_ignored_name_typos,
    pf_numbers_for_records,
    save_ignored_name_typos,
    typo_warning_key,
    format_email_subject,
    default_formatting_profiles,
    extract_pf_identifier,
    FilterSelection,
    find_open_excel_workbook,
    filter_employee_names,
    FormattingProfile,
    get_app_root,
    is_alert_triggered,
    is_allowed_quickload_cancel_hotkey,
    is_path_like_table_column,
    is_unmetered_wifi_profile,
    load_email_templates,
    load_formatting_profiles,
    load_app_settings,
    normalize_ui_hex_color,
    parse_ui_theme_payload,
    DailyRecord,
    build_tracker_data_with_status,
    tracker_data_invalidated_for_cache_clear,
    load_employee_emails,
    load_employee_groups,
    load_table_layouts,
    normalize_person_name,
    normalize_quickload_cancel_hotkey,
    normalize_windows_path,
    parse_sheet_revision,
    parse_threshold_value,
    daily_rollup_sort_key,
    weekly_rollup_sort_key,
    remove_config_keys,
    save_employee_emails,
    save_employee_groups,
    save_formatting_profiles,
    save_email_templates,
    save_app_settings,
    save_table_layout,
    select_preferred_dated_sheets,
)
from test_dss_hours_tracker_fixtures import DssHoursTrackerFixtures


class DssHoursTrackerTests(DssHoursTrackerFixtures):
    def test_daily_rollup_sort_key_keeps_sections_grouped(self) -> None:
        rows = [
            ("PF26024-2 Electrical", "2026-04-07", "Whole Crew", "255.5", "30", "60", "345.5", "420.5", "Crew Total"),
            ("PF26024-1 Instrumentation", "2026-04-07", "Grady Redden", "50", "0", "6", "56", "62", "Employee"),
            ("PF26024-2 Electrical", "2026-04-07", "Dexter Olshewski", "10", "0", "4", "14", "18", "Employee"),
            ("PF26024-1 Instrumentation", "2026-04-07", "Chuck Ehr", "50", "0", "6", "56", "62", "Employee"),
            ("PF26024-1 Instrumentation", "2026-04-07", "Whole Crew", "288", "0", "32.5", "320.5", "353", "Crew Total"),
        ]
        sorted_rows = sorted(rows, key=lambda row: daily_rollup_sort_key("work_date", row, True))
        self.assertEqual(sorted_rows[0][0], "PF26024-1 Instrumentation")
        self.assertEqual(sorted_rows[0][2], "Chuck Ehr")
        self.assertEqual(sorted_rows[1][0], "PF26024-1 Instrumentation")
        self.assertEqual(sorted_rows[2][2], "Whole Crew")
        self.assertEqual(sorted_rows[3][0], "PF26024-2 Electrical")
        self.assertEqual(sorted_rows[3][2], "Dexter Olshewski")
        self.assertEqual(sorted_rows[4][2], "Whole Crew")

    def test_weekly_rollup_sort_key_keeps_sections_grouped(self) -> None:
        rows = [
            ("PF26024-2 Electrical", "2026-04-27", "2026-05-03", "Whole Crew", "255.5", "30", "60", "345.5", "420.5", "Crew Total"),
            ("PF26024-1 Instrumentation", "2026-04-27", "2026-05-03", "Grady Redden", "50", "0", "6", "56", "62", "Employee"),
            ("PF26024-2 Electrical", "2026-04-27", "2026-05-03", "Dexter Olshewski", "10", "0", "4", "14", "18", "Employee"),
            ("PF26024-1 Instrumentation", "2026-04-27", "2026-05-03", "Chuck Ehr", "50", "0", "6", "56", "62", "Employee"),
            ("PF26024-1 Instrumentation", "2026-04-27", "2026-05-03", "Whole Crew", "288", "0", "32.5", "320.5", "353", "Crew Total"),
        ]
        sorted_rows = sorted(rows, key=lambda row: weekly_rollup_sort_key("week_start", row, True))
        self.assertEqual(sorted_rows[0][0], "PF26024-1 Instrumentation")
        self.assertEqual(sorted_rows[0][3], "Chuck Ehr")
        self.assertEqual(sorted_rows[1][0], "PF26024-1 Instrumentation")
        self.assertEqual(sorted_rows[2][3], "Whole Crew")
        self.assertEqual(sorted_rows[3][0], "PF26024-2 Electrical")
        self.assertEqual(sorted_rows[3][3], "Dexter Olshewski")
        self.assertEqual(sorted_rows[4][3], "Whole Crew")

    def test_parse_sheet_revision_handles_common_patterns(self) -> None:
        self.assertEqual(parse_sheet_revision("2026-04-07"), 0)
        self.assertEqual(parse_sheet_revision("2026-04-07 R1"), 1)
        self.assertEqual(parse_sheet_revision("2026-04-07 r 2"), 2)
        self.assertEqual(parse_sheet_revision("2026-04-07 rev 3"), 3)
        self.assertEqual(parse_sheet_revision("2026-04-07 Revision_4"), 4)
        self.assertEqual(parse_sheet_revision("2026-04-07 r-1"), 1)
        self.assertEqual(parse_sheet_revision("2026-04-07 r.1"), 1)
        self.assertEqual(parse_sheet_revision("2026-04-07 rev 12"), 12)
        self.assertEqual(parse_sheet_revision("2026-04-26 rv1"), 1)
        self.assertEqual(parse_sheet_revision("2026-04-26 RV2"), 2)

    def test_az2_revision_matches_sheet_name(self) -> None:
        ok, _ = az2_revision_matches_sheet_name("2026-04-07 R2", 2)
        self.assertTrue(ok)
        ok, msg = az2_revision_matches_sheet_name("2026-04-07 R2", 1)
        self.assertFalse(ok)
        self.assertIn("revision level 1", msg or "")
        ok, msg = az2_revision_matches_sheet_name("2026-04-07", "")
        self.assertTrue(ok)
        ok, msg = az2_revision_matches_sheet_name("2026-04-07 R1", "")
        self.assertFalse(ok)
        self.assertIn("blank", (msg or "").lower())

    def test_combine_sheet_hashes_stable(self) -> None:
        a = {"2026-04-07": "aa", "2026-04-08": "bb"}
        self.assertEqual(combine_sheet_hashes(a), combine_sheet_hashes(dict(reversed(list(a.items())))))

    def test_extract_pf_identifier_prefers_pf_token(self) -> None:
        self.assertEqual(
            extract_pf_identifier("PF26024-2 Electrical Tech Maintenance Ongoing DSS.xlsx"),
            "PF26024-2",
        )
        self.assertEqual(
            extract_pf_identifier("ja tech pf26024-7 phase dss.xlsx"),
            "PF26024-7",
        )

    def test_extract_pf_identifier_falls_back_to_stem(self) -> None:
        self.assertEqual(extract_pf_identifier("Electrical Tech Maintenance.xlsx"), "Electrical Tech Maintenance")

    def test_select_preferred_dated_sheets_prefers_highest_revision(self) -> None:
        selected = select_preferred_dated_sheets(
            ["2026-04-07", "2026-04-07 R1", "2026-04-08", "2026-04-08 rev 2", "notes"]
        )
        self.assertEqual(
            selected,
            [
                (date(2026, 4, 7), "2026-04-07 R1"),
                (date(2026, 4, 8), "2026-04-08 rev 2"),
            ],
        )

    def test_permission_message_mentions_onedrive_guidance(self) -> None:
        message = build_permission_denied_message(
            Path(r"C:\Users\Test\OneDrive - JA Tech\SharePoint - PF26024\Outage DSS.xlsx")
        )
        self.assertIn("OneDrive / SharePoint", message)
        self.assertIn("Always keep on this device", message)
        self.assertIn("Update View", message)

    def test_normalize_windows_path(self) -> None:
        self.assertEqual(
            normalize_windows_path(r"C:/Users/Test/OneDrive/File.xlsx\\"),
            r"c:\users\test\onedrive\file.xlsx",
        )

    def test_normalize_person_name(self) -> None:
        self.assertEqual(normalize_person_name("  Alice   Smith "), "alice smith")

    def test_find_open_excel_workbook_matches_by_full_path_or_unique_name(self) -> None:
        class FakeWorkbook:
            def __init__(self, full_name: str, name: str):
                self.FullName = full_name
                self.Name = name

        class FakeExcel:
            def __init__(self, workbooks):
                self.Workbooks = workbooks

        exact = FakeWorkbook(
            r"C:\Users\Test\OneDrive - JA Tech\SharePoint - PF26024\Outage DSS.xlsx",
            "Outage DSS.xlsx",
        )
        result = find_open_excel_workbook(FakeExcel([exact]), Path(exact.FullName))
        self.assertIs(result, exact)

        by_name = FakeWorkbook(r"C:\Other\Folder\Outage DSS.xlsx", "Outage DSS.xlsx")
        result = find_open_excel_workbook(FakeExcel([by_name]), Path(r"C:\Missing\Outage DSS.xlsx"))
        self.assertIs(result, by_name)

    def test_normalize_release_version_and_compare(self) -> None:
        self.assertEqual(normalize_release_version('v0.2.1'), '0.2.1')
        self.assertTrue(is_newer_version('0.2.1', '0.2.0'))
        self.assertFalse(is_newer_version('0.2.0', '0.2.1'))
        self.assertFalse(is_newer_version('0.2.0', '0.2.0'))

    def test_parse_latest_release_payload(self) -> None:
        payload = {
            'tag_name': 'v0.2.1',
            'name': 'Release 0.2.1',
            'html_url': 'https://github.com/LochlanRoss/DSS-Viewer/releases/tag/v0.2.1',
            'published_at': '2026-05-01T00:00:00Z',
            'body': 'Notes',
            'assets': [
                {'name': 'DSSViewerSetup.exe', 'browser_download_url': 'https://example.invalid/setup.exe', 'size': 1234, 'content_type': 'application/octet-stream'},
                {'name': 'checksums.txt', 'browser_download_url': 'https://example.invalid/checksums.txt'},
            ],
        }
        info = parse_latest_release_payload(payload)
        self.assertEqual(info['version'], '0.2.1')
        self.assertEqual(info['tag_name'], 'v0.2.1')
        self.assertEqual(info['asset_names'], ['DSSViewerSetup.exe', 'checksums.txt'])
        self.assertEqual(info['assets'][0]['download_url'], 'https://example.invalid/setup.exe')

    def test_release_asset_selection_and_checksum_parsing(self) -> None:
        release_info = {
            'assets': [
                {'name': 'notes.zip', 'download_url': 'https://example.invalid/notes.zip'},
                {'name': 'DSSViewerSetup.exe', 'download_url': 'https://example.invalid/setup.exe'},
                {'name': 'checksums.txt', 'download_url': 'https://example.invalid/checksums.txt'},
            ]
        }
        installer = choose_release_installer_asset(release_info)
        checksum_asset = choose_release_checksum_asset(release_info)
        manifest = (
            'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa *DSSViewerSetup.exe\n'
            'bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb  other.msi\n'
        )
        self.assertEqual(installer['name'], 'DSSViewerSetup.exe')
        self.assertEqual(checksum_asset['name'], 'checksums.txt')
        self.assertEqual(
            checksum_for_asset_name(manifest, 'DSSViewerSetup.exe'),
            'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa',
        )
        self.assertIn('other.msi', parse_checksum_manifest(manifest))

    def test_is_unmetered_wifi_profile(self) -> None:
        self.assertTrue(
            is_unmetered_wifi_profile(
                {
                    'supported': True,
                    'connected': True,
                    'is_wlan': True,
                    'network_cost_type': 'Unrestricted',
                    'roaming': False,
                    'over_data_limit': False,
                    'approaching_data_limit': False,
                    'background_restricted': False,
                }
            )
        )
        self.assertFalse(is_unmetered_wifi_profile({'supported': True, 'connected': True, 'is_wlan': False, 'network_cost_type': 'Unrestricted'}))

    def test_compute_bytes_hash_is_stable(self) -> None:
        payload = b"example workbook bytes"
        self.assertEqual(compute_bytes_hash(payload), compute_bytes_hash(payload))
        self.assertNotEqual(compute_bytes_hash(payload), compute_bytes_hash(payload + b"!"))

    def test_parse_threshold_value_allows_blank(self) -> None:
        self.assertIsNone(parse_threshold_value(""))
        self.assertEqual(parse_threshold_value("40"), 40.0)

    def test_is_alert_triggered_uses_configured_thresholds(self) -> None:
        profile = FormattingProfile(name="Job A", st_threshold=40.0, ot_threshold=10.0, daily_st_threshold=8.0, max_hours_per_day=12.0)
        self.assertFalse(is_alert_triggered(40.0, 10.0, 2.0, profile))
        self.assertTrue(is_alert_triggered(41.0, 0.0, 0.0, profile))
        self.assertTrue(is_alert_triggered(0.0, 11.0, 0.0, profile))

    def test_formatting_profiles_round_trip(self) -> None:
        with self.workspace_json("formatting") as path:
            profiles = default_formatting_profiles()
            profiles["Job B"] = FormattingProfile(
                name="Job B",
                st_threshold=36.0,
                ot_threshold=8.0,
                daily_st_threshold=9.0,
                max_hours_per_day=12.0,
            )
            save_formatting_profiles(path, profiles, "Job B")

            loaded_profiles, current_profile = load_formatting_profiles(path)

            self.assertEqual(current_profile, "Job B")
            self.assertEqual(loaded_profiles["Job B"].st_threshold, 36.0)
            self.assertEqual(loaded_profiles["Job B"].ot_threshold, 8.0)
            self.assertEqual(loaded_profiles["Job B"].daily_st_threshold, 9.0)
            self.assertEqual(loaded_profiles["Job B"].max_hours_per_day, 12.0)

    def test_employee_emails_round_trip(self) -> None:
        with self.workspace_json("emails") as path:
            save_employee_emails(path, {"Alice Smith": "alice@example.com", "Bob Jones": "bob@example.com"})
            loaded = load_employee_emails(path)
            self.assertEqual(loaded["Alice Smith"], "alice@example.com")
            self.assertEqual(loaded["Bob Jones"], "bob@example.com")

    def test_employee_groups_round_trip(self) -> None:
        with self.workspace_json("groups") as path:
            save_employee_groups(path, {"Crew A": ["Alice Smith", "Bob Jones"], "Crew B": ["Charlie West"]})
            loaded = load_employee_groups(path)
            self.assertEqual(loaded["Crew A"], ["Alice Smith", "Bob Jones"])
            self.assertEqual(loaded["Crew B"], ["Charlie West"])

    def test_table_layout_round_trip(self) -> None:
        with self.workspace_json("table_layout") as path:
            save_table_layout(
                path,
                "weekly_summary",
                ["week_start", "employee", "st"],
                {"week_start": 120, "employee": 240, "st": 90},
                sort_column="week_start",
                sort_descending=True,
            )

            layouts = load_table_layouts(path)

            self.assertEqual(layouts["weekly_summary"]["visible_columns"], ["week_start", "employee", "st"])
            self.assertEqual(layouts["weekly_summary"]["column_widths"]["employee"], 240)
            self.assertEqual(layouts["weekly_summary"]["sort_column"], "week_start")
            self.assertTrue(layouts["weekly_summary"]["sort_descending"])
            self.assertEqual(layouts["weekly_summary"].get("column_filters", {}), {})

    def test_table_layout_column_filters_round_trip(self) -> None:
        with self.workspace_json("table_layout_cf") as path:
            save_table_layout(
                path,
                "t_filter",
                ["employee", "st"],
                {"employee": 200, "st": 80},
                sort_column="employee",
                sort_descending=False,
                column_filters={"employee": {"Alice Smith", "Bob Jones"}},
            )
            layouts = load_table_layouts(path)
            self.assertEqual(layouts["t_filter"]["column_filters"], {"employee": {"Alice Smith", "Bob Jones"}})

    def test_app_settings_round_trip(self) -> None:
        with self.workspace_json("app_settings") as path:
            save_app_settings(
                path,
                AppSettings(
                    disable_name_typo_notifications=True,
                    hash_poll_minutes=12,
                    show_daily_raw_tab=False,
                    quickload_last_sources_enabled=False,
                    quickload_cancel_hotkey="<F9>",
                    auto_update_check_enabled=False,
                    auto_download_updates_on_unmetered_wifi=False,
                ),
            )

            loaded = load_app_settings(path)

            self.assertTrue(loaded.disable_name_typo_notifications)
            self.assertEqual(loaded.hash_poll_minutes, 12)
            self.assertFalse(loaded.show_daily_raw_tab)
            self.assertFalse(loaded.quickload_last_sources_enabled)
            self.assertEqual(loaded.quickload_cancel_hotkey, "<F9>")
            self.assertFalse(loaded.auto_update_check_enabled)
            self.assertFalse(loaded.auto_download_updates_on_unmetered_wifi)
            self.assertEqual(loaded.ui_theme, DEFAULT_UI_THEME)

    def test_normalize_ui_hex_color(self) -> None:
        self.assertEqual(normalize_ui_hex_color("#abc"), "#aabbcc")
        self.assertEqual(normalize_ui_hex_color("#aABBcc"), "#aabbcc")
        self.assertIsNone(normalize_ui_hex_color("red"))
        self.assertIsNone(normalize_ui_hex_color("#gggggg"))

    def test_parse_ui_theme_payload_partial(self) -> None:
        merged = parse_ui_theme_payload({"alert_row_background": "#001122", "tooltip_foreground": "not-a-colour"})
        self.assertEqual(merged.alert_row_background, "#001122")
        self.assertEqual(merged.tooltip_foreground, DEFAULT_UI_THEME.tooltip_foreground)

    def test_parse_ui_theme_payload_includes_chrome_fields(self) -> None:
        merged = parse_ui_theme_payload({})
        self.assertEqual(merged.table_background, DEFAULT_UI_THEME.table_background)
        self.assertEqual(merged.content_chrome_background, DEFAULT_UI_THEME.content_chrome_background)
        self.assertEqual(merged.top_toolbar_background, DEFAULT_UI_THEME.top_toolbar_background)

    def test_tracker_data_invalidated_for_cache_clear(self) -> None:
        p = Path("fixture.xlsx")
        rec = DailyRecord(
            source_path=p,
            source_file="fixture.xlsx",
            work_date=date(2026, 1, 2),
            source_sheet="2026-01-02",
            employee="Test",
            st=1.0,
            ot=0.0,
            dt=0.0,
            source_ranges="",
        )
        data = build_tracker_data_with_status(
            [p],
            {p: "hash1"},
            [p],
            [],
            [rec],
            cache_status_by_path={p: "Memory Hit"},
        )
        cleared = tracker_data_invalidated_for_cache_clear(data)
        self.assertEqual(cleared.file_hashes, {})
        self.assertEqual(cleared.reused_paths, [])
        self.assertEqual(cleared.cache_status_by_path.get(p), "Miss")

    def test_discover_app_version_env_override(self) -> None:
        import dss_hours_tracker as m

        with mock.patch.dict(os.environ, {"DSS_APP_VERSION": "9.8.7"}):
            self.assertEqual(m.discover_app_version(), "9.8.7")

    def test_discover_app_version_frozen_bundle_file(self) -> None:
        import dss_hours_tracker as m

        with tempfile.TemporaryDirectory() as tdir:
            Path(tdir, "dss_app_version.txt").write_text("2.3.4", encoding="utf-8")
            with mock.patch.dict(os.environ, {"DSS_APP_VERSION": ""}):
                with mock.patch.object(m.sys, "frozen", True, create=True), mock.patch.object(
                    m.sys, "_MEIPASS", tdir, create=True
                ):
                    self.assertEqual(m.discover_app_version(), "2.3.4")

    def test_binding_sequence_from_keypress_event(self) -> None:
        class KeyEvt:
            __slots__ = ("keysym", "state")

            def __init__(self, keysym: str, state: int = 0):
                self.keysym = keysym
                self.state = state

        self.assertEqual(binding_sequence_from_keypress_event(KeyEvt("Escape")), "<Escape>")
        self.assertEqual(binding_sequence_from_keypress_event(KeyEvt("F9")), "<F9>")
        self.assertEqual(binding_sequence_from_keypress_event(KeyEvt("q", 0x4)), "<Control-Key-q>")
        self.assertIsNone(binding_sequence_from_keypress_event(KeyEvt("q", 0)))
        self.assertEqual(binding_sequence_from_keypress_event(KeyEvt("Escape", 0x1)), "<Shift-Escape>")

    def test_quickload_cancel_hotkey_validation(self) -> None:
        self.assertTrue(is_allowed_quickload_cancel_hotkey("<Escape>"))
        self.assertTrue(is_allowed_quickload_cancel_hotkey("<Control-Key-z>"))
        self.assertTrue(is_allowed_quickload_cancel_hotkey("<Control-q>"))
        self.assertEqual(normalize_quickload_cancel_hotkey("Escape"), "<Escape>")
        self.assertFalse(is_allowed_quickload_cancel_hotkey("<Key-a>"))
        self.assertFalse(is_allowed_quickload_cancel_hotkey("<bogus>"))

    def test_path_like_table_column_ids(self) -> None:
        self.assertTrue(is_path_like_table_column("source_file", None))
        self.assertTrue(is_path_like_table_column("sources", None))
        self.assertTrue(is_path_like_table_column("workbook_path", None))
        self.assertTrue(is_path_like_table_column("dss_file", "dss_file"))
        self.assertFalse(is_path_like_table_column("employee", "source_file"))
        self.assertFalse(is_path_like_table_column("details", None))

    def test_filter_employee_names_supports_all_employee_and_group_modes(self) -> None:
        employees = ["Alice Smith", "Bob Jones", "Charlie West"]
        groups = {"Crew A": ["Alice Smith", "Charlie West"]}
        self.assertEqual(
            filter_employee_names(employees, FilterSelection(mode="all", value="All Employees"), groups),
            {"Alice Smith", "Bob Jones", "Charlie West"},
        )
        self.assertEqual(
            filter_employee_names(employees, FilterSelection(mode="employee", value="Bob Jones"), groups),
            {"Bob Jones"},
        )
        self.assertEqual(
            filter_employee_names(employees, FilterSelection(mode="group", value="Crew A"), groups),
            {"Alice Smith", "Charlie West"},
        )
        self.assertEqual(
            filter_employee_names(
                employees,
                FilterSelection(mode="employee_multi", value="Alice Smith, Bob Jones", values=("Alice Smith", "Bob Jones")),
                groups,
            ),
            {"Alice Smith", "Bob Jones"},
        )

    def test_ignored_name_typos_round_trip(self) -> None:
        with self.workspace_json("ignored_typos") as path:
            ignored = {typo_warning_key("Alic Smith", "Alice Smith")}
            save_ignored_name_typos(path, ignored)
            self.assertEqual(load_ignored_name_typos(path), ignored)

    def test_pf_numbers_for_records_and_subject_formatting(self) -> None:
        records = [
            type('Record', (), {'source_file': 'PF26024-2 Alpha DSS.xlsx'})(),
            type('Record', (), {'source_file': 'PF26024-3 Beta DSS.xlsx'})(),
            type('Record', (), {'source_file': 'No PF Here.xlsx'})(),
        ]
        self.assertEqual(pf_numbers_for_records(records), 'PF26024-2, PF26024-3')
        subject = format_email_subject(
            'Hours for {first_name} - {week_start} to {week_end}',
            'Alice Smith',
            date(2026, 4, 6),
            date(2026, 4, 12),
            records,
        )
        self.assertIn('PF26024-2', subject)
        self.assertIn('PF26024-3', subject)

    def test_email_templates_round_trip(self) -> None:
        with self.workspace_json("email_templates") as path:
            save_email_templates(path, "Subject {first_name}", "<p>Hello {first_name}</p>{hours_table}")
            subject_template, body_template = load_email_templates(path)

            self.assertEqual(subject_template, "Subject {first_name}")
            self.assertIn("{hours_table}", body_template)

    def test_build_bug_report_html_includes_key_fields(self) -> None:
        html = build_bug_report_html(
            current_profile_name="Default",
            app_root=Path(r"C:\Temp\DSSHoursTracker"),
            snapshot_path=Path(r"C:\Temp\DSSHoursTracker\diagnostic_snapshot.json"),
            loaded_sources=[Path(r"C:\Work\Alpha.xlsx")],
            cache_status_by_path={Path(r"C:\Work\Alpha.xlsx"): "Disk Hit"},
        )

        self.assertIn("Summary", html)
        self.assertIn("Steps to reproduce", html)
        self.assertIn("diagnostic_snapshot.json", html)
        self.assertIn("Alpha.xlsx", html)
        self.assertIn("Disk Hit", html)

    def test_get_app_root_uses_localappdata_when_available(self) -> None:
        with mock.patch.dict(os.environ, {"LOCALAPPDATA": r"C:\Temp\AppData"}, clear=False):
            app_root = get_app_root()
            self.assertTrue(str(app_root).endswith("DSSHoursTracker"))

    def test_remove_config_keys_keeps_remaining_payload(self) -> None:
        with self.workspace_json("config_keys") as path:
            path.write_text(
                '{"employee_emails": {"Alice Smith": "alice@example.com"}, "table_layouts": {"daily_raw": {}}}',
                encoding="utf-8",
            )

            remove_config_keys(path, ["table_layouts"])

            payload = path.read_text(encoding="utf-8")
            self.assertIn("employee_emails", payload)
            self.assertNotIn("table_layouts", payload)


if __name__ == "__main__":
    unittest.main()
