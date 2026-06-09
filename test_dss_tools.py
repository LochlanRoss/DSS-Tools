from __future__ import annotations

"""
Fast unit tests (no RUN_SLOW_TESTS required).

Workbook / load_tracker_data / cache integration tests live in
`test_dss_tools_integration.py` and run when RUN_SLOW_TESTS=1.
"""

import os
import subprocess
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

from dss_hours_tracker import (
    AppSettings,
    DEFAULT_UI_THEME,
    OUTLOOK_NAME_RULE_LABEL,
    _bug_report_attachment_strings_to_try,
    build_employee_email_list_label,
    load_employee_name_overrides,
    _is_updater_executable_path,
    binding_sequence_from_keypress_event,
    az2_revision_matches_sheet_name,
    build_bug_report_html,
    create_bug_report_draft,
    checksum_for_asset_name,
    choose_preferred_outlook_resolution,
    choose_release_checksum_asset,
    choose_release_installer_asset,
    build_permission_denied_message,
    is_newer_version,
    normalize_release_version,
    parse_checksum_manifest,
    parse_latest_release_payload,
    process_workbook_bytes,
    compute_bytes_hash,
    compute_all_dated_sheet_hashes,
    format_email_address_display,
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
    find_outlook_display_name_typos,
    find_address_book_name_typos,
    filter_employee_names,
    FormattingProfile,
    get_app_root,
    is_alert_triggered,
    is_allowed_quickload_cancel_hotkey,
    is_path_like_table_column,
    is_unmetered_wifi_profile,
    load_email_templates,
    load_missing_email_suppressions,
    load_formatting_profiles,
    load_app_settings,
    normalize_ui_hex_color,
    parse_ui_theme_payload,
    DailyRecord,
    build_outlook_name_mismatch_findings,
    build_tracker_data_with_status,
    tracker_data_invalidated_for_cache_clear,
    load_employee_emails,
    load_employee_outlook_display_names,
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
    save_missing_email_suppressions,
    save_employee_name_overrides,
    save_formatting_profiles,
    save_email_templates,
    save_app_settings,
    save_table_layout,
    select_preferred_dated_sheets,
)
from test_dss_tools_fixtures import DssToolsFixtures


class DssToolsTests(DssToolsFixtures):
    def test_main_installer_postinstall_cleanup(self) -> None:
        from dss_hours_tracker import main

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "DSSTools"
            root.mkdir()
            (root / "dss_hours_tracker_config.json").write_text("{}", encoding="utf-8")
            (root / "leftover.bin").write_bytes(b"x")
            extra = root / "extra_dir"
            extra.mkdir()
            (extra / "a.txt").write_text("a", encoding="utf-8")
            with mock.patch("dss_hours_tracker.get_app_root", return_value=root):
                with mock.patch.object(sys, "argv", ["dss_hours_tracker.py", "--installer-postinstall-cleanup"]):
                    rc = main()
            self.assertEqual(rc, 0)
            self.assertTrue((root / "dss_hours_tracker_config.json").is_file())
            self.assertFalse((root / "leftover.bin").exists())
            self.assertFalse(extra.exists())

    def test_is_updater_executable_path_staged_temp_name(self) -> None:
        self.assertTrue(_is_updater_executable_path(Path(r"C:\Temp\dss_tools_updater_ab12cd.exe")))
        self.assertTrue(_is_updater_executable_path(Path(r"C:\Program Files\DSS Tools\DSSToolsUpdater.exe")))
        self.assertFalse(_is_updater_executable_path(Path(r"C:\Windows\notepad.exe")))

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

    def test_process_workbook_bytes_parses_signin_rows_and_time_fallback(self) -> None:
        with self.workspace_files("signin_source") as path:
            self.build_signin_source_workbook(path)
            records, warnings, health = process_workbook_bytes(path, path.read_bytes())

        self.assertFalse(warnings)
        self.assertFalse(health)
        self.assertEqual(len(records), 3)
        self.assertEqual(records[0].employee, "Lochlan Ross")
        self.assertEqual(records[0].st, 2.0)
        self.assertEqual(records[0].source_sheet, "Sheet1")
        self.assertEqual(records[0].source_ranges, "Sign-in name A5; data C5:R5")
        self.assertEqual(records[1].employee, "Lochlan Ross")
        self.assertEqual(records[1].st, 8.0)
        self.assertEqual(records[1].source_ranges, "Sign-in name A5; continuation C6:R6")
        self.assertEqual(records[2].employee, "Hayden Roddis")
        self.assertEqual(records[2].st, 10.0)
        self.assertEqual(records[2].source_ranges, "Sign-in name A7; data C7:R7")

    def test_compute_all_dated_sheet_hashes_includes_signin_weekly_sheets(self) -> None:
        with self.workspace_files("signin_weekly") as path:
            self.build_signin_weekly_workbook(path)
            hashes = compute_all_dated_sheet_hashes(path.read_bytes())

        self.assertEqual(set(hashes), {"2026-06-08", "2026-06-09"})
        self.assertTrue(all(hashes.values()))

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
            'html_url': 'https://github.com/LochlanRoss/DSS-Tools/releases/tag/v0.2.1',
            'published_at': '2026-05-01T00:00:00Z',
            'body': 'Notes',
            'assets': [
                {'name': 'DSSToolsSetup.exe', 'browser_download_url': 'https://example.invalid/setup.exe', 'size': 1234, 'content_type': 'application/octet-stream'},
                {'name': 'checksums.txt', 'browser_download_url': 'https://example.invalid/checksums.txt'},
            ],
        }
        info = parse_latest_release_payload(payload)
        self.assertEqual(info['version'], '0.2.1')
        self.assertEqual(info['tag_name'], 'v0.2.1')
        self.assertEqual(info['asset_names'], ['DSSToolsSetup.exe', 'checksums.txt'])
        self.assertEqual(info['assets'][0]['download_url'], 'https://example.invalid/setup.exe')

    def test_release_asset_selection_and_checksum_parsing(self) -> None:
        release_info = {
            'assets': [
                {'name': 'notes.zip', 'download_url': 'https://example.invalid/notes.zip'},
                {'name': 'DSSToolsSetup.exe', 'download_url': 'https://example.invalid/setup.exe'},
                {'name': 'checksums.txt', 'download_url': 'https://example.invalid/checksums.txt'},
            ]
        }
        installer = choose_release_installer_asset(release_info)
        checksum_asset = choose_release_checksum_asset(release_info)
        manifest = (
            'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa *DSSToolsSetup.exe\n'
            'bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb  other.msi\n'
        )
        self.assertEqual(installer['name'], 'DSSToolsSetup.exe')
        self.assertEqual(checksum_asset['name'], 'checksums.txt')
        self.assertEqual(
            checksum_for_asset_name(manifest, 'DSSToolsSetup.exe'),
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

    def test_employee_outlook_display_names_round_trip(self) -> None:
        with self.workspace_json("outlook_names") as path:
            save_employee_emails(
                path,
                {"Andrea Kolodinski": "a@example.com"},
                {"Andrea Kolodinski": "Andrea Kolodinsky"},
            )
            self.assertEqual(load_employee_emails(path)["Andrea Kolodinski"], "a@example.com")
            self.assertEqual(
                load_employee_outlook_display_names(path)["Andrea Kolodinski"],
                "Andrea Kolodinsky",
            )

    def test_save_employee_emails_without_outlook_arg_preserves_display_names(self) -> None:
        with self.workspace_json("preserve_outlook") as path:
            save_employee_emails(
                path,
                {"Andrea Kolodinski": "a@example.com"},
                {"Andrea Kolodinski": "Andrea Kolodinsky"},
            )
            save_employee_emails(path, {"Andrea Kolodinski": "a@example.com"})
            self.assertEqual(
                load_employee_outlook_display_names(path)["Andrea Kolodinski"],
                "Andrea Kolodinsky",
            )

    def test_build_outlook_name_mismatch_findings_lists_day_and_file(self) -> None:
        records = [
            DailyRecord(
                source_path=Path("C:/data/PF1.xlsx"),
                source_file="PF26024-1.xlsx",
                work_date=date(2026, 4, 8),
                source_sheet="Sheet1",
                employee="Andrea Kolodinski",
                st=8.0,
                ot=0.0,
                dt=0.0,
                source_ranges="",
            ),
        ]
        outlook = {"Andrea Kolodinski": "Andrea Kolodinsky"}
        findings = build_outlook_name_mismatch_findings(records, outlook, set())
        self.assertEqual(len(findings), 1)
        f = findings[0]
        self.assertTrue(f.outlook_name_rule)
        self.assertEqual(f.hour_type, OUTLOOK_NAME_RULE_LABEL)
        self.assertEqual(f.trigger_date, date(2026, 4, 8))
        self.assertEqual(f.source_files, "PF26024-1.xlsx")
        self.assertIn("Andrea Kolodinsky", f.reason)

    def test_find_outlook_display_name_typos_matches_address_book_spelling(self) -> None:
        records = [
            DailyRecord(
                source_path=Path("C:/data/PF1.xlsx"),
                source_file="PF26024-1.xlsx",
                work_date=date(2026, 4, 8),
                source_sheet="Sheet1",
                employee="Andrea Kolodinski",
                st=1.0,
                ot=0.0,
                dt=0.0,
                source_ranges="",
            ),
        ]
        outlook = {"Andrea Kolodinski": "Andrea Kolodinsky"}
        warnings = find_outlook_display_name_typos(["Andrea Kolodinski"], outlook, records)
        self.assertEqual(len(warnings), 1)
        self.assertEqual(warnings[0].similar_employee, "Andrea Kolodinsky")
        self.assertIn("PF26024-1.xlsx", warnings[0].locations[0])

    def test_choose_preferred_outlook_resolution_prefers_jatech_domain(self) -> None:
        chosen = choose_preferred_outlook_resolution(
            [
                type("R", (), {"email": "name@gmail.com", "display_name": "Dwayne Moffatt"})(),
                type("R", (), {"email": "dwayne.moffatt@jatechpowersystems.com", "display_name": "Dwayne Moffatt"})(),
            ]
        )
        self.assertIsNotNone(chosen)
        self.assertEqual(chosen.email, "dwayne.moffatt@jatechpowersystems.com")

    def test_find_address_book_name_typos_catches_small_surname_typo_only(self) -> None:
        records = [
            DailyRecord(
                source_path=Path("C:/data/PF1.xlsx"),
                source_file="PF26024-1.xlsx",
                work_date=date(2026, 4, 8),
                source_sheet="Sheet1",
                employee="Dwayne Moffat",
                st=8.0,
                ot=0.0,
                dt=0.0,
                source_ranges="",
            ),
        ]
        warnings = find_address_book_name_typos(
            ["Dwayne Moffat", "Chris Georget"],
            ["Dwayne Moffatt", "Chris McClean"],
            records,
        )
        self.assertEqual(len(warnings), 1)
        self.assertEqual(warnings[0].employee, "Dwayne Moffat")
        self.assertEqual(warnings[0].similar_employee, "Dwayne Moffatt")

    def test_employee_groups_round_trip(self) -> None:
        with self.workspace_json("groups") as path:
            save_employee_groups(path, {"Crew A": ["Alice Smith", "Bob Jones"], "Crew B": ["Charlie West"]})
            loaded = load_employee_groups(path)
            self.assertEqual(loaded["Crew A"], ["Alice Smith", "Bob Jones"])
            self.assertEqual(loaded["Crew B"], ["Charlie West"])

    def test_employee_name_overrides_round_trip(self) -> None:
        with self.workspace_json("employee_name_overrides") as path:
            save_employee_name_overrides(path, {"Manual Add"}, {"Hidden Typo"})
            added, hidden = load_employee_name_overrides(path)
            self.assertEqual(added, {"Manual Add"})
            self.assertEqual(hidden, {"Hidden Typo"})

    def test_missing_email_suppressions_round_trip(self) -> None:
        with self.workspace_json("missing_email_suppressions") as path:
            save_missing_email_suppressions(path, {"Alice Smith", "Bob Jones"})
            loaded = load_missing_email_suppressions(path)
            self.assertEqual(loaded, {"Alice Smith", "Bob Jones"})

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
            app_root=Path(r"C:\Temp\DSSTools"),
            snapshot_path=Path(r"C:\Temp\DSSTools\diagnostic_snapshot.json"),
            loaded_sources=[Path(r"C:\Work\Alpha.xlsx")],
            cache_status_by_path={Path(r"C:\Work\Alpha.xlsx"): "Disk Hit"},
        )

        self.assertIn("Summary", html)
        self.assertIn("Steps to reproduce", html)
        self.assertIn("diagnostic_snapshot.json", html)
        self.assertIn("Alpha.xlsx", html)
        self.assertIn("Disk Hit", html)
        self.assertIn("DSS Tools", html)

    def test_bug_report_attachment_strings_to_try(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            snap = Path(tmp) / "diagnostic_snapshot.json"
            snap.write_text('{"ok": true}', encoding="utf-8")
            strings, cleanup = _bug_report_attachment_strings_to_try(snap)
            self.assertGreaterEqual(len(strings), 1)
            self.assertEqual(len(cleanup), 1)
            self.assertTrue(cleanup[0].is_file())
            self.assertEqual(cleanup[0].read_text(encoding="utf-8"), '{"ok": true}')
            for path in cleanup:
                path.unlink(missing_ok=True)

    def test_format_email_address_display_marks_missing(self) -> None:
        self.assertEqual(format_email_address_display("person@example.com"), ("person@example.com", False))
        self.assertEqual(format_email_address_display("   "), ("(missing) \u26A0", True))

    def test_build_employee_email_list_label_marks_missing(self) -> None:
        self.assertEqual(
            build_employee_email_list_label("Alice Smith", ""),
            ("Alice Smith | (missing) \u26A0", True),
        )

    def test_create_bug_report_draft_retries_without_attachment_when_save_fails(self) -> None:
        class FakeAttachments:
            def __init__(self) -> None:
                self.paths: list[str] = []

            def Add(self, path: str, _mode: int) -> None:
                self.paths.append(path)

        class FakeMailItem:
            def __init__(self) -> None:
                self.To = ""
                self.Subject = ""
                self.HTMLBody = ""
                self.Attachments = FakeAttachments()
                self.saved = False

            def Save(self) -> None:
                if self.Attachments.paths:
                    raise RuntimeError("Path does not exist")
                self.saved = True

        class FakeOutlook:
            def __init__(self) -> None:
                self.items: list[FakeMailItem] = []

            def CreateItem(self, _kind: int) -> FakeMailItem:
                item = FakeMailItem()
                self.items.append(item)
                return item

        fake_outlook = FakeOutlook()
        fake_dispatch = mock.Mock(return_value=fake_outlook)
        fake_pythoncom = mock.Mock()
        with tempfile.TemporaryDirectory() as tmp:
            snapshot = Path(tmp) / "diagnostic_snapshot.json"
            snapshot.write_text('{"ok": true}', encoding="utf-8")
            with mock.patch("dss_hours_tracker.pythoncom", fake_pythoncom), mock.patch("dss_hours_tracker.win32com") as fake_win32com:
                fake_win32com.client.Dispatch = fake_dispatch
                warning = create_bug_report_draft(
                    "lross@jatechpowersystems.com",
                    "Bug",
                    "<p>Body</p>",
                    attachment_path=snapshot,
                )
        self.assertIsNotNone(warning)
        self.assertEqual(len(fake_outlook.items), 2)
        self.assertTrue(fake_outlook.items[-1].saved)
        self.assertIn("saved without the attachment", warning or "")

    def test_ensure_dss_tools_ico_script_uses_fallback_in_empty_repo(self) -> None:
        script = Path(__file__).resolve().parent / "tools" / "ensure_dss_tools_ico.py"
        fallback = Path(__file__).resolve().parent / "tools" / "default_dss_tools.ico"
        if not script.is_file() or not fallback.is_file():
            self.skipTest("packaged icon tooling not present")
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            proc = subprocess.run(
                [sys.executable, str(script), "--repo", str(root)],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(proc.returncode, 0, msg=proc.stderr + proc.stdout)
            self.assertTrue((root / "dss_tools.ico").is_file())

    def test_ensure_dss_tools_ico_script_prefers_png_when_present(self) -> None:
        try:
            from PIL import Image
        except ImportError:
            self.skipTest("pillow not installed")
        script = Path(__file__).resolve().parent / "tools" / "ensure_dss_tools_ico.py"
        if not script.is_file():
            self.skipTest("ensure_dss_tools_ico.py not present")
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            png = root / "dss_tools.png"
            Image.new("RGBA", (64, 48), (200, 40, 40, 255)).save(png, format="PNG")
            proc = subprocess.run(
                [sys.executable, str(script), "--repo", str(root)],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(proc.returncode, 0, msg=proc.stderr + proc.stdout)
            ico = root / "dss_tools.ico"
            self.assertTrue(ico.is_file())
            self.assertGreater(ico.stat().st_size, 100)
            with Image.open(ico) as im:
                im.load()
                sizes = im.info.get("sizes", set())
                self.assertGreaterEqual(len(sizes), 6, msg=str(sizes))
                self.assertTrue(all(w == h for w, h in sizes), msg=str(sizes))
                self.assertEqual(im.size, (256, 256), msg="Pillow ICO primary must be largest or Windows shows a smeared icon")

    def test_ensure_dss_tools_ico_force_overwrites_existing_ico(self) -> None:
        try:
            from PIL import Image
        except ImportError:
            self.skipTest("pillow not installed")
        script = Path(__file__).resolve().parent / "tools" / "ensure_dss_tools_ico.py"
        if not script.is_file():
            self.skipTest("ensure_dss_tools_ico.py not present")
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            png = root / "dss_tools.png"
            Image.new("RGBA", (64, 48), (200, 40, 40, 255)).save(png, format="PNG")
            ico = root / "dss_tools.ico"
            ico.write_bytes(b"not-a-real-ico")
            proc = subprocess.run(
                [sys.executable, str(script), "--repo", str(root), "--force"],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(proc.returncode, 0, msg=proc.stderr + proc.stdout)
            self.assertGreater(ico.stat().st_size, 100)

    def test_ensure_dss_tools_ico_canonical_png_beats_lone_ico(self) -> None:
        """When DSS-Tools Icon.png exists, a single stray *.ico must not replace the branded output."""
        try:
            from PIL import Image
        except ImportError:
            self.skipTest("pillow not installed")
        script = Path(__file__).resolve().parent / "tools" / "ensure_dss_tools_ico.py"
        if not script.is_file():
            self.skipTest("ensure_dss_tools_ico.py not present")
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            canonical = root / "DSS-Tools Icon.png"
            # Flat color survives ICO encode/verify (high-frequency gradients do not).
            Image.new("RGB", (256, 256), (80, 120, 200)).save(canonical, format="PNG")
            stray = root / "stray.ico"
            Image.new("RGBA", (32, 32), (10, 20, 30, 255)).save(
                stray, format="ICO", sizes=[(32, 32)]
            )
            proc = subprocess.run(
                [sys.executable, str(script), "--repo", str(root)],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(proc.returncode, 0, msg=proc.stderr + proc.stdout)
            self.assertIn("branding png", proc.stdout.lower())
            with Image.open(root / "dss_tools.ico") as im:
                im.load()
                self.assertEqual(im.size, (256, 256))

    def test_ensure_dss_tools_ico_canonical_regenerates_corrupt_ico_without_force(self) -> None:
        try:
            from PIL import Image
        except ImportError:
            self.skipTest("pillow not installed")
        script = Path(__file__).resolve().parent / "tools" / "ensure_dss_tools_ico.py"
        if not script.is_file():
            self.skipTest("ensure_dss_tools_ico.py not present")
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            canonical = root / "DSS-Tools Icon.png"
            Image.new("RGB", (256, 256), (80, 120, 200)).save(canonical, format="PNG")
            ico = root / "dss_tools.ico"
            ico.write_bytes(b"bogus")
            proc = subprocess.run(
                [sys.executable, str(script), "--repo", str(root)],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(proc.returncode, 0, msg=proc.stderr + proc.stdout)
            with Image.open(ico) as im:
                im.load()
                self.assertEqual(im.size, (256, 256))

    def test_get_app_root_uses_localappdata_when_available(self) -> None:
        with mock.patch.dict(os.environ, {"LOCALAPPDATA": r"C:\Temp\AppData"}, clear=False):
            app_root = get_app_root()
            self.assertTrue(str(app_root).endswith("DSSTools"))

    def test_get_app_root_prefers_legacy_folder_when_present(self) -> None:
        import tempfile

        with tempfile.TemporaryDirectory() as tmp:
            legacy = Path(tmp) / "DSSHoursTracker"
            legacy.mkdir()
            with mock.patch.dict(os.environ, {"LOCALAPPDATA": tmp}, clear=False):
                app_root = get_app_root()
            self.assertEqual(app_root, legacy)

    def test_get_app_root_prefers_primary_when_both_folders_exist(self) -> None:
        import tempfile

        with tempfile.TemporaryDirectory() as tmp:
            legacy = Path(tmp) / "DSSHoursTracker"
            primary = Path(tmp) / "DSSTools"
            legacy.mkdir()
            primary.mkdir()
            with mock.patch.dict(os.environ, {"LOCALAPPDATA": tmp}, clear=False):
                app_root = get_app_root()
            self.assertEqual(app_root, primary)

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
