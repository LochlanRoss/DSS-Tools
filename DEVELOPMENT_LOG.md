# Development Log

This log tracks moderate and larger changes, along with active feature requests, notable bug work, and handoff-ready context.
Small fixes and very short edits are intentionally omitted unless they materially changed behavior.

## Logging Rule
- Record moderate or large feature work, behavioral changes, UI workflow changes, architecture changes, persistence changes, diagnostics, packaging, caching, and integration work.
- Do not record tiny cosmetic edits or very small bug fixes unless they changed user behavior in an important way.

## 2026-05-01

### Test layout: fast default vs slow integration

- Split workbook-heavy coverage into `test_dss_hours_tracker_integration.py`, gated with **`RUN_SLOW_TESTS`** (`1` / `true` / `yes` / `all`) and `@unittest.skipUnless` on the integration class so **`python -m unittest test_dss_hours_tracker`** stays fast.
- Shared temp-file / sample-workbook helpers live in **`test_dss_hours_tracker_fixtures.py`** (`DssHoursTrackerFixtures`) to avoid duplicating setup between modules.
- **`test_dss_hours_tracker.py`** keeps sort keys, pure helpers, settings/layout round-trips, updater helpers, and similar tests without requiring the env var.
- Full DSS test discovery remains **`python -m unittest discover -s . -p "test_dss*.py"`**; with `RUN_SLOW_TESTS=1`, integration tests execute instead of skipping (still **63** tests total: **35** fast + **28** slow).
- **`test_dss_hours_tracker.py`** was trimmed so workbook/cache/`load_tracker_data` cases exist only in the integration module (no duplicate `test_*` definitions); imports were reduced to match the slimmer fast suite. Integration module imports include helpers such as **`format_email_subject`** and **`compute_bytes_hash`** where cache and email tests need them.
- **`README.md` Tests** subsection documents the split, the env var, fixture module path, and PowerShell examples for fast-only vs full discovery runs.

### UI theme colours (Configuration)

- Added **`UiThemeColors`** (frozen defaults: soft rose alert rows, teal-leaning crew totals, slate tooltips, pink report outline) persisted under `app_settings.ui_theme` in config JSON.
- **Settings ? Configuration ? Appearance:** hex entry + colour picker per semantic slot, **Reset colours to sample defaults**, validated on Apply; all `DataTable` instances and email preview table refresh tags; Reports outline and Formatting Rules **ToolTip** popups use the saved tooltip colours.
- Helpers: **`normalize_ui_hex_color`**, **`parse_ui_theme_payload`** (partial / invalid keys fall back to defaults).

### Loading hints file (auxiliary; UI not wired yet)

- Added **`LOADING_HINTS.txt`** at the repository root: many single-line hints covering parsing boundaries (`K25:AZ36`), revision behaviour, AZ2 warnings, cache and semantic hash, quick load and cancel hotkey, filters and layouts, Outlook and email drafts, formatting rules, combined-name caveats, maintenance buttons, `%LOCALAPPDATA%\DSSHoursTracker`, and similar limitations.
- First lines in the file are **`#` comments** documenting convention: a future loader can skip lines starting with `#` and rotate the rest on splash or progress UI.

### Maintainer / handoff notes (no code change)

- **Distribution:** Clarified for maintainers that end users should receive a single Windows installer (e.g. PyInstaller-built binary wrapped with Inno Setup, NSIS, or WiX) for double-click install without separate Python or `pip`; standard installers register **Settings ? Apps** uninstall and do not require a separate uninstaller download.
- **Visual design:** Reviewed current styling (`DataTable` alert and crew-total tags, Reports outline highlight, default `ttk` elsewhere); recorded recommendation to centralize palette tokens and optionally add a header/toolbar band and accent before larger dark-mode work.

### Completed Changes
- Added GitHub-based update checking with installed-version detection.
- Added automatic background update checks on startup.
- Added automatic update download on unmetered Wi-Fi with installer handoff after app close.
- Added release asset parsing and checksum verification support for downloaded installers.
- Added persistent ignore support for repeated name-typo warnings.
- Added PF number support in email draft subjects.
- Added drag-to-reorder columns in table views.
- Added right-click header menus for per-column filtering while keeping left-click sorting intact.
- Added manual `Check Name Typos` action in the `Error Report` page.
- Added a note in the typo warning popup that the warning can be disabled in settings.
- Added `UI Standardization Notes.md` for future project handoffs and reusable UI patterns.
- Added `DEVELOPMENT_LOG.md` for ongoing medium/large change tracking.

### Persistence / Configuration Changes
- App settings now include automatic update-check and auto-download preferences.
- Table layouts continue to persist visibility, widths, sort state, and user column order through saved display-column order.
- Diagnostic snapshots now include update-download state and the update cache path.

### Bug Fixes
- Fixed the weekly rollup cross-PF grouping bug when multi-employee filtering was combined with sorting by week.
- The fix adds a section-aware custom sort path for the weekly-by-PF rollup table so employee rows stay grouped with the correct PF/source-file crew-total row.

### Hash / Cache Improvements
- Replaced the previous workbook-content hash preference with a DSS-semantic hash path.
- The semantic hash now considers only:
  - preferred dated sheets after revision selection
  - the DSS-relevant cell window `K25:AZ36`
- The semantic hash now ignores:
  - workbook metadata churn
  - non-dated sheets
  - changes outside the DSS-relevant range
- This is intended to reduce false cache misses and false "changed" alerts caused by Excel/OneDrive autosave behavior.

### Safety / Source Workbook Audit
- Reviewed source-workbook access paths in `dss_hours_tracker.py`.
- Current conclusion: the app reads source DSS workbooks and does not intentionally write back to the selected source workbook path.
- Source-workbook handling is read-only in the app path:
  - direct `read_bytes()` for normal access
  - `openpyxl` loaded from in-memory bytes in `read_only=True` mode
  - Excel fallback uses `SaveCopyAs(...)` to a temporary file, then reads the temp copy
- The app does write to:
  - JSON cache files
  - config files
  - diagnostic snapshot files
  - explicit user-chosen export paths
- It does not contain a known code path that overwrites, saves, or deletes the selected source DSS workbook.

### Testing / Verification
- Added and updated unit coverage for:
  - updater helpers
  - checksum parsing
  - app-settings round-trip
  - typo-warning helper behavior
  - weekly rollup grouped sorting behavior
  - DSS semantic hashing behavior
- Full unit suite passing after the above changes.

### Follow-up (same release cycle)

- **Summaries:** added `Daily by PF#` and `Combined Summary by Day`; renamed tabs to `Weekly by PF#` and `Combined Summary by Week`. `TrackerData` now carries `daily_summary`, `daily_rollup`, and `combined_daily_summary` built beside the weekly model.
- **Table layouts:** persisted `column_filters` in `table_layouts` JSON; restored on launch with the rest of each table's layout.
- **`Feature request list.txt`:** converted to a markdown checklist with completed items checked.
- Unit tests: added coverage for daily aggregation/rollup, daily rollup sort key, layout column-filter round-trip; **57** tests passing.

### Later same day (feature batch)

- **AZ2 vs sheet revision:** `AZ2` is read on each dated sheet; when it disagrees with whether the tab name carries a revision suffix (or revision level), a `SheetParseWarning` is emitted (`Revision Indicator AZ2 Mismatch`).
- **Revision parsing:** `REVISION_PATTERN` / `parse_sheet_revision` accept additional separators (for example `r-1`, `r.1`, `rev 12`) so rev-up logic stays consistent across naming styles.
- **Per-sheet hashing and partial reload:** dated-sheet cell-window digests are stored with cache entries; `Update View` / `load_tracker_data` can reuse unchanged sheets and only re-parse sheets whose digest changed (**Partial Refresh**), while the overall semantic hash still summarizes the workbook.
- **Reports chrome:** when filtered errors or parse warnings exist, the Error Report / Sheet Parse Warnings tabs (and the parent Reports group) show an alert outline until the loaded data no longer has those issues.
- **UI:** double-click **source file** opens the workbook; column drag shows a vertical insertion indicator; Summaries notebook order is daily views before weekly (`Daily by PF#`, `Weekly by PF#`, combined day, combined week).
- **Load order:** multi-DSS loads sort sources by file **mtime** (newest first); `process_workbook_bytes` progress text includes **PF#-#** plus sheet name.
- **Quick load:** last successfully opened DSS paths persist in config; optional startup reload (800 ms after launch if no CLI sources), progress hint label, cancel button plus **configurable global hotkey** (presets and a "Press keys..." capture dialog, stored as a Tk virtual event string). `AppSettings` adds `quickload_last_sources_enabled` and `quickload_cancel_hotkey`.
- **Tests:** app-settings round-trip and hotkey helper tests; full suite **63** tests passing (includes path-like column helper and cache/hash cases).

### Comprehensive progress log (recent implementation work)

Consolidated detail for the same feature batch (parsing, cache, UI, settings, tests, and documentation).

**Parsing and sheet validation**

- **`AZ2` revision cross-check:** Helpers such as `az2_revision_matches_sheet_name`, `revision_level_from_az2` / `revision_presence_from_az2`, and `_normalize_az2_cell_text` compare fixed cell `AZ2` on each dated sheet to the revision implied by the tab name. Mismatches emit `SheetParseWarning` with issue `Revision Indicator AZ2 Mismatch` (including blank `AZ2` when the sheet name still carries a revision suffix).
- **Numeric `AZ2` vs yes/no:** Ambiguous yes/no handling was tightened so a numeric `AZ2` such as `1` is interpreted through revision-level logic instead of colliding with a yes/no token set (which could wrongly validate against a sheet name like `R2`).
- **`REVISION_PATTERN` / `parse_sheet_revision`:** Pattern tolerates more separators (e.g. `r-1`, `r.1`, `rev 12`, `Revision 1`) so highest-revision selection stays consistent.

**Caching, hashing, and reload**

- **Per-sheet digests:** `compute_all_dated_sheet_hashes` and related code hash the dated-sheet window of interest; the value map used for semantic hashing includes **`AZ2`** so revision-indicator edits invalidate the correct sheet digest.
- **Combined digest / semantic hash:** `combine_sheet_hashes` and `compute_dss_semantic_hash` fold per-sheet digests into the workbook-level semantic fingerprint alongside full-file hashing where used.
- **Cache payload:** Cached daily records / analysis store `sheet_hashes` where applicable; older disk hits can backfill missing `sheet_hashes` when analysis is re-saved.
- **Partial workbook reload:** `merge_workbook_from_cache_by_sheet_hashes` with `read_workbook_cache_payload` reloads only sheets whose digests changed when appropriate; status can read **Partial Refresh**. `process_workbook_bytes` supports `restrict_to_sheet_names` for targeted re-parse.
- **Multi-file load order:** `load_tracker_data` sorts normalized paths by filesystem **mtime** (newest first). True "newest by latest in-file sheet date" ordering was not implemented; the existing preview callback still surfaces partial UI while long parses run.

**Reports and tables (UI)**

- **Parse warnings:** Parse-warning rows use the same `alert` styling as error-report rows where wired.
- **Tab chrome:** `reports_outline` and `_sync_reports_alert_chrome` drive outline / `(!)` labelling on **Error Report**, **Sheet Parse Warnings**, and the parent **Reports** group when filtered errors or parse warnings exist; cleared when data is refreshed clean or views are cleared.
- **Open source file:** `DataTable` uses `open_source_file_callback` / `source_file_column`; double-click opens via `_open_displayed_source_file` (shell open) on supported tables (daily, summaries, parse warnings, workbook health, audit trail, etc.).
- **Column reorder:** Drag shows a vertical insertion line (`_column_drag_line`); hidden on release or invalid drop.
- **Column auto-width:** `DataTable` measures heading + cell strings (`tkinter.font`) after each render; path-like column ids are capped tighter; saved layout `column_widths` are no longer restored on load; columns use `stretch=False` so neighbour columns are not resized by the layout engine (see later toolbar word wrap).
- **Summaries tab order:** Notebook order is **Daily by PF#**, **Weekly by PF#**, **Combined Summary by Day**, **Combined Summary by Week**.

**Progress text**

- Per-sheet progress uses `extract_pf_identifier` so messages show **PF#-#** from the filename together with the sheet name.

**Quick load, settings, persistence**

- **Last paths:** `save_last_open_dss_paths` / `load_last_open_dss_paths` persist `last_open_dss_paths` in config; successful loads update the list.
- **Startup:** With quick load enabled, no CLI `initial_source`, and no data loaded, `after(800, _maybe_quickload_last_sources)` restores paths that still exist.
- **Session / hint:** `_quickload_session` gates cancel behaviour; `_refresh_quickload_hint_label` shows a short hint under the progress bar during quick load.
- **Hotkey:** `quickload_cancel_hotkey` (Tk virtual event, default `<Escape>`) validated via `is_allowed_quickload_cancel_hotkey` / `normalize_quickload_cancel_hotkey`; `_register_quickload_cancel_hotkey` uses `bind_all` / `unbind_all` and runs after apply settings and major config resets. Settings UI: presets plus **Press keys...** capture using `binding_sequence_from_keypress_event`.
- **`AppSettings`:** Adds `quickload_last_sources_enabled` and `quickload_cancel_hotkey`; round-tripped in JSON; invalid saved hotkeys fall back to `<Escape>`.

**Diagnostics**

- Diagnostic snapshot JSON includes the quick-load flags under `app_settings`.

**Testing**

- Extended tests: revision parsing, preferred-sheet / AZ2 behaviour, `combine_sheet_hashes`, app-settings round-trip (quick-load fields), hotkey helpers, path-like column detection, disk cache `file_hash` contract, `rv` sheet suffix. **63** tests, `python -m unittest test_dss_hours_tracker -q`.

**Documentation**

- **`README.md`:** Summaries order, quick load, partial refresh / per-sheet hashing, PF-in-progress text, double-click source, AZ2 warnings, column-drag indicator.
- **`Feature request list.txt`:** Related checklist items marked complete.

### Table column auto-width (same cycle)

- **`DataTable`:** After each filtered render, column widths are set from `tkinter.font` text measurement of the heading and visible cell values (newlines flattened for width).
- **Path-like columns** (`source_file`, configured `source_file_column`, `sources`, any `*_path`) use a lower pixel cap than general text; other columns use a high ceiling so wide text (for example **Details**) still fits without unbounded growth.
- **Saved `column_widths` in `table_layouts`:** Widths from disk are no longer applied on load so content-based sizing stays consistent; widths are still written when layouts save for backward compatibility.
- **Stretch:** Only the last visible column uses `stretch=True` so extra horizontal space is absorbed there.
- **Tests:** `is_path_like_table_column` unit coverage; full suite **63** tests with `python -m unittest test_dss_hours_tracker -q`.

### Table UX and cache/revision fixes (same cycle)

- **Treeview columns:** all logical columns use `stretch=False` so interactive width changes do not auto-shrink the next column; horizontal scroll covers overflow.
- **Word wrap:** per-`DataTable` toolbar button `Word wrap (off)` / `(on)`; wraps displayed cell text to the current column pixel width and applies a custom `ttk.Style` row height so wrapped lines show; source-file double-click strips embedded newlines before opening.
- **Disk cache key:** after a **Partial Refresh**, `save_cached_daily_records` now stores the same **workbook content hash** as `load_cached_daily_records` expects (previously the semantic fingerprint was written, causing a perpetual disk miss for that file until re-parse).
- **Sheet revision names:** `REVISION_PATTERN` recognises **`rv`** + digits (e.g. `2026-04-26 rv1`) so `parse_sheet_revision` aligns with AZ2 revision level `1`.
- **Tests:** `parse_sheet_revision` cases for `rv1`/`RV2`; `test_disk_cache_load_requires_workbook_content_hash`; suite **63** tests.

## Current Open Bugs
These remain on `Known bugs List.txt` until explicitly confirmed fixed by the user.

- Cache inconsistency on reload:
  - same three DSS files sometimes report `0 memory hit / 2 disk hit / 1 miss` despite no intentional file changes
  - likely next area to investigate if still reproducible after the semantic hash update

## Active Feature Requests / Objectives
- Continue monitoring whether the DSS-semantic hash fully resolves OneDrive/autosave false-change behavior (see open bug note below).
- Continue using this development log for medium and large features going forward.

## Handoff Notes
- `Feature request list` remains the working scratchpad for requests.
- `Known bugs List.txt` should only be edited after direct user confirmation of a fix.
- `UI Standardization Notes.md` is the reusable generalized UI handoff/reference document.
- `DEVELOPMENT_LOG.md` is the higher-signal historical record for meaningful completed work and notable active items.
- The repo path is now writable in-session:
  - `C:\Users\LochlanRoss\Documents\GitHub\DSS Viewer`
- The current Codex config is using Windows `sandbox = "unelevated"` with per-project trust entries.
