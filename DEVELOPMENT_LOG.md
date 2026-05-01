# Development Log

This log tracks moderate and larger changes, along with active feature requests, notable bug work, and handoff-ready context.
Small fixes and very short edits are intentionally omitted unless they materially changed behavior.

## Logging Rule
- Record moderate or large feature work, behavioral changes, UI workflow changes, architecture changes, persistence changes, diagnostics, packaging, caching, and integration work.
- Do not record tiny cosmetic edits or very small bug fixes unless they changed user behavior in an important way.

## 2026-05-01

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
- This is intended to reduce false cache misses and false ùchangedù alerts caused by Excel/OneDrive autosave behavior.

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
- **Table layouts:** persisted `column_filters` in `table_layouts` JSON; restored on launch with the rest of each tableùs layout.
- **`Feature request list.txt`:** converted to a markdown checklist with completed items checked.
- Unit tests: added coverage for daily aggregation/rollup, daily rollup sort key, layout column-filter round-trip; **57** tests passing.

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
