# Development Log

This log tracks moderate and larger changes, along with active feature requests worth carrying forward.
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

### Persistence / Configuration Changes
- App settings now include automatic update-check and auto-download preferences.
- Table layouts continue to persist visibility, widths, and sort state, and now preserve user column order through saved display-column order.
- Diagnostic snapshots now include update-download state and the update cache path.

### Testing / Verification
- Added and updated unit coverage for updater helpers, checksum parsing, app-settings round-trip, and typo-warning helper behavior.
- Full unit suite passing after the above changes.

## Active Feature Requests
- Make the hash/diff process even more robust against non-DSS Excel metadata churn, especially with OneDrive autosave and open workbooks.
- Persist active per-column header filter selections across launches.
- Continue using this log for medium and large features going forward.

## Notes
- `Feature request list` remains the working scratchpad for requests.
- This log is the higher-signal historical record for meaningful completed work and notable active items.
