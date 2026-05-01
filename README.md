# DSS Hours Tracker

Desktop GUI for opening one or more DSS `.xlsx` workbooks, extracting labour hours from the protected daily sheets, and working with the results directly in the app.

## What It Does

- Opens one or more DSS Excel workbooks
- Parses only `K25:AZ36`
- Detects dated daily sheets named like `YYYY-MM-DD`, including suffixes such as `R1`, `rev 2`, and similar variants
- Prefers the highest revision sheet for a date when duplicates exist
- Extracts employee hours from the left and right worker blocks
- Aggregates hours into Monday-Sunday weeks
- Supports multi-DSS rollups across matching employee names
- Keeps all results usable without Excel formulas linked back to the source workbook

## Run

Launch the GUI:

```powershell
python dss_hours_tracker.py
```

Open workbook(s) immediately on launch:

```powershell
python dss_hours_tracker.py "Phase1 DSS.xlsx" "Phase2 DSS.xlsx"
```

Install locally as an app:

```powershell
pip install .
dss-hours-tracker
```

## Tests

Workbook parsing, `load_tracker_data`, disk cache, and related cases live in `test_dss_hours_tracker_integration.py`. That module’s class is **skipped by default** so day-to-day runs stay quick.

**Fast suite** (unit tests and JSON round-trips only; integration skipped):

```powershell
python -m unittest test_dss_hours_tracker -q
```

**Full suite** (include integration; allow roughly a minute on a typical machine):

```powershell
$env:RUN_SLOW_TESTS = "1"
python -m unittest discover -s . -p "test_dss*.py" -q
```

`RUN_SLOW_TESTS` is treated as enabled when its value is `1`, `true`, `yes`, or `all` (case-insensitive). Shared workbook helpers for both modules are in `test_dss_hours_tracker_fixtures.py`.

## Current Interface

The app uses grouped navigation with two tab rows:

- Top row groups:
  - `Data`
  - `Summaries`
  - `Reports`
  - `Settings`

- Pages inside each group:
  - `Data`
    - `Daily Raw` (optional, can be hidden in Configuration)
    - `Week Totals`
  - `Summaries`
    - `Daily by PF#`
    - `Weekly by PF#`
    - `Combined Summary by Day`
    - `Combined Summary by Week`
  - `Reports`
    - `Error Report`
    - `Email Drafts`
  - `Settings`
    - `Configuration`
    - `Employee List`
    - `Employee Groups`
    - `Formatting Rules`

## Key Features

- `Update View` reloads all currently selected DSS files
- Background loading keeps the UI responsive
- Overall progress bar shows multi-DSS progress; sheet progress text includes the **PF#-#** token from the filename where available
- `Cancel` can abort DSS loading and Outlook email sync; during **quick load** (re-opening the last DSS set on startup) you can also cancel with a configurable hotkey (defaults to Escape; presets and key capture live under **Settings → Configuration**)
- Optional **quick load**: on startup the app can re-open the same DSS workbook paths as last time (saved in config), with a short hint next to the progress bar; turn the behaviour off in Configuration
- Multi-file loads prefer **newer files first** (by filesystem modified time); the parser can emit an early partial preview while older workbooks are still loading
- Same-file updates are skipped when the **DSS semantic hash** (dated sheets and `K25:AZ36` only) is unchanged; per-sheet digests allow **partial refresh** when only some dated sheets change
- Parsed DSS data is cached on disk for up to 7 days
- Optional **GitHub release** update check; automatic download on unmetered Wi‑Fi; install prompt after download
- Loaded DSS files are checked periodically in the background for changes
- Revision sheets such as `R1` / `rev 2` override the non-revised sheet for the same date
- Conditional highlighting uses saved job-specific formatting profiles
- Table layouts persist per page:
  - shown/hidden columns
  - column order (drag column headers to reorder; a drop indicator shows the insertion point)
  - sort column and direction
  - per-column header filter selections (where filters are used)
- Column widths **auto-fit** to the heading and visible cell text after each refresh (saved pixel widths in config are not reapplied so layout stays data-driven); **Source File** / path-like columns (`sources`, `*_path`) use a tighter cap so long filenames stay readable without stretching the whole grid
- Column headers use **no stretch-to-fill**: resizing one column does not steal width from the neighbour; if the combined column widths exceed the view, use the **horizontal scrollbar**
- Each data table has a **Word wrap (off) / (on)** toolbar toggle (per table): when on, long cell text wraps within the column width and row height increases so wrapped lines stay visible
- Double-click a **source file** cell in supported tables to open that workbook in the desktop shell
- **AZ2** is checked against the sheet name revision pattern; mismatches appear as sheet parse warnings, and Reports tab chrome highlights when errors or parse warnings are present

## Filtering

The top `Filter` control is a checklist popup:

- `All Employees`
- `Uncheck All`
- one checkbox per employee

You can keep the filter popup open while selecting multiple employees. It closes when you click away or press `Esc`.

The filter applies across the main data views and the email draft preview.

## Summaries views

- `Weekly by PF#`
  - one row per employee per source DSS (PF) per week
  - includes `Whole Crew` total rows per PF per week
- `Daily by PF#`
  - one row per employee per source DSS per calendar day
  - includes `Whole Crew` total rows per PF per day
- `Combined Summary by Week`
  - combines matching employee names across all loaded DSS files, aggregated by week
- `Combined Summary by Day`
  - same cross-file name matching, aggregated by calendar day
- `Week Totals` (under **Data**)
  - one row per week for whole-crew totals across the currently filtered result set

## Error Reporting

`Error Report` explains triggered rules with details such as:

- employee
- week
- rule triggered
- trigger date
- actual vs limit
- daily ST / OT / DT for the trigger day
- source files involved
- readable reason
- daily breakdown

## Formatting Rules

Formatting profiles are saved locally and can vary by job.

Current supported rule fields:

- `Daily ST Alert`
- `Weekly ST Alert`
- `Weekly OT Alert`
- `Max Hours Per Day`

Leave a field blank to disable that rule.

## Email Drafts

`Email Drafts` can:

- preview a selected week by employee
- create one Outlook draft per employee
- use editable saved subject and body templates
- greet each employee by first name
- create drafts only for selected employees in the preview table, or everyone in the week if nothing is selected

Template placeholders:

- `{employee}`
- `{first_name}`
- `{week_start}`
- `{week_end}`
- `{hours_table}`

Employee email addresses can be:

- entered manually
- edited by double-clicking an employee from supported lists
- filled from desktop Outlook for missing names

## Employee Lists and Groups

- `Employee List` stores unique names detected from the loaded DSS data
- Employee names are grouped by exact spelling
- `Employee Groups` lets you build reusable crews or custom sets
- Group member selection supports normal Windows multiselect behavior:
  - `Ctrl` for individual add/remove
  - `Shift` for range select

## Configuration

The `Configuration` page currently includes:

- disable name typo notifications
- check source DSS(s) frequency in minutes
- show/hide the `Daily Raw` tab
- enable automatic update checks against GitHub releases
- optionally download release installers automatically on unmetered Wi‑Fi
- **Appearance:** eight configurable `#RRGGBB` colours (with **Pick…** dialogs) for alert table rows, crew-total rows, formatting-rule tooltips, and the Reports group outline when errors or parse warnings exist; **Reset colours to sample defaults** restores the built-in palette

### Maintenance Buttons

- `Reset All Settings to Default`
- `Clear Cached DSSs`
- `Clear Stored Emails`
- `Clear All Stored Data`

### Diagnostic Buttons

- `Show App Data Folder`
- `Export Diagnostic Snapshot`
- `Test Outlook Connection`
- `Show Loaded DSS Status`

## Outlook Integration

The app uses desktop Outlook through `pywin32` when available.

Supported Outlook features:

- sync missing employee email addresses
- create Outlook draft emails
- test Outlook connectivity from Configuration

The app also performs one automatic Outlook email lookup about one minute after startup when DSS data is loaded.

If Outlook cannot resolve a name and the name looks similar to another loaded employee, the app can warn about a possible typo unless that notification is disabled in Configuration.

## App Data

On Windows the app stores its data in:

```text
%LOCALAPPDATA%\DSSHoursTracker
```

This includes:

- config
- formatting profiles
- employee emails
- employee groups
- email templates
- table layouts
- parsed DSS cache
- diagnostic snapshot exports

## Packaging

### GitHub Actions (automated Windows release)

The workflow **`.github/workflows/release-windows.yml`** runs on:

- **Push of a version tag** matching `v*` (for example `v0.2.0`), or  
- **Manual run** (“Run workflow”) with an **existing** tag name to rebuild and re-upload assets for that tag.

It builds **`dist/DSSHoursTracker.exe`** with PyInstaller (bundles `pywin32`), embeds the tag’s numeric version via **`dss_app_version.txt`** (so in-app “installed version” and GitHub update checks work), writes **`checksums.txt`** for the app’s optional SHA-256 verification, and publishes both files on the **GitHub Release** for that tag (`softprops/action-gh-release`).

**Repository settings:** **Settings → Actions → General → Workflow permissions** → enable **Read and write permissions** (and allow GitHub Actions to create and approve pull requests if your org defaults to read-only), so `GITHUB_TOKEN` can attach release assets.

**Before tagging:** bump **`pyproject.toml`** `[project] version` to match the release when you want the repo metadata and frozen bundle to agree (the workflow still stamps the exe from the **tag** for `discover_app_version()`).

**Commands (maintainer):**

```bash
git tag -a v0.2.0 -m "Release 0.2.0"
git push origin v0.2.0
```

### Local PyInstaller build

Optional one-file build (version will follow `pyproject.toml` or installed package unless you set **`DSS_APP_VERSION`** or ship **`dss_app_version.txt`** next to the spec / add it with `--add-data` as in the workflow):

```powershell
pyinstaller --noconsole --onefile dss_hours_tracker.py
```
