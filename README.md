# DSS Tools

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

## Application icon (Windows)

**Why shortcuts / taskbar showed a generic icon:** the GitHub release build only picks up an icon if **`dss_tools.ico`** exists at the repo root when CI runs. A root **`DSS-Tools Icon.png`** was ignored by `.gitignore` (`*.png`), so it never reached Actions—PyInstaller had no `--icon` and Inno had no `SetupIconFile`.

**What to do:** track **one** of these at the repository root (they are un-ignored in `.gitignore`):

- **`dss_tools.ico`** (best for Windows—include 16–256 px sizes), or  
- **`DSS-Tools Icon.png`** (preferred), **`DSS Tools Icon.png`** (spaces), **`DSS-Tools-Icon.png`**, **`app-icon.png`**, **`dss_tools.png`**, or **`DSSTools Icon.png`** (CI uses the first existing file in that order and converts PNG → multi-size `dss_tools.ico` with Pillow), or  
- **exactly one** other `*.ico` file in the root (it is copied to `dss_tools.ico`).

If none of those are present, CI copies a small **placeholder** from `tools/default_dss_tools.ico` so the installer still gets a real icon—replace it with your artwork using one of the options above.

When you run **`python dss_hours_tracker.py`**, the window looks for, in order: **`dss_tools.ico`**, `DSSTools.ico`, `app_icon.ico`, `icon.ico`, `app.ico`, then a **single** root `*.ico`. Before packaging you can run **`python tools/ensure_dss_tools_ico.py`** to refresh `dss_tools.ico` from a tracked PNG.

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
dss-tools
```

## Tests

Workbook parsing, `load_tracker_data`, disk cache, and related cases live in `test_dss_tools_integration.py`. That module’s class is **skipped by default** so day-to-day runs stay quick.

**Fast suite** (unit tests and JSON round-trips only; integration skipped):

```powershell
python -m unittest test_dss_tools -q
```

**Full suite** (include integration; allow roughly a minute on a typical machine):

```powershell
$env:RUN_SLOW_TESTS = "1"
python -m unittest discover -s . -p "test_dss*.py" -q
```

`RUN_SLOW_TESTS` is treated as enabled when its value is `1`, `true`, `yes`, or `all` (case-insensitive). Shared workbook helpers for both modules are in `test_dss_tools_fixtures.py`.

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
- Optional **quick load**: on startup the app can re-open the same DSS workbook paths as last time (saved in config), with a short hint to the **left** of the progress bar on the toolbar row; turn the behaviour off in Configuration
- Multi-file loads prefer **newer files first** (by filesystem modified time); the parser can emit an early partial preview while older workbooks are still loading
- Same-file updates are skipped when the **DSS semantic hash** (dated sheets and `K25:AZ36` only) is unchanged; per-sheet digests allow **partial refresh** when only some dated sheets change
- Parsed DSS data is cached on disk for up to 7 days
- Optional **GitHub release** update check; automatic download on unmetered Wi‑Fi; **Check for Updates** can download on other networks when you confirm; after download, **Install now** launches **`DSSToolsUpdater.exe`** from **`%TEMP%`** (staged from the install or embedded copy) so the old **`Program Files`** tree can be removed safely. The helper opens a **compact window** with a **percentage progress bar** (**0–50%** uninstall, **50–100%** install—estimated from elapsed time because Inno is silent), then **silent Inno uninstall** (same product id), clears **`%LOCALAPPDATA%\DSSTools`** **cache**, **updates**, logs, and diagnostic JSON (not **`dss_hours_tracker_config.json`**), runs **`DSSToolsSetup.exe`** with **`/VERYSILENT`** (reusing the previous install directory via **`UsePreviousAppDir`**), and starts **DSS Tools** when finished—if silent install fails, it offers the **full wizard**. Current frozen builds **embed** the helper; Inno also ships it beside **`DSSTools.exe`**. Very old builds may still use a hidden PowerShell handoff and **`update_handoff.log`**.
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
- **Application version** line (same value as the built-in / frozen version used for updates)
- Long **Settings** subtabs and **Email Drafts** use a vertical scrollbar when the window is too short to show all controls
- **Appearance:** configurable `#RRGGBB` colours (with **Pick…** next to each hex field) for alert rows, crew totals, tooltips, **main content background**, and **table cell background**; report tabs use the alert-row tint when issues exist; **Reset colours to sample defaults** restores the built-in palette

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
- store each resolved employee’s **Outlook display name** (address book / GAL) next to their email for **DSS roster vs address book** checks
- create Outlook draft emails
- test Outlook connectivity from Configuration

The app also performs one automatic Outlook email lookup about one minute after startup when DSS data is loaded.

If Outlook cannot resolve a name and the name looks similar to another loaded employee, the app can warn about a possible typo unless that notification is disabled in Configuration. After a successful sync, if the **DSS spelling** still differs from the **resolved Outlook display name** (for example a one-letter surname typo), **Check Name Typos** and the **Error Report** can flag those mismatches with the **trigger date** and **source file** so you can fix the roster.

## App Data

On Windows the app stores its data in:

```text
%LOCALAPPDATA%\DSSTools
```

If you already have data from an older build, the app may still be using your existing folder under `%LOCALAPPDATA%` until you move files manually; use **Show App Data Folder** in Configuration to see the active path.

This includes:

- config
- formatting profiles
- employee emails
- Outlook display names (paired with saved emails for typo checks)
- employee groups
- email templates
- table layouts
- parsed DSS cache
- diagnostic snapshot exports

## Packaging

### GitHub Actions (automated Windows release)

The workflow **`.github/workflows/release-windows.yml`** runs on:

- **Push of a version tag** matching **`MAJOR.MINOR.PATCH`** (for example `0.2.0`) or the same with a leading **`v`** (`v0.2.0`), or  
- **Manual run** (“Run workflow”) with an **existing** tag name to rebuild and re-upload assets for that tag.

It builds **`dist/DSSTools.exe`** with PyInstaller (bundles `pywin32`), embeds the tag’s numeric version via **`dss_app_version.txt`**, compiles **`installer/DSSTools.iss`** with **Inno Setup** into **`dist/DSSToolsSetup.exe`** (click-through wizard, Start-menu entry, uninstaller in **Settings → Apps**), writes **`checksums.txt`** for the **setup** program’s SHA-256 (for the app’s optional download verification), and publishes **`DSSToolsSetup.exe`** + **`checksums.txt`** on the **GitHub Release** (`softprops/action-gh-release`). The raw PyInstaller `.exe` is not attached to the release so the in-app updater picks the installer asset.

**Repository settings:** **Settings → Actions → General → Workflow permissions** → enable **Read and write permissions** (and allow GitHub Actions to create and approve pull requests if your org defaults to read-only), so `GITHUB_TOKEN` can attach release assets.

**Before tagging:** bump **`pyproject.toml`** `[project] version` to match the release when you want the repo metadata and frozen bundle to agree (the workflow still stamps the exe from the **tag** for `discover_app_version()`).

**Commands (maintainer):**

```bash
git tag -a 0.2.0 -m "Release 0.2.0"
git push origin 0.2.0
```

### Local build (PyInstaller + Inno wizard)

0. **Icon:** The repo ships **`DSS-Tools Icon.png`** and generated **`dss_tools.ico`** at the root. **PyInstaller** embeds that ICO into **`DSSTools.exe`** (`--icon`); the **Inno wizard** still uses **`SetupIconFile`**. Start-menu and desktop shortcuts use **`DSSTools.exe`** as **`IconFilename`** so the shell shows the **same** embedded resource the taskbar uses (a correct loose **`dss_tools.ico`** alone cannot fix a bad embedded icon). To refresh the ICO after editing the PNG: `pip install pillow` then `python tools/ensure_dss_tools_ico.py --force`. Without sources, the script falls back to **`tools/default_dss_tools.ico`** (generic blue tile).

1. **One-file app** (same flags as CI; version from tag in CI is simulated here with `dss_app_version.txt`):

```powershell
Set-Content -NoNewline dss_app_version.txt "0.1.0"
$args = @("--noconsole", "--onefile", "--name", "DSSTools", "--collect-all", "pywin32", "--add-data", "dss_app_version.txt;.")
if (Test-Path -LiteralPath "dss_tools.ico") { $args += @("--icon", "dss_tools.ico", "--add-data", "dss_tools.ico;.") }
pyinstaller @args dss_hours_tracker.py
```

2. **Installer** (requires [Inno Setup 6](https://jrsoftware.org/isdl.php) installed):

```powershell
& "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe" /DMyAppVersion=0.1.0 installer\DSSTools.iss
```

Output: **`dist\DSSToolsSetup.exe`**. Uninstall: **Settings → Apps → Installed apps → DSS Tools**.
