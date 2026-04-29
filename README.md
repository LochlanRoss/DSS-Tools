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
    - `Weekly Rollup`
    - `Combined Summary`
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
- Overall progress bar shows multi-DSS progress
- `Cancel` can abort DSS loading and Outlook email sync
- Same-file updates are skipped when the workbook hash is unchanged
- Parsed DSS data is cached on disk for up to 7 days
- Loaded DSS files are checked periodically in the background for changes
- Revision sheets such as `R1` / `rev 2` override the non-revised sheet for the same date
- Conditional highlighting uses saved job-specific formatting profiles
- Table layouts persist per page:
  - shown/hidden columns
  - column widths
  - sort column and direction

## Filtering

The top `Filter` control is a checklist popup:

- `All Employees`
- `Uncheck All`
- one checkbox per employee

You can keep the filter popup open while selecting multiple employees. It closes when you click away or press `Esc`.

The filter applies across the main data views and the email draft preview.

## Weekly Views

- `Weekly Rollup`
  - one row per employee per source DSS per week
  - includes `Whole Crew` total rows
- `Combined Summary`
  - combines matching employee names across all loaded DSS files
- `Week Totals`
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

Optional PyInstaller build:

```powershell
pyinstaller --noconsole --onefile dss_hours_tracker.py
```
