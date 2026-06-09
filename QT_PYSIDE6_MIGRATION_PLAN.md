# Qt / PySide6 Migration Plan

## Goal

Rebuild the desktop interface on `PySide6` while preserving the existing DSS parsing, aggregation, caching, Outlook, export, and configuration logic already implemented in `dss_hours_tracker.py`.

This is intentionally a **new shell around the existing core**, not a simultaneous rewrite of the business rules.

## Why Migrate

The current `tkinter` UI has reached the point where:

- large table redraws and long scrollable settings forms visibly fragment during heavy loading
- progressive preview updates compete with scroll painting
- the single-threaded Tk widget tree is carrying a lot of UI complexity

`PySide6` gives us:

- better table/view infrastructure
- smoother scrolling and repaint behavior
- more predictable threaded UI updates through signals/slots
- a cleaner path for large-table virtualization and incremental refresh

## Scope

### Keep

- DSS workbook reading and parsing rules
- semantic DSS hashing and cache reuse
- revision sheet selection logic
- Monday-Sunday aggregation
- error finding logic
- Outlook email resolution and draft creation
- config file storage and app data directories
- updater / installer cleanup entry handling

### Replace

- all `tkinter` windows, tabs, tables, dialogs, menus, and scroll containers
- Tk-specific progress/update scheduling
- Tk-specific layout persistence / column menu plumbing

## Architecture

### Existing Backend

`dss_hours_tracker.py` remains the backend/core for now and continues to own:

- data models (`DailyRecord`, `TrackerData`, `FormattingProfile`, etc.)
- loading / parsing / aggregation
- export helpers
- Outlook helpers
- config persistence helpers
- diagnostics and maintenance helpers

### New Frontend

Add a new module:

- `dss_qt_app.py`

This module becomes the PySide6 shell and owns:

- main window
- grouped navigation
- Qt table models / proxy filtering
- background worker wiring
- dialogs / forms / list editing
- Qt-specific layout persistence

## Migration Phases

### Phase 1: Shell + Infrastructure

- Add `PySide6` dependency and packaging support
- Create `dss_qt_app.py`
- Build:
  - main window
  - toolbar / top actions
  - grouped two-row navigation
  - status / progress / cancel area
  - reusable Qt table view wrapper
  - global employee / PF filters
- Keep backend loading in `dss_hours_tracker.py`

### Phase 2: Loading and Core Views

- Wire `load_tracker_data(...)` through a background worker
- Support:
  - open DSSs
  - add DSS
  - remove DSSs
  - update view
  - cancel
- Port core data/report tabs:
  - `Daily Raw`
  - `Daily by PF`
  - `Weekly by PF`
  - `Combined Summary Daily`
  - `Combined Summary Weekly`
  - `Week Totals`
  - `Sheet Parse Warnings`
  - `Workbook Health`
  - `Audit Data Trail`

### Phase 3: Workflows and Settings

- Port:
  - formatting rules / profiles
  - job presets
  - employees / groups / notes / suppression
  - email templates / preview / Outlook draft creation
  - error reports
  - maintenance / diagnostics actions

### Phase 4: Entry Point and Cleanup

- Switch `main()` to launch Qt
- Preserve installer cleanup mode
- Keep Tk code in place as backend-compatible legacy code until the Qt shell settles
- Update docs and development log

## View Mapping

### Top-Level Groups

- `Data`
- `Summaries`
- `Reports`
- `Settings`

### Group Pages

#### Data

- `Daily Raw`
- `Week Totals`

#### Summaries

- `Daily by PF`
- `Weekly by PF`
- `Combined Summary Daily`
- `Combined Summary Weekly`

#### Reports

- `Error Report`
- `Email Drafts`
- `Sheet Parse Warnings`
- `Workbook Health`
- `Audit Data Trail`

#### Settings

- `Configuration`
- `Employees`
- `Formatting Rules`

## Performance Strategy

The Qt rewrite should not try to mimic the old Tk repaint behavior.

Key principles:

- only update visible views during loading
- keep data transforms off the UI thread
- use model/view tables instead of manually inserting widget rows
- batch table refreshes
- avoid rebuilding hidden tabs during progress updates
- preserve progressive loading, but deliver it through lightweight model refreshes

## Risks

- feature parity breadth is large
- Outlook integration remains Windows-specific
- packaging will need PyInstaller / installer updates after the shell swap
- some Tk-only layout persistence details will need a Qt-native equivalent

## Success Criteria

- Qt app launches from the existing entry point
- existing DSS parsing behavior is unchanged
- multi-DSS loads, filters, reports, and email drafting continue to work
- UI remains responsive during heavy loads
- scrolling and large table redraws feel materially smoother than Tk

## Rollback Strategy

Before switching `main()` to the Qt shell, create a Git rollback point.

If needed, rollback is straightforward because:

- backend/core logic stays in place
- the Qt shell is introduced as a separate module
- the entry point swap is isolated
