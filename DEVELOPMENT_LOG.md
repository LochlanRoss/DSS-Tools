# Development Log

This log tracks moderate and larger changes, along with active feature requests, notable bug work, and handoff-ready context.
Small fixes and very short edits are intentionally omitted unless they materially changed behavior.

## 2026-05-04

### Branding PNG path (DSS-Tools vs DSS Viewer)

- **Note:** ``DSS Tools Icon.png`` was authored under the sibling folder ``GitHub\DSS-Tools``; this app’s repo is ``GitHub\DSS Viewer``. The file was copied to the **DSS Viewer** root so ``ensure_dss_tools_ico.py`` could regenerate ``dss_tools.ico``. Keep the canonical PNG in **DSS Viewer** for CI and PyInstaller.

### Icon PNG: multiple tracked root names

- **Change:** ``tools/ensure_dss_tools_ico.py`` picks the first existing file from ``BRAND_PNG_PRIORITY`` (``DSS-Tools Icon.png``, ``DSS Tools Icon.png`` with spaces, ``DSS-Tools-Icon.png``, ``app-icon.png``, etc.). **``.gitignore``** and **release CI** accept the same set so a replacement artwork filename still ships and drives ``dss_tools.ico``.

### Windows taskbar / desktop icon vs Explorer (Tk + Shell)

- **Cause:** The taskbar often ignored Tk ``iconbitmap`` and showed the generic Tcl/Python glyph; pinned shortcuts could group separately from the running process without a consistent **App User Model ID**.
- **Fix:** Frozen builds call ``SetCurrentProcessExplicitAppUserModelID`` with **`WIN_APP_USER_MODEL_ID`** before creating the root window; after map, **WM_SETICON** loads icon resource **id 1** from the frozen exe (matches PyInstaller ``--icon``). Inno **[Icons]** sets the same string via **`AppUserModelID`** on Start-menu and desktop shortcuts so Shell and runtime agree.

### Installer: clean upgrades + personal Desktop + AppData scrub

- **Desktop shortcut:** Inno now uses **`{userdesktop}`** (installing user’s profile Desktop) instead of **`{autodesktop}`**, which under **`PrivilegesRequired=admin`** targeted **Public Desktop** only—File Explorer’s personal Desktop folder showed no shortcut. **`[InstallDelete]`** removes a legacy **`{commondesktop}\DSS Tools.lnk`** on upgrade.
- **Program Files:** **`[InstallDelete]`** deletes **`{app}`** before copying files so orphaned artifacts from older builds cannot remain beside the new exe set.
- **%LOCALAPPDATA%\\DSSTools:** **`clean_transient_app_data`** now removes **everything** except **`dss_hours_tracker_config.json`** (cache, updates, logs, staged updater copy, stray files/dirs). The updater already calls this before silent reinstall; the wizard installer runs **`DSSTools.exe --installer-postinstall-cleanup`** (hidden, wait) after install so interactive upgrades match.
- **Temp:** After the GUI updater finishes, **`cleanup_staged_temp_artifacts`** schedules deletion of **`dss_tools_setup_*`** / **`dss_tools_updater_*`** copies under **`%TEMP%`** via a short **`cmd`** delay (cannot delete a running exe immediately).
- **Build:** PyInstaller for **`DSSTools`** includes **`--hidden-import dss_tools_updater`** (and **`DSSTools.spec`** lists it) so the cleanup flag works in the frozen exe.

### Inno shortcuts: desktop vs taskbar icon mismatch

- **Cause:** Shortcuts used **``{app}\dss_tools.ico``** while the taskbar / pinned app icon usually comes from the **first icon resource embedded in ``DSSTools.exe``** (PyInstaller ``--icon``). After fixing the repo ICO, the loose file could look fine while an older **PE-embedded** multi-size set still looked like a solid smear.
- **Fix:** Start-menu and desktop **``IconFilename``** now point at **``{app}\DSSTools.exe``** (same as **``UninstallDisplayIcon``**). The wizard still uses **``SetupIconFile``** = ``dss_tools.ico``. **``dss_tools.ico``** is still installed beside the exe for in-app / ``_MEIPASS`` copy paths.

### ICO script: installer / shell showed solid blue smear

- **Cause:** Pillow’s ICO ``_save`` skips any requested ``sizes`` larger than the **first** image’s width/height. The script passed a **16×16** frame first, so only 16×16 was written; Windows/Inno scaled that up to a flat color. A non-square PNG could also yield a **256×255** frame. ``append_images`` alone did not fix ordering.
- **Fix:** ``tools/ensure_dss_tools_ico.py`` now letterboxes to a square, builds all sizes, and saves with the **256×256** image as the primary plus ``append_images`` for the rest. After write, a **mean RGB error** check compares a 64×64 fit of the PNG to the decoded ICO (tunable ``--no-verify`` to skip).

### Updater: hide flashing console windows

- **Cause:** ``CREATE_NO_WINDOW`` alone was not always enough when a GUI process spawned ``taskkill``, ``cmd /c …``, or Inno from ``subprocess``; a brief console host could appear.
- **Fix:** Shared ``_windows_hidden_subprocess_kwargs()`` adds ``STARTUPINFO`` (``STARTF_USESHOWWINDOW`` + ``SW_HIDE``), null stdio for ``Popen``, and keeps ``DETACHED_PROCESS`` on the delayed temp-file delete helper.

### Updater: Inno uninstall registry miss

- **Cause:** The helper only opened the fixed subkey ``{AppId}_is1``. Older or alternate builds can register a different subkey name; a **portable** ``DSSTools.exe`` tree has **no** Uninstall entry at all. The log line *“No existing Inno uninstall entry”* was easy to read as “Inno is broken” when it really meant “expected key not found.”
- **Fix:** :func:`discover_inno_dss_tools_uninstall_info` tries the canonical AppId subkey first, then **enumerates** ``HKLM``/``HKCU`` ``…\\Uninstall`` for Inno-style ``unins*.exe`` strings plus DSS Tools ``DisplayName`` / ``Publisher`` / ``InstallLocation`` / path heuristics. :func:`read_inno_install_location` reuses the same discovery. Clearer log text when nothing matches.

### Update helper: stop stalling on "waiting for main app to close"

- **Cause:** The GUI polled ``parent_process_exists_windows``, which treated ``OpenProcess`` **ERROR_ACCESS_DENIED (5)** as "parent still running." After the main window closed, **PID reuse** or **elevated updater vs. medium-IL target** could keep returning 5 so the loop never finished.
- **Fix:** Replace the wait loop with ``taskkill /PID … /T /F`` (process tree), a short sleep, then uninstall/install. Same for ``--headless`` legacy handoff.
- **Follow-up:** ``taskkill`` runs only after ``QueryFullProcessImageName`` confirms the PID still maps to the **same executable path** recorded at handoff (or ``DSSTools.exe`` when no path is passed). If the main app already exited, ``taskkill`` is skipped so a recycled PID is never killed. The main app still exits only via Tk ``destroy()`` after starting the sidecar; there is no ``taskkill`` in ``dss_hours_tracker``. The handoff passes the **live** process image path (not only ``sys.executable``) so Store Python shims match the real interpreter.

### Update helper: UAC elevation (WinError 740)

- **Cause:** ``DSSToolsUpdater.exe`` is built with PyInstaller **``uac_admin``** (administrator manifest). Starting it with **``subprocess.Popen``** from a normal session fails with **WinError 740** (``ERROR_ELEVATION_REQUIRED``) because no UAC handoff occurs.
- **Fix:** On Windows, launch the frozen updater via **``ShellExecuteW``** with the **``runas``** verb so the standard elevation consent prompt appears. Dev handoff (**``python``** + **``dss_tools_updater.py``**) still uses **``subprocess``** (no admin manifest).
- **Follow-up (0.1.10 still saw 740):** The main app **stages** the updater to ``%%TEMP%%`` as **``dss_tools_updater_*.exe``**. The elevation branch only matched **``DSSToolsUpdater.exe``**, so the staged path fell through to **``Popen``** and still hit 740. **``_is_updater_executable_path``** now treats the temp prefix like the shipped name.

### App icon committed (no CI placeholder)

- **Cause:** Root `*.png` was gitignored until recently; clean clones had no **`DSS-Tools Icon.png`**, so **`tools/ensure_dss_tools_ico.py`** fell through to **`tools/default_dss_tools.ico`** (generic blue tile) for **`dss_tools.ico`**, PyInstaller, Inno, and the Tk title bar.
- **Fix:** Added **`DSS-Tools Icon.png`** (timesheet artwork) and regenerated **`dss_tools.ico`** in the repo root. Release workflow runs **`ensure_dss_tools_ico.py --force`** so the ICO always tracks the PNG when both are present. Script **`--force`** regenerates even if **`dss_tools.ico`** already exists.

### Icon tests + updater exe branding (follow-up)

- **Tests:** Canonical ICO tests used a per-pixel gradient PNG; post-write **MAE** verification correctly failed (>32). Tests now use **flat 256×256** canonical PNGs and an in-test **stray.ico** (no dependency on missing **`tools/default_dss_tools.ico`**).
- **CI / spec:** **`DSSToolsUpdater`** PyInstaller step now passes **`--icon`** / **`--add-data`** for **`dss_tools.ico`** (matches main exe). **`DSSTools.spec`** mirrors this for local **`pyinstaller DSSTools.spec`** builds.

### DSS roster vs Outlook address book (Error Report + name typo checks)

- **Cause:** Name typo detection only compared **unresolved** DSS names to other DSS names. Once Outlook filled an email, the roster could still spell the person differently than the address book (e.g. Kolodinski vs Kolodinsky) with no warning.
- **Change:** Outlook sync now stores **`employee_outlook_display_names`** in config (resolved display name when SMTP is found). **`find_outlook_display_name_typos`** / **`build_outlook_name_mismatch_findings`** compare normalized DSS text to that display name (difflib floor **0.84**); ignored pairs reuse **`ignored_name_typos`** / **`typo_warning_key`**. **Error Report** adds rows with rule **“Name does not match email address book”**, **Trigger Date**, **Source Files** (workbook), day ST/OT/DT for that day+file, and narrative in **Reason** / **Daily Breakdown**. **`ErrorFinding.outlook_name_rule`** drives non-hour formatting in the Actual/Limit/Delta columns. Manual **Check Name Typos** always includes address-book mismatches when display names exist. Editing or clearing an employee email drops the cached Outlook display name for that person.
- **Tests:** `test_dss_tools.py` (round-trip, preserve-on-email-only-save, Kolodinski vs Kolodinsky findings and typo list).

### In-app update mini-app (silent uninstall / clean / reinstall)

- **`dss_tools_updater.py`** is now a compact **Tk** window after **Install now**: waits for the main PID, runs **Inno silent uninstall** (registry `UninstallString` for the fixed `AppId`), clears **`%LOCALAPPDATA%\DSSTools`** transients (**`cache/`**, **`updates/`**, `*.log`, `diagnostic_snapshot_*.json`) while keeping **`dss_hours_tracker_config.json`**, then runs **`DSSToolsSetup.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /CLOSEAPPLICATIONS`**. A **determinate** progress bar shows **0–50%** during uninstall and **50–100%** during install (time-smoothed while each subprocess runs; Inno does not expose byte-level progress). On silent failure, offers the **full wizard**. Launches **`DSSTools.exe`** from `InstallLocation` when done.
- **Main app** stages the **installer** into `%TEMP%` when it was under app data (so **`updates/`** can be removed) and stages **`DSSToolsUpdater.exe`** into `%TEMP%` so uninstall can remove the old **`Program Files`** tree while the helper is still running.
- **Inno** `[Setup]`: **`CloseApplications=yes`**, **`UsePreviousAppDir=yes`**, **`RestartApplications=no`** for smoother silent upgrades.
- **Build:** **`DSSToolsUpdater`** frozen with **`--uac-admin`** and **`--collect-all tkinter`** (spec mirrors). **`--headless`** keeps the old wait+`startfile` behaviour for tests.

### Windows installer / taskbar / desktop icon missing on GitHub builds

- **Cause:** Root **`*.png`** was gitignored, so **`DSS-Tools Icon.png`** never appeared in CI; **`dss_tools.ico`** was absent, so PyInstaller had no **`--icon`** and Inno skipped **`HasAppIco`** (shortcuts and uninstall used the generic executable look).
- **Fix:** **`tools/ensure_dss_tools_ico.py`** builds **`dss_tools.ico`** from a lone root **`.ico`**, named PNGs (**`DSS-Tools Icon.png`**, etc.), or a single *icon* PNG; otherwise copies **`tools/default_dss_tools.ico`**. Release workflow runs this before PyInstaller. **`.gitignore`** now allows those PNG filenames to be tracked.
- **Maintainer:** commit **`DSS-Tools Icon.png`** or **`dss_tools.ico`** at the repo root so releases use your branding instead of the placeholder.

### Reliable in-place updates (sidecar updater)

- **Cause:** Post-download **Install now** spawned a hidden PowerShell wait loop, then closed the app; failures were silent (stdout/stderr discarded), so users often saw the window close with no installer.
- **Fix:** Added **`dss_tools_updater.py`** / frozen **`DSSToolsUpdater.exe`**: waits on the parent PID via **`OpenProcess`** (locale-independent), then launches the downloaded setup with **`os.startfile`**. The main app prefers this sidecar when present beside **`DSSTools.exe`** (or runs the script via **`sys.executable`** in dev); otherwise falls back to the old PowerShell handoff with **`%LOCALAPPDATA%\DSSTools\update_handoff.log`** for command output.
- **Packaging:** **`DSSTools.spec`**, **`installer/DSSTools.iss`**, and **`.github/workflows/release-windows.yml`** now build and ship **`DSSToolsUpdater.exe`** into the install directory.
- **Follow-up:** Legacy PowerShell handoff showed an empty console and often failed silently; it now uses **`CREATE_NO_WINDOW`**, **`STARTUPINFO` / `SW_HIDE`**, and **`stdout/stderr` to `DEVNULL`**. The frozen **`DSSTools.exe`** build **embeds** **`DSSToolsUpdater.exe`** (`--add-binary` / spec `binaries`); at install time the app **copies** it from **`_MEIPASS`** to **`%LOCALAPPDATA%\DSSTools\`** before exit so the helper still runs after PyInstaller deletes its temp folder (portable single-file use without Inno’s second file).
- **UX:** Manual **Check for Updates** when not on unmetered Wi‑Fi now offers **Download the installer now?** instead of only linking the release page.
- **Tests:** **`test_dss_tools_updater.py`** (Windows-only PID checks; skipped elsewhere).

### Bug report Outlook attachment & diagnostics shortcuts

- **Bug:** Submit Bug Report failed with Outlook **“Cannot find this file”** (`0x80070002`) when attaching the JSON snapshot; COM is strict about path shape and some `%LOCALAPPDATA%` paths.
- **Fix:** **`Attachments.Add(Source=…, Type=olByValue)`**, try **`GetShortPathName`** on Windows when available, then a **short-named copy under `%TEMP%`** before giving up; **`Save()`** still runs and the UI reports a **warning** with on-disk snapshot path if Outlook refuses the attachment.
- **Settings:** Diagnostics frame adds **Sync Outlook Emails** and **Check Name Typos** (same actions as elsewhere); manual sync now explains **load DSS first** or **operation in progress** instead of doing nothing silently.
- **Tests:** **`test_bug_report_attachment_strings_to_try`** in **`test_dss_tools.py`**.

## Logging Rule
- Record moderate or large feature work, behavioral changes, UI workflow changes, architecture changes, persistence changes, diagnostics, packaging, caching, and integration work.
- Do not record tiny cosmetic edits or very small bug fixes unless they changed user behavior in an important way.

## 2026-05-01

### Product rename: DSS Tools

- User-facing name is **DSS Tools** (`DISPLAY_APP_NAME`); window title, dialogs, bug-report copy, argparse description, and HTTP **User-Agent** use **`dss-tools/`** plus the app version string.
- **PyInstaller / Inno / CI:** executable **`DSSTools.exe`**, installer script **`installer/DSSTools.iss`**, published asset **`DSSToolsSetup.exe`**; workflow **`.github/workflows/release-windows.yml`** updated accordingly.
- **Distribution:** **`pyproject.toml`** project name **`dss-tools`**, GUI entrypoint **`dss-tools`**; **`discover_app_version()`** tries **`dss-tools`** package metadata first, then any older registered distribution name for the same codebase so editable or transitional installs still resolve a version.
- **Data directory:** default **`%LOCALAPPDATA%\DSSTools`** (`APP_DIRNAME`); if the previous per-user folder already exists it is reused automatically (`LEGACY_APP_DIRNAME`) so existing installs keep their config and cache until migrated manually.
- **Updates:** **`GITHUB_REPO_SLUG`** is **`LochlanRoss/DSS-Tools`** (rename the GitHub repository to match, or adjust the constant if releases stay on another slug).
- Removed obsolete root Inno script (superseded by **`installer/DSSTools.iss`**). PyInstaller spec renamed to **`DSSTools.spec`** with **`name='DSSTools'`**.
- **Test modules:** **`test_dss_tools.py`** (fast), **`test_dss_tools_integration.py`** (slow, gated), **`test_dss_tools_fixtures.py`** (shared helpers); fast CI command is **`python -m unittest test_dss_tools -q`**.
- **Windows app icon:** repo-root **`dss_tools.ico`** (or a single **`*.ico`** picked up in dev) drives **`apply_tk_window_icon`** on the main **`tk.Tk`**, **PyInstaller** `--icon` / `--add-data`, **`DSSTools.spec`** bundling, and **`installer/DSSTools.iss`** wizard + shortcut icons via Inno preprocessor **`#if FileExists`** when that file is present; release workflow normalizes a lone **`*.ico`** to **`dss_tools.ico`** before the frozen build.

### Scrollable Settings / Email Drafts pages

- **VerticalScrollablePage** (`Canvas` + `ttk.Scrollbar` + inner frame) wraps long notebook tabs so a vertical scrollbar appears when the window is short: **Configuration**, **Employee List**, **Employee Notes**, **Employee Groups**, **Formatting Rules**, and **Reports ? Email Drafts**; mouse wheel scrolls the page except on `Text` / `Listbox` / `Treeview` (those keep local scroll).
- **Employee Groups**: vertical scrollbar on the **Groups** listbox; **Employee Notes**: scrollbar on the names listbox and on the note **Text**; **Email Drafts**: scrollbars on subject and body **Text** fields.

### Main-window chrome, cache clear, reports tabs, toolbar UX

- **Clear Cached DSSs** now also clears in-memory reuse flags (`file_hashes`, `reused_paths`, cache status to Miss) for the loaded set so stats and the next **Update View** do not report stale memory hits after disk cache deletion (`tracker_data_invalidated_for_cache_clear`).
- **Open DSS** removed; a single **Add DSS Workbook(s)** control always merges paths; **Remove DSS(s)** clears sources as before.
- **Reports alerts:** dropped the pink outline frame; Error Report / Sheet Parse Warnings / parent **Reports** tabs show a left **swatch** in the alert-row background colour (PhotoImage stripe) while keeping the `(!)` labels.
- **Toolbar row:** hint sits left of the progress bar; **Cancel** aligns on the same row as the bar; standard ttk toolbar and progress row (no dark strip) (`UiThemeColors` **table_background**, **content_chrome_background**); Configuration Appearance lists those fields with **Pick�** beside each hex entry; **Application version** line on Configuration.
- Shared **`DssTable.Treeview`** style for table backgrounds; word-wrap styles copy table colours.

### Test layout: fast default vs slow integration

- Split workbook-heavy coverage into `test_dss_tools_integration.py`, gated with **`RUN_SLOW_TESTS`** (`1` / `true` / `yes` / `all`) and `@unittest.skipUnless` on the integration class so **`python -m unittest test_dss_tools`** stays fast.
- Shared temp-file / sample-workbook helpers live in **`test_dss_tools_fixtures.py`** (`DssToolsFixtures`) to avoid duplicating setup between modules.
- **`test_dss_tools.py`** keeps sort keys, pure helpers, settings/layout round-trips, updater helpers, and similar tests without requiring the env var.
- Full DSS test discovery remains **`python -m unittest discover -s . -p "test_dss*.py"`**; with `RUN_SLOW_TESTS=1`, integration tests execute instead of skipping (still **71** tests total: **43** fast + **28** slow).
- **`test_dss_tools.py`** was trimmed so workbook/cache/`load_tracker_data` cases exist only in the integration module (no duplicate `test_*` definitions); imports were reduced to match the slimmer fast suite. Integration module imports include helpers such as **`format_email_subject`** and **`compute_bytes_hash`** where cache and email tests need them.
- **`README.md` Tests** subsection documents the split, the env var, fixture module path, and PowerShell examples for fast-only vs full discovery runs.

### GitHub Actions: Windows release workflow

- Added **`.github/workflows/release-windows.yml`**: on **semver-style tag push** (`[0-9]*.[0-9]*.[0-9]*` or `v[0-9]*.[0-9]*.[0-9]*`) or manual dispatch for an existing tag, checks out the ref, writes **`dss_app_version.txt`**, runs fast unit tests, runs **PyInstaller** `--onefile` with **`--collect-all pywin32`**, compiles **`installer/DSSTools.iss`** via **Inno Setup** (`choco install innosetup`) to **`dist/DSSToolsSetup.exe`**, emits **`checksums.txt`** for that setup program, and publishes **setup + checksums** (not the loose PyInstaller exe) via **`softprops/action-gh-release`** (`contents: write`).
- **`discover_app_version()`** now honours **`DSS_APP_VERSION`** env, then **`dss_app_version.txt`** in **`sys._MEIPASS`** when **`sys.frozen`**, then package metadata / **`pyproject.toml`** � so frozen CI builds report the correct version for update comparison.

### UI theme colours (Configuration)

- Added **`UiThemeColors`** (frozen defaults: soft rose alert rows, teal-leaning crew totals, slate tooltips, pink report outline) persisted under `app_settings.ui_theme` in config JSON.
- **Settings ? Configuration ? Appearance:** hex entry + colour picker per semantic slot, **Reset colours to sample defaults**, validated on Apply; all `DataTable` instances and email preview table refresh tags; Reports outline and Formatting Rules **ToolTip** popups use the saved tooltip colours.
- Helpers: **`normalize_ui_hex_color`**, **`parse_ui_theme_payload`** (partial / invalid keys fall back to defaults).

### Loading hints file (auxiliary; UI not wired yet)

- Added **`LOADING_HINTS.txt`** at the repository root: many single-line hints covering parsing boundaries (`K25:AZ36`), revision behaviour, AZ2 warnings, cache and semantic hash, quick load and cancel hotkey, filters and layouts, Outlook and email drafts, formatting rules, combined-name caveats, maintenance buttons, `%LOCALAPPDATA%\DSSTools`, and similar limitations.
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
- **Tests:** app-settings round-trip and hotkey helper tests; full suite **71** tests passing (includes path-like column helper and cache/hash cases).

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

- Extended tests: revision parsing, preferred-sheet / AZ2 behaviour, `combine_sheet_hashes`, app-settings round-trip (quick-load fields), hotkey helpers, path-like column detection, disk cache `file_hash` contract, `rv` sheet suffix. **71** tests, `python -m unittest test_dss_tools -q`.

**Documentation**

- **`README.md`:** Summaries order, quick load, partial refresh / per-sheet hashing, PF-in-progress text, double-click source, AZ2 warnings, column-drag indicator.
- **`Feature request list.txt`:** Related checklist items marked complete.

### Table column auto-width (same cycle)

- **`DataTable`:** After each filtered render, column widths are set from `tkinter.font` text measurement of the heading and visible cell values (newlines flattened for width).
- **Path-like columns** (`source_file`, configured `source_file_column`, `sources`, any `*_path`) use a lower pixel cap than general text; other columns use a high ceiling so wide text (for example **Details**) still fits without unbounded growth.
- **Saved `column_widths` in `table_layouts`:** Widths from disk are no longer applied on load so content-based sizing stays consistent; widths are still written when layouts save for backward compatibility.
- **Stretch:** Only the last visible column uses `stretch=True` so extra horizontal space is absorbed there.
- **Tests:** `is_path_like_table_column` unit coverage; full suite **71** tests with `python -m unittest test_dss_tools -q`.

### Table UX and cache/revision fixes (same cycle)

- **Treeview columns:** all logical columns use `stretch=False` so interactive width changes do not auto-shrink the next column; horizontal scroll covers overflow.
- **Word wrap:** per-`DataTable` toolbar button `Word wrap (off)` / `(on)`; wraps displayed cell text to the current column pixel width and applies a custom `ttk.Style` row height so wrapped lines show; source-file double-click strips embedded newlines before opening.
- **Disk cache key:** after a **Partial Refresh**, `save_cached_daily_records` now stores the same **workbook content hash** as `load_cached_daily_records` expects (previously the semantic fingerprint was written, causing a perpetual disk miss for that file until re-parse).
- **Sheet revision names:** `REVISION_PATTERN` recognises **`rv`** + digits (e.g. `2026-04-26 rv1`) so `parse_sheet_revision` aligns with AZ2 revision level `1`.
- **Tests:** `parse_sheet_revision` cases for `rv1`/`RV2`; `test_disk_cache_load_requires_workbook_content_hash`; suite **71** tests.

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
  - `C:\Users\LochlanRoss\Documents\GitHub\DSS-Tools` (local clone path; adjust to match your machine)
- The current Codex config is using Windows `sandbox = "unelevated"` with per-project trust entries.
