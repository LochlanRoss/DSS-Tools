# UI Standardization Notes

These guidelines are intentionally general so they can be reused for future desktop data-viewer or workflow tools.

## Layout and Navigation
- Prefer tabbed or grouped-tab navigation when the app has several distinct workflows.
- Keep operational pages separate from configuration and diagnostics pages.
- Reserve a dedicated Settings or Configuration area for user-tunable behavior.
- Keep the primary workflow visible near the top of the app with obvious open, refresh, export, and cancel controls.

## Table and Viewer Behavior
- Tables should support sorting from headers.
- Tables should support per-column filtering without removing the faster one-click sort behavior.
- Tables should allow column show or hide controls.
- Tables should allow column width changes and column reordering.
- Viewer defaults should prefer the most recent operational data first when recency matters.
- Scrolling should feel responsive for both mouse wheel and horizontal navigation.

## Persistence
- Persist user interface preferences that change how people work day to day.
- Persist at minimum: visible columns, column widths, column order, and sort state.
- Persist user settings, saved templates, group definitions, and similar operator-entered metadata.
- Restore persisted settings automatically on launch.

## Background Work
- Long-running file loads, hashing, network checks, and external-app lookups should run off the UI thread.
- Provide visible progress and a cancel button for long operations.
- Keep the last good data on screen if a refresh fails.
- Prefer partial or progressive loading when a useful early subset can be shown safely.

## Diagnostics and Trust
- Include a diagnostics area for health checks, status summaries, and bug-report helpers.
- Make cache use visible when it affects performance or trust.
- Prefer explicit health or warning pages for parse issues instead of silent failures.
- Preserve enough audit information that users can understand where a reported value came from.

## Settings and Safety
- Put feature toggles, polling frequency, notification controls, and storage-path settings in one place.
- Add reset and cleanup tools for cached data and saved settings.
- Use confirmation prompts for destructive maintenance actions.
- Keep integrations with external systems optional and failure-tolerant.

## Notifications and Resolution Workflows
- Warnings should explain what happened, where it happened, and what the user can do next.
- If a warning can be disabled, say so in the warning dialog.
- Where practical, provide both automatic detection and a manual re-check button.

## Export and Interop
- Allow export of the current view with the active filters applied.
- Keep exported data standalone whenever possible instead of depending on live source files.
- When integrating with email or office tools, prefer draft creation over automatic sending.
