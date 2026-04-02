# UI Redesign Regression Checklist

Use this checklist before redesign work begins and after each redesign phase.

Date: 2026-04-02
Tester: Manual QA (User)
Phase: Phase 4
Result: PASS

## Startup and Settings

- [x] App launches without errors.
- [x] Saved settings load on startup.
- [x] App window opens centered and usable at minimum size.
- [x] Preference toggles persist after restart.

## Pending Files Workflow

- [x] Pending folder browse works.
- [x] Pending list loads current PDFs.
- [x] Auto-refresh updates list when files are added/removed externally.
- [x] Selection supports click, Ctrl-select, Shift-range, and select-all toggle.
- [x] Pending count label remains accurate.

## Pending Actions

- [x] Preview opens selected PDF.
- [x] Rotate window opens selected files and saves correctly.
- [x] Rotated output preserves file integrity and expected page orientation.

## New Record

- [x] New Record window opens from selected pending file(s).
- [x] Name and year validation works.
- [x] Record file is created with expected naming rules.
- [x] Source pending file is archived to processed after success.

## Merge Existing

- [x] Merge Existing window opens from selected pending file(s).
- [x] Employee folder selection/autocomplete works.
- [x] Merge output is correct and save flow succeeds.
- [x] Source pending file is archived to processed after success.

## Employee Sources and Parsed Names

- [x] Add/Remove/Clear source files works.
- [x] Parsing progress status updates correctly.
- [x] Parsed names can be viewed and exported.
- [x] Strict/Lenient name filter behavior remains correct.

## Employee Details Editor

- [x] Employee Details window opens.
- [x] File list loads for selected folder.
- [x] External add/remove/rename is auto-detected.
- [x] Bulk date updates still work.
- [x] Open Folder and related actions work.

## About and Updates

- [x] About dialog opens.
- [x] Version is correct.
- [x] Build commit/date metadata resolves correctly.
- [x] Update status message behavior remains correct.

## Final Quality Gate

- [x] No new diagnostics errors in edited files.
- [x] No workflow regressions observed in smoke run.
- [x] UI remains responsive during common operations.

## Notes

- Issues found: None.
- Reproduction steps: N/A.
- Screenshots taken: N/A.
