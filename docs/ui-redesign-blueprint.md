# PDF Record Manager UI Redesign Blueprint

## 1) Goal

Modernize the full UI to look professional, organized, and enterprise-grade while preserving all existing workflows and behavior.

## 2) Non-Negotiables

- No feature regressions in New Record, Merge Existing, Rotate, Preview, and Employee Details flows.
- Keep all existing data and file-processing rules unchanged.
- Keep settings persistence behavior unchanged.
- Keep update/about/release metadata behavior unchanged.
- Keep keyboard and mousewheel behavior unchanged unless intentionally improved and verified.

## 3) Current Architecture Reality

- UI composition and logic are heavily centralized in main.py.
- Styling is already partially centralized (theme and ttk style setup), which is a good base for redesign.
- Main dashboard layout is currently assembled in one section, so layout modernization can be done with controlled refactors.

## 4) Redesign Strategy (Phased, Low Risk)

### Phase 0: Safety Net (No Visual Changes)

Deliverables:

- Regression checklist for critical flows.
- Simple smoke-test routine to run after each UI phase.
- Snapshot list of existing UI states (normal, empty, long lists, error dialogs).

Exit criteria:

- We can repeatedly verify behavior after every design change.

### Phase 1: Design System Foundation

Deliverables:

- Semantic color tokens (surface, panel, border, text-primary, text-muted, success, warning, danger, focus).
- Consistent spacing scale (4, 8, 12, 16, 24, 32).
- Typography scale (app title, section title, label, metadata, status).
- Button hierarchy (primary, secondary, subtle, destructive, toolbar icon button).

Rules:

- Do not change command callbacks.
- Do not change business logic.
- Only map existing widgets to semantic style names.

Exit criteria:

- Whole app theme can be tuned by editing centralized tokens only.

### Phase 2: Dashboard Layout Recomposition

Deliverables:

- Structured shell:
  - Top: app identity + status strip.
  - Left: folders and employee source management.
  - Right: pending files operations and selection actions.
  - Bottom: primary workflow action bar.
- Better visual grouping with consistent card spacing and section headers.
- Cleaner alignment, spacing, and interaction density.

Rules:

- Keep existing state variables and handlers.
- Rewire UI containers only.

Exit criteria:

- Dashboard looks organized and modern without behavior changes.

### Phase 3: Dialog and Window Consistency

Deliverables:

- Standardized dialog chrome for New Record, Merge Existing, Rotate Preview, and Employee Details.
- Consistent button order and footer actions.
- Shared helper builders for repeated row patterns and field/action blocks.

Rules:

- Preserve all current function signatures and event behavior.

Exit criteria:

- All major windows feel like one cohesive product.

### Phase 4: Enterprise UX Polish

Deliverables:

- Empty states and loading states with clearer guidance.
- Better inline validation messaging and conflict warnings.
- Accessibility improvements: focus ring consistency, contrast checks, keyboard traversal checks.
- Optional compact mode switch for dense operations.

Exit criteria:

- UI quality is release-ready for enterprise users.

## 5) Regression Checklist (Run After Every Phase)

- Launch app and load settings successfully.
- Select pending/root folders and refresh pending list behavior.
- Multi-select pending files (single, ctrl, shift, select-all toggle).
- Preview selected PDF.
- Rotate selected pending PDFs and confirm saved output.
- New Record flow creates destination and archives pending file.
- Merge Existing flow merges correctly and archives pending file.
- Employee Details loads folder and reflects external file add/remove/rename.
- Parsed names window opens and export works in supported formats.
- About dialog shows version/build metadata and update feed state.
- App restart and exit paths work.

## 6) Implementation Order in Code

1. Refactor style tokens and ttk style names.
2. Recompose dashboard container layout only.
3. Standardize toolbar/button styles and section headers.
4. Normalize key dialogs one by one.
5. Run regression checklist after each step.

## 7) Risk Controls

- Keep each commit small and focused by phase.
- Avoid mixing style changes with processing logic changes.
- Keep a rollback-safe history (one checkpoint commit per phase).
- If behavior breaks, revert only the latest phase commit.

## 8) Definition of Done

- UI is visibly modern and consistent across all major screens.
- Critical workflows pass the regression checklist.
- No new diagnostics/runtime errors introduced.
- Theme and layout are maintainable through centralized style tokens and small UI builder helpers.

## 9) Progress Checkpoint

Completed:

- Phase 1 design tokens and semantic style hierarchy.
- Phase 2 dashboard shell re-layout.
- Phase 3 dialog consistency pass completed for New Record, Merge Existing, Rotate Preview, and Employee Details.
- Bulk PDF Date Update popup aligned to the same header/card/action-bar pattern.
- Phase 4 polish pass applied for inline year guidance, clearer empty-state messaging, and keyboard focus/escape handling in major dialogs.

Remaining:

- None. Phase 4 regression checklist completed with no errors reported.

## 10) Immediate Next Step

Create a Phase 4 completion checkpoint commit, then optionally prepare a release tag and release notes.
