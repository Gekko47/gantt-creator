# ADR-0042 — Destructive confirmation opt-out checkbox

- **Status**: Accepted
- **Date**: 2026-10-13
- **Context**: ADR-0008 D3 makes the warned confirmation the only
  pre-execution guard for destructive commands, which per D2 have
  no undo and no rollback. The delete-row command (`RemoveRow`)
  therefore confirms on every invocation. Operators deleting many
  rows in one Excel session found the per-delete confirmation
  repetitive and asked whether it could be shown once per session.
  Silently suppressing it would remove the only guard for an
  irreversible operation, so any opt-out must be explicit and
  user-controlled, not a silent default.
- **Decision** (human-approved 2026-10-13, amending ADR-0008 D3):
  - **D1 — Default guard unchanged.** Every destructive command
    still shows its confirmation dialog by default. ADR-0008
    D2/D3 are unchanged: the dialog is the only pre-execution
    guard, and its body still states the change cannot be undone
    (the exact phrasing remains pinned by test).
  - **D2 — Explicit opt-out checkbox.** The delete-row confirmation
    dialog offers a verification checkbox labelled "Don't ask
    again this session". The opt-out is recorded only when the
    user both confirms the change (Yes) and checks the box. A
    declined dialog, or a confirmation without the box checked,
    leaves the session un-opted-out.
  - **D3 — Session scope.** "Session" is the add-in's loaded
    lifetime (one Excel process). The opt-out state is
    process-scoped and held in memory only; it is not persisted to
    the workbook or to the `_GanttCreatorConfig` worksheet (which
    stores catalogues, styles, metrics and configuration metadata,
    never session state), and it resets when the add-in reloads. It
    is not per-workbook: a second workbook opened in the same
    Excel process inherits the opt-out.
  - **D4 — Effect.** Once opted out, subsequent delete-row commands
    in the same session proceed without the confirmation dialog,
    straight to the mutator. The user has explicitly chosen to drop
    the ADR-0008 D3 guard for the rest of the session; this is the
    informed trade-off recorded here.
  - **D5 — Fallback.** When the managed TaskDialog cannot be shown
    and the code falls back to a plain MessageBox, no verification
    checkbox exists, so no opt-out is recorded on that path; the
    confirmation still gates the delete.
  - **D6 — Scope.** The opt-out is adopted by the delete-row
    confirmation only. The configuration-repair confirmation
    (`RepairConfigCommand`) and any other destructive command keep
    the unconditional per-invocation confirmation unless a future
    ADR extends this pattern.
- **Consequences**: The delete-row confirmation is the only
  destructive command whose guard is user-dismissable. The
  confirmation body text is unchanged. The opt-out short-circuit
  (`ConfirmationSession.DontAskAgain`) and the checkbox-state
  recording (`ConfirmationSession.RecordChoice`) are pure and
  unit-tested without a dialog host; the live checkbox read after
  `TaskDialog.ShowDialog` follows the documented TaskDialog pattern
  and is recorded as F5 evidence.
- **Alternatives considered**:
  - Silently show the confirmation only on the first deletion per
    session (rejected — removes the guard without the user's
    explicit consent, contradicting ADR-0008 D3).
  - Per-workbook opt-out persisted in the configuration worksheet
    (rejected — the invariant forbids session state in the
    configuration worksheet, and a persisted opt-out would outlive
    the session).
  - An application-level undo (rejected — ADR-0008 D2 records why
    undo is out of scope).
