# ADR-0008 — Destructive-command policy: no undo, warned confirmation

- **Status**: Accepted
- **Date**: 2026-09-22
- **Context**: R2.7a introduces COM-mutating commands (clearing the visible
  `tblGanttData` body, recolumnising it, and resetting the
  `_GanttCreatorConfig` catalogues back to code-owned defaults). These
  operations overwrite live workbook state, so the add-in needs a written
  policy for whether it offers an undo, what the confirmation says, and
  what the live gate for "command ran and corrupted something" is. Three
  open questions (manifest decision D-G2) required human approval:
  whether mutating commands expose a reversible undo, what confirmation
  text they present, and which observable proves a destructive command
  actually ran.
- **Decision** (human-approved 2026-09-22):
  - **D1 — Scope.** "Destructive command" means any Office-mutating adapter
    path that clears, overwrites, or recreates user-visible cell content or
    table geometry in `tblGanttData` or the configuration worksheets. Normal
    worksheet edits, Type/colour/label-position/selection changes, and
    Refresh/render are not destructive commands under this ADR; they do not
    render and do not need this policy.
  - **D2 — No undo.** Destructive commands run without an Excel undo-stack
    entry and without a transactional rollback. The add-in does not build an
    application-level undo for them and does not attempt to snapshot-and-restore
    workbook state. Rationale and trade-offs are recorded in the running
    KNOWN-LIMITATIONS ledger under "destructive command undo"; they are not
    repeated here.
  - **D3 — Warned confirmation.** Each destructive command exposes a
    confirmation dialog whose body explicitly states that the change cannot
    be undone, using exactly that phrasing (the Core text builder pins the
    exact string and the confirmation path has a positive test). The dialog is
    the only pre-execution guard; there is no silent truncation, partial
    execution, or "best-effort" fallback.
  - **D4 — Guard ordering.** The workbook-protection guard is the first check
    in every mutating adapter (protected worksheet → refuse before any COM
    mutation), followed by the confirmation dialog when the operation is
    user-initiated. The confirmation is not a substitute for the protection
    guard, and the protection guard is not a substitute for the confirmation.
  - **D5 — Live gate.** The acceptance gate for R2.7a is the protected-sheet
    refusal: Initialise (or any future destructive command) on a protected
    worksheet must return a typed refusal outcome and perform zero partial
    mutation before it returns. The live exercise where a destructive command
    is confirmed and actually mutates a real workbook is recorded separately
    at R5.8 and cross-referenced here; it is not the primary acceptance gate,
    because that gate must prove the refusal path, not the happy path.
  - **D6 — Cross-reference to manual destructive run.** The confirmed-destructive
    Office run (a destructive command confirmed and executed against a live
    workbook with the result verified) lives in R5.8's evidence and is
    referenced from this ADR; it is not duplicated as an automated test, because
    it requires a live COM host. Any future automation of that path is a
    separate decision.
- **Consequences**: Every future Office mutating adapter must start with the
  protection guard (D4) and, when user-initiated, the warned confirmation (D3).
  The Core text for the confirmation is pinned by exact-string tests so the
  "cannot be undone" phrasing cannot drift. The protected-sheet refusal is a
  first-class automated contract test with a zero-mutation assertion. The
  manual destructive run in R5.8 remains the only live observation of a
  confirmed destructive command actually mutating a workbook.
- **Alternatives considered**:
  - Application-level undo via an Excel undo-stack entry or an add-in snapshot
    (rejected — no approved implementation path, and the running ledger records
    why undo is out of scope; not pursued here).
  - Making the confirmation dialog the only guard and dropping the protection
    guard (rejected — a protected sheet must refuse before any mutation, not
    after a dialog).
  - Using the happy-path "command actually ran" as the acceptance gate instead
    of the protected-sheet refusal (rejected — the gate must prove the refusal
    path; the destructive run is recorded separately).
