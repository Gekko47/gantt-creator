# ADR-0057 — Repair rebuilds a missing data table

- **Status:** Accepted (owner ruling 2026-11-11)
- **Date:** 2026-11-11
- **Amends:** ADR-0056's three-state ribbon model, by giving its `Needs repair`
  state a working recovery action
- **Related:** ADR-0011 (safe configuration repair policy), ADR-0055 (add-in-owned
  sheet protection), R2.10 (configuration repair), R5.16 (ribbon states)

## Context

R5.16 introduced three ribbon modes. `Needs repair` means *the workbook has
`_GanttCreatorConfig` but no readable `tblGanttData`*; in that mode the ribbon
shows the full surface with the table-dependent controls disabled and **Fix** as
the one enabled action. That is the whole point of the state: it exists so a
user who damages the table is not left with nothing to press.

It does not work. Measured on 2026-11-11: the owner deleted `tblGanttData` and
pressed **Fix**, and nothing happened at all.

The cause is mechanical and was read from the code, not inferred:

- [`RepairConfigCommand.Run`](../../src/GanttCreator.AddIn/RepairConfigCommand.cs)
  builds a plan from the configuration integrity check and then
  `if (plan.Findings.Count == 0) { return; }` — a **silent** return.
- Deleting the data table leaves the configuration sheet intact, so the checker
  finds nothing wrong with the configuration, the plan is empty, and the command
  exits without a word.
- Every other path is closed as well. `Initialise` refuses
  `ConfigSheetExists` while `_GanttCreatorConfig` exists, every table command
  refuses `TableMissing` with "Run Initialise sheet first", and the
  configuration sheet cannot even be deleted by hand to make Initialise legal:
  it is `xlSheetVeryHidden`, so it is absent from the sheet-tab list, and a
  programmatic `Worksheet.Delete()` on it is refused by the host with
  **`0x800A03EC`** (measured, 2026-11-10).

So `Needs repair` currently advertises a recovery that does not exist, and the
only enabled action is a silent no-op. That is a dead end in a state the
add-in itself creates.

## Decision

**Fix is the recovery action for `Needs repair`, and a repair that finds nothing
must never be silent.**

1. **Rebuild.** When the table read refuses with `TableMissing` *and* reports
   `ConfigurationPresent`, **Fix** rebuilds `tblGanttData` on the existing Gantt
   worksheet before running the ordinary configuration repair.
2. **On the existing sheet, never a second one.** The rebuild targets the single
   worksheet carrying `GanttWorkbookContract.GanttSheetLabel` (case-insensitive)
   that holds no table. Any other shape — no such sheet, more than one, a sheet
   with a table — is reported and nothing is changed. Creating another visible
   worksheet would break the one-visible-worksheet invariant that the whole
   workbook contract rests on.
3. **The user's rows are adopted, not discarded.** The rebuild writes the schema
   header row at the sheet's used-range origin, creates the table over the header
   and whatever body rows exist, and names it `tblGanttData`. Rows that were on
   the sheet before the table was lost become body rows.
4. **Then the ordinary repair.** After the table exists, the Type-column list
   rule, the plot anchor, and the configuration catalogues are re-materialised
   through the existing repair path, so the rebuilt workbook is the same state
   Initialise would have produced.
5. **Never silent.** Every Fix outcome presents something: what was rebuilt, what
   was repaired, what was skipped, or the typed reason nothing could be done.
   The `Findings.Count == 0` early return becomes a reported "nothing to repair".
6. **Confirm when rows would be adopted.** A rebuild over a sheet that already
   holds body rows is a confirm-required change; a rebuild over a header-only or
   empty sheet is automatic. This follows ADR-0011's classification: changing
   user data asks first.

## What this trades

- **Fix now mutates a table, not only configuration.** The command's blast radius
  grows from the hidden configuration sheet to the visible one. The mitigations
  are the narrow trigger (only `TableMissing` + `ConfigurationPresent`, only the
  one standard-labelled sheet, never a second sheet), and the confirmation when
  user rows would be adopted.
- **A second opinion about the same workbook state.** Two commands can now
  produce a table: `Initialise` (first run, no configuration) and `Fix`
  (configuration present, table lost). They must not diverge, so the rebuild
  must reuse the same schema, neutral appearance, Type rule, and plot anchor
  rather than defining its own.
- **The invariant count stays one.** Because the rebuild never creates a
  worksheet, "exactly one visible user worksheet plus one VeryHidden
  configuration worksheet" remains true in every reachable state.

## Consequences

- `Needs repair` finally has the recovery its name promises, and the three-state
  model becomes honest: first run → `Initialise`, damaged table → `Fix`, healthy
  → the full surface.
- `RepairConfigCommand` grows a pre-step and its tests grow accordingly: a
  present-config/missing-table fixture, an adoption confirmation, and the
  "nothing to repair" message.
- The rebuild needs a typed outcome of its own (rebuilt / refused because the
  sheet is missing, ambiguous, or protected), because a silent refusal here is
  the defect this ADR exists to remove.
- The live gate must prove the state end to end: delete the table, press Fix,
  see the table return with the sheet's rows intact and the Type rule applied.
