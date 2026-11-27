# ADR-0073 — Additive settings schema repair: upgrade a one-step, settings-only bump in place

- **Status:** Accepted
- **Date:** 2026-11-26
- **Relates to:** ADR-0007 D2/D3/D4, ADR-0011, ADR-0029 D6
- **Implements:** R8.9
- **Decided by:** product owner, 2026-11-26

## Context

`CurrentSchemaVersion` moves from 28 to 29 because `tblGanttSettings` gains the
custom-size keys `SizeWidthCm` and `SizeHeightCm`. The existing policy is
ADR-0029 D6: no migration, a mismatch is reported, and the remedy is
Initialise.

Two facts make that policy unusable for this bump:

1. **Initialise is not reachable in the current UI flow.** The Get-started
   group is not available to an already-initialised workbook, so telling the
   user to run it strands the file.
2. **The existing repair path destroys user data.** A schema bump surfaces as
   `CatalogueHashMismatch` or `RowCountMismatch`; the checker classifies it as
   an automatic catalogue repair, and `ExcelConfigRepairer` calls the full
   `ExcelConfigCatalogueWriter.Write()`. That path writes `BuildMetricRows()`
   over `tblGanttMetrics`, discarding every theme-edited metric value.

The bump itself is additive: two settings rows are appended at the end of
`tblGanttSettings`, and the version and catalogue hash in `tblGanttConfig`
move. No column, type, style, or metric changes.

## Decision

- **D1 — A one-step, settings-only schema bump is repaired in place.** The
  repair appends the missing settings rows with their catalogue defaults,
  preserves every present value, and rewrites the version and hash. The user
  does not run Initialise.
- **D2 — Detection is strict and pure.** `AdditiveSchemaUpgrade.TryPlan`
  succeeds only when the stored version is exactly one behind the current
  version, the stored setting keys are an exact prefix of
  `GanttCatalogues.Settings`, the stored set is non-empty, and every missing
  key is an appended settings row. A reorder, an extra key, a missing middle
  key, an empty set, a newer workbook, or a multi-version gap is refused.
- **D3 — The repair is narrow.** It writes only `tblGanttSettings` and
  `tblGanttConfig`, under one application-state scope and one protection
  guard, with the config version and hash written last and together. It never
  calls the full catalogue writer.
- **D4 — The finding is distinct.** `SchemaVersionOlderAdditive` is
  `AutoRepair`. Structural `SchemaVersionOlder` remains `ConfirmRequired`, so
  a mismatch that is not provably additive still asks before it rewrites.
- **D5 — The message names the reachable action.** `SchemaVersionMismatchMessage`
  names **Repair**, never Initialise.
- **D6 — The ceiling is one version step.** A multi-version gap falls back to
  the structural repair, because a settings-prefix match cannot prove that
  `tblGanttTypes` or `tblGanttMetrics` did not also change in between. The
  upgrade path is a version-range migration table when a multi-step repair is
  needed.
- **D7 — User data is the invariant.** `tblGanttMetrics`, `tblGanttStyles`,
  present settings values and the workbook ID survive the repair. Tests pin
  each of them, and a second run writes nothing.

## Alternatives considered

- **Keep Initialise as the remedy.** Rejected: the control is not reachable in
  the current flow, and the message would send the user to a dead end.
- **Keep the full catalogue rewrite.** Rejected: it resets metric edits, which
  is data loss on a change that does not touch metrics.
- **Encode the two values inside the existing `SizePreset` scalar.** This
  avoids the bump, but it hides structured data in a scalar key and does not
  solve the next settings-only bump. Rejected in favour of a reusable additive
  repair.
- **Auto-repair every older schema.** Rejected: an older schema may carry
  column or catalogue changes the settings-prefix test cannot see.

## Consequences

- A version-28 workbook opens and repairs itself; no user action is required.
- A workbook more than one version behind still uses the confirmation-required
  structural repair, which is the only path that knows how to regenerate
  types and metrics.
- The `SchemaVersionOlder` finding now has a production producer as well as a
  test producer, so the enum is no longer dead.
- `SchemaVersionMismatchMessage` and its test change from Initialise to Repair.
- The plan's ADR-0029 D6 policy is amended for additive, one-step settings
  bumps only; every other schema change keeps the existing rule.

## Amendment — 2026-11-27: prove the other tables before claiming additive

A live upgrade needed two Fixes. The user saw the `RowCountMismatch` sentence, pressed
Fix, was told "1 repaired", and saw the same refusal on the next Refresh; only a second
Fix cleared it.

**Cause.** The additive plan was chosen from the schema version alone. Any
`RowCountMismatch` became `SchemaVersionOlderAdditive` whenever `ObserveSchema` found a
version exactly one step behind, and that observation proved only the types and metrics
row counts. The reader's other `RowCountMismatch` arms — a styles table short of a
built-in preset, or a ragged style, settings or config row — were invisible to it. The
narrow write then rewrote `tblGanttSettings` and `tblGanttConfig`, which cannot reach a
styles table, so the fault survived and the next Refresh refused verbatim. The second Fix
found the version already current, fell through to the structural rewrite, and worked.

**Rule now enforced.** `ObserveSchema` returns an observation only when the settings row
set, the stored version or the stored hash is the *only* thing that can be wrong:
types and metrics counts match, the styles table carries at least every built-in preset,
and no style, settings or config row is ragged. Anything else returns `null` and leaves
Fix on the confirm-required structural path, which regenerates the table that actually
refused.

**Evidence gap closed.** The reader also records one technical record per refusal naming
the table and its actual and expected body rows, because the user-facing sentence names
only the reason and the log could not attribute the loop at all.

**Proof.** `ConfigCatalogueReaderTests` pins the additive case (one narrow repair leaves
the workbook readable), the regression (a short styles table is refused as non-additive),
the one-Fix guarantee (one structural rewrite repairs it) and the refusal record. The
regression was proved non-vacuous by disabling the styles count guard, which restored the
misclassification `StoredVersion = 28` for a short styles table.

## Evidence

- Full writer metric reset: `ExcelConfigCatalogueWriter.WriteContractTables`
  writes `BuildMetricRows()` (`src/GanttCreator.Office/ExcelConfigCatalogueWriter.cs:483`),
  while the theme path writes `BuildThemedMetricRows(request)` (line 291).
- Settings append and preservation: `BuildSettingRows` (line 1346).
- Config version and hash: `BuildConfigRows` (line 1377).
- Current mismatch message: `SchemaVersionMismatchMessage.Build`
  (`src/GanttCreator.Core/ConfigIntegrity/ConfigIntegrity.cs:119`).
- Current classification: `ConfigIntegrity.Classify`
  (`src/GanttCreator.Core/ConfigIntegrity/ConfigIntegrity.cs:195`).
