# ADR-0029 — Workbook schema version 4: hierarchy, derived Duration, fixed row geometry

- **Status:** Accepted
- **Date:** 2026-09-29
- **Relates to:** ADR-0007 D2/D3/D7, ADR-0026, ADR-0027, ADR-0028
- **Implements:** roadmap row R4.7C
- **Decided by:** product owner, 2026-09-29 (*"No need to migrate anything as
  there are no active users, this is still in development."*)

## Context

`GanttSchemaVersion.CurrentSchemaVersion` is **3**. Its own documented rule says to
bump it when *any* of the following changes: a column name or column order, a
catalogue display name, catalogue contract metadata (including **allowed label
positions**), or the exact `tblGanttSettings` key/value contract.

Four decisions in this pre-R4.9 sequence each independently require a bump, and
two of them change the visible column set.

`GanttTableSchema.Default` currently declares 13 required columns plus an optional
`SortOrder`, in this order: `Id`, `LaneId`, `StackIndex`, `Type`, `Description`,
`Start`, `Finish`, `ParentId`, `StyleKey`, `LabelPosition`, `FillColour`,
`StrokeColour`, `Visible`, `SortOrder`.

## Decision

- **D1 — `Duration` is added as a required column by R4.7C**, positioned after
  `Finish` and before `ParentId`. It is **visible and locked**: GanttCreator owns the value and
  recalculates it on Refresh, because the analyst wants to read it but must not
  edit it. REV5 §10 and REV6 §11 both require this, with
  `DurationDays = (Finish.Date - Start.Date).Days + 1` for inclusive spans, and
  an approved non-duration marker (`-`) for milestones and delineators, blank for
  `Splitter`/`Spacer`, and **no misleading value** for invalid or missing dates.
  Dates are date-only; time-of-day is normalised away.
- **D2 — `SiblingOrder` is added as a required column by R4.7A**, positioned
  immediately after `ParentId`. It is engine-maintained and hidden. REV6 §3 requires it so that
  hierarchy order is not defined solely by physical Excel row number, which would
  make a user sort silently redefine the tree.
- **D3 — `GanttRowHeightPt` is added to the metric-token catalogue** (ADR-0026
  D1), and **`CriticalLinePt` is removed from it** (ADR-0027 D4). The
  `CriticalInterval` preset keeps its `ActivityHeightPt` reference and gains a
  `CriticalFill` colour token carrying `#FF0000` (ADR-0027 D2). The
  `CriticalInterval` type gains the `Fill` colour capability (ADR-0027 D5).
- **D4 — `Above` and `Below` are removed from the permitted label-position sets**
  (ADR-0028 D1). Retained: `None`, `Auto`, `Left`, `Right`, `Inside`, the four
  delineator corners, and the splitter positions.
- **D5 — The schema advances in two steps, one per row that changes it, and
  never two rows for one version number.** This ADR originally said "3 → 4" for
  all of D1-D4, which was **wrong**: `R4.7A` adds the `SiblingOrder` column and
  `R4.7C` adds `Duration`, the metric-token changes and the label set. Two rows
  claiming one bump would either collide or skip a version, and
  `ConfigIntegrity` cannot detect that, because the version is the signal it
  compares against. Corrected on 2026-09-29 to:

  | Step | Row | Change | Version |
  | --- | --- | --- | --- |
  | 1 | R4.7A | `SiblingOrder` column | **3 → 4** |
  | 2 | R4.7C | `Duration` column, `GanttRowHeightPt`, `CriticalLinePt` retired, `CriticalFill`, label set | **4 → 5** |

  Within each row, the bump is the **last** change in that row's commit sequence,
  so all the other changes land first and the integrity checker is the gate rather
  than a manual review. An unchanged version would make every change above
  invisible to `ConfigIntegrity` and to the R2.10 repair workflow, which exists
  precisely to notice a contract change.
- **D6 — There is no migration path.** The product owner confirmed there are no
  active users, so `GanttTableSchema.Default` is edited **in place** and no
  legacy-workbook upgrade, conversion, or back-compat test is written. The
  integrity checker still reports a version mismatch rather than silently
  accepting one, because a mismatched workbook is still a fact the user needs
  told.
- **D7 — The engine columns stay physically in `tblGanttData`** rather than
  moving to the VeryHidden configuration sheet. REV5 §11 and REV6 §1 both permit
  this where it avoids disruptive migration, and it keeps one table for one
  logical row. They are initialised **hidden and locked**, and that state is
  repaired if a user unhides them.
- **D8 — Column classification is a first-class contract, not a comment.** Every
  managed column is classified as *user-visible and editable*, *user-visible and
  locked* (`Duration`), or *engine-hidden and locked* (`Id`, `ParentId`,
  `SiblingOrder`, `LaneId`, `StackIndex`, `StyleKey`, and the fill/stroke/label
  overrides). R4.7C owns producing that classification and repairing it.

## Consequences

- `GanttRowDto` gains `DurationCell` and `SiblingOrderCell` in the same
  constructor positions as the schema, so the reader, the validator and the
  schema cannot disagree about column order.
- A user's first save after this change produces a version-4 workbook. Because
  no migration is written, an older workbook is reported, not converted — which
  is the correct behaviour for a pre-release product and must be stated in the
  Ribbon-visible error rather than left to a generic integrity message.
- Retiring `CriticalLinePt` removes a row from the materialised metrics table on
  the VeryHidden sheet. D5's bump is what makes the integrity checker notice that
  the materialised catalogue no longer matches the code.
- ADR-0028 D1's removal of two label positions means a stored `LabelPosition` of
  `Above`/`Below` becomes unresolvable. D6 applies: reported, never coerced.
- Every D1-D4 change is a schema contract change, so **all of them must land
  before the version is bumped**, and the bump is the last of them. A partial
  landing is caught by the integrity checker, which is why the checker is the
  gate rather than a manual review.

## Alternatives considered

- **Move engine columns to `_GanttCreatorConfig`.** Rejected: it splits one
  logical row across two sheets and makes the outline/row alignment harder to
  reason about, for no user-visible gain, since the columns are hidden either
  way. D7 keeps one table.
- **Write a migration for version-3 workbooks anyway.** Rejected on the owner's
  explicit direction, and it would be untestable — no version-3 workbook with real
  user content exists to migrate.
- **Leave the version at 3** because there are no users. Rejected: the version is
  not a user-count concept, it is a contract-drift detector. Leaving it at 3
  would disable the integrity check that exists to catch exactly these changes.
- **Make `Duration` a user-editable Excel formula rather than a written value.**
  Rejected per REV5 §10: the add-in owns the value, one bulk write is cheaper
  than a per-row recalculation, and a formula would be user-visible in the
  formula bar, which contradicts the locked-column presentation.
- **Make `SiblingOrder` optional.** Rejected: it is the engine's authoritative
  ordering, so it must always be present. A blank means the engine has not yet
  assigned one, which validation reports rather than treating as ordering.
