# ADR-0029 — Workbook schema versions 4, 5 and 6: hierarchy, derived Duration, fixed row geometry, horizontal-only labels

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
- **D4 — The label-set change belongs to R4.7G, not to R4.7C.** ADR-0028 D1
  removes `Above` and `Below` from the permitted label-position sets, and the
  retained set is `None`, `Auto`, `Left`, `Right`, `Inside`, the four delineator
  corners, and the splitter positions. This is a schema contract change and so
  needs a bump, but **it is R4.7G's bump**: ADR-0028 and `03-ROADMAP.md` both
  assign the removal to R4.7G, and the R4.7C work item's own scope list omits it.
  Earlier drafts of this ADR listed the label set among R4.7C's changes and put
  it in step 2 of the table below; that is corrected to a step of its own, taken
  by R4.7G at **5 → 6**.
- **D5 — The schema advances in four steps, one per change that requires it, and
  never two changes for one version number.** This ADR originally said "3 → 4" for
  all of D1-D4, which was **wrong**: `R4.7A` adds the `SiblingOrder` column,
  `R4.7C` adds `Duration` and the metric-token changes, and `R4.7G` changes the
  permitted label-position set. Three rows each claiming one bump would either
  collide or skip a version, and
  `ConfigIntegrity` cannot detect that, because the version is the signal it
  compares against. Corrected on 2026-09-29 to, and extended on 2026-09-30 when
  the label-set change was moved off R4.7C onto R4.7G (D4):

  | Step | Row | Change | Version |
  | --- | --- | --- | --- |
  | 1 | R4.7A | `SiblingOrder` column | **3 → 4** |
  | 2 | R4.7C | `Duration` column, `GanttRowHeightPt`, `CriticalLinePt` retired, `CriticalFill` | **4 → 5** |
  | 3 | R4.7G | permitted label-position set (`Above`/`Below` removed, ADR-0028 D1) | **5 → 6** |
  | 4 | R4 QA review | `SizePreset` and `RangePaddingDays` settings keys added (owner ruling 2026-10-01) | **6 → 7** |
  | 5 | R4.7I | reserved row above the table carrying the table title and year band; the header row moves from worksheet row 1 to row 2 and the plot anchor moves with it (ADR-0030 D4) | **7 → 8** |
  | 6 | R4.7I | chart padding rows above the title row and below the last activity row; the header moves from row 2 to row 3 (ADR-0031 D1) | **8 → 9** |
  | 7 | R4.7I | the bottom margin becomes a **reserved** worksheet row resolved from the body's measured span rather than derived arithmetic, and `MaximumExternalLabelWidthPt` is retired from the metric catalogue (ADR-0035 D1/D2) | **9 → 10** |

  **Steps 5, 6 and 7 are the R4.7I layout changes, and version 7 does NOT cover
  them.** Version 7 is the two added settings keys and nothing else. The layout is
  version **8**, added by ADR-0030's reserved row, version **9** by ADR-0031's
  padding rows, and version **10** by ADR-0035's reserved bottom padding row.
  This table previously stopped at step 4 while
  `GanttSchemaVersion.CurrentSchemaVersion` was 9, which left a reader unable to tell
  which change a given workbook carried — and `ConfigIntegrity` distinguishes
  pre-layout from post-layout workbooks **by this number and nothing else**. A
  version-7 workbook has its header on row 1, a version-8 one on row 2, and a
  version-9 one on row 3; because the anchor repair writes the *expected* address, a
  workbook carrying the wrong version would otherwise pass its own integrity check
  while its plot anchor pointed at the wrong row. Recording the steps here is what
  keeps that signal readable, and D6's "reported, never coerced" is what keeps it
  safe.

  **Step 7 is a schema change on both of its own axes**, which is why it takes its
  own version rather than riding on step 6. The reservation changes what a bottom
  margin *is* — a version-9 workbook may carry user content in the row version 10
  treats as the margin, so the row-height normaliser and the panel measurement now
  refuse rather than write to it — and the metric catalogue loses
  `MaximumExternalLabelWidthPt`, so a version-9 workbook's `tblGanttMetrics` carries
  a row the running add-in no longer knows and its stored catalogue hash no longer
  matches. Both are contract changes under the same bump rule D5 applies to steps
  1-3, and neither is migration (D6): the remedy is Initialise, which is also what
  performs the reservation. **Step 7 was added on 2026-10-02** because the table
  stopped at 6 while `CurrentSchemaVersion` was already 10, which is the same
  unreadable-version defect the step-4 gap created.

  Step 4 was added by the senior QA review of R4 and is **not** a roadmap row. Both
  keys were read by the scene-request factory from the first version of R4.8A but
  were never part of the approved key set, so `ValidateSettings` could not return
  them: every live chart was permanently A4-portrait with a 7-day plot-range pad,
  while `A_known_preset_key_is_honoured` stayed green because it injected the key into
  a hand-built dictionary. Adding a key is a change to the exact key set the bump rule
  names, so it takes its own version rather than riding on 6 — the same reasoning D5
  applies to steps 1-3. There is no migration: a version-6 workbook reports a version
  mismatch and is reported, never coerced (D6).

  **No row is excluded from the progression**; each schema change lands in the
  row whose change requires it, and each row bumps exactly once.

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
- A user's first save after R4.7C produces a version-**5** workbook, because
  R4.7C's bump is the second of the three steps D5 covers. Because
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
