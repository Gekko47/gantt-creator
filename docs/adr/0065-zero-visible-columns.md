# ADR-0065 — Zero visible table columns are a supported state, and the Columns control toggles all of them

- **Status:** Accepted (owner-approved R5.22, 2026-11-17)
- **Date:** 2026-11-17
- **Amends:** the `AGENTS.md` invariant "the worksheet data table is visible" and
  [`docs/02-ARCHITECTURE.md`](../02-ARCHITECTURE.md)'s "core data columns remain visible";
  ADR-0029 D8's classification is reaffirmed, not changed
- **Related:** ADR-0029 (column classification and schema versioning), ADR-0048 (the live frame
  encloses the table), ADR-0050 (empty live chart and coalesced refresh), ADR-0059 and ADR-0064
  (visible column order), ADR-0061 (initialise presents the chart surface), R5.18h (persisted
  column visibility), R5.22

## Context

R5.18h made the six user-visible columns' hidden flags a persisted per-column choice and
refused the toggle that would hide the last authoring column, on the reasoning that a table with
no authoring column visible could not be typed into by hand. That was a usability assumption,
not a data-integrity rule: a hidden column is still present in `tblGanttData`, the reader and
writer still address it by name, and the ribbon can restore it.

The owner now requires the analyst to be able to hide the whole table and work with the chart
alone, and a one-click way to hide or show all six columns. The engine columns remain hidden and
locked under the schema's own classification; the change is limited to the six user columns.

The host behaviour was measured before the decision, not assumed. On Excel 16.0 a synthetic
schema-26 workbook, opened in a private instance:

- accepted `EntireColumn.Hidden = true` for all six user columns;
- reported every hidden width as `0` and the header/body `Left` as the table origin;
- moved the plot anchor column's `Left` onto that same origin while leaving the anchor column
  itself visible;
- left `DataBodyRange` present with its rows; and
- left the active cell on the now-hidden `B2` rather than moving it to a visible column.

The probe is committed as [`scripts/probe-hide-all-columns.ps1`](../../scripts/probe-hide-all-columns.ps1)
and its run is recorded in the R5.22 work item. Freeze panes were not exercised by the probe and
remain an Office-gate observation, not a supported contract.

## Decision

### D1 — Zero visible user columns is a legal persisted state

`HiddenColumns` may name all six user columns. The per-column toggle no longer refuses the last
visible authoring column, and the ribbon no longer carries a "must stay visible" warning.
`ColumnVisibilityRefusal.LastVisibleColumn` and `AnyAuthoringColumnVisible` are deleted rather
than left dead.

### D2 — Hiding is presentation only, and the engine columns stay schema-owned

The six user columns' hidden flags are the user's choice. The engine columns' hidden and locked
state remains the schema's (ADR-0029 D8), and every column of `tblGanttData` stays on the one
visible worksheet. Nothing is moved to `_GanttCreatorConfig`, and no schedule value changes when
a column is hidden.

### D3 — The Columns control is a split button whose primary toggles all six

The primary's action is: none hidden → hide all six; any hidden (partial or all) → show all six.
The primary's label states the action (`Hide all columns` / `Show all columns`), never the
state, because a split button has no pressed state. The menu keeps the six per-column toggles,
each of which states its current state (`Type (shown)` / `Type (hidden)`) and carries the
tick/dash image. The split button is gated on workbook + initialised + Gantt sheet only, never
on selection, because it is the recovery path when the table is invisible.

### D4 — The render path accepts a zero-column measurement, and the live title is disabled

The live measurement returns an empty-column grid rather than refusing, and `PanelCellGrid`
accepts zero columns with a zero summed width. On the live sheet the grid's right edge is the
measured plot-anchor column's left edge, so with every table column hidden the anchor and the
origin coincide and the plot expands into the vacated table width.

The live title is **disabled** while no column is visible (owner amendment 2026-11-17). Its
textbox spans the table width, so at zero columns it either collapses or — widened over the
plot — overlaps the chart, which the owner observed. Omitting the title request makes the
reconciler delete the owned `ChartTitle` shape, and the next refresh that finds any visible
column re-creates it. The title therefore returns by itself when a column is shown; no separate
command or setting is involved. The export title is unaffected, because an export composition
with no panel has nothing to overlap.

`PanelCellGridRefusal.NonPositiveWidth` is retained. A hidden column passed to the grid is still
a caller error; the caller filters hidden columns from the measured list and never hands the grid
a zero width.

### D5 — Persistence, repair, reopen and Initialise agree with the sheet

The toggle applies the worksheet first, persists `HiddenColumns` second, and verifies the
result. A refused apply or persist is compensated back to the previous set and reported; no
refresh is scheduled over a half-changed sheet. A repair/rebuild re-applies the stored set after
the table is recreated. A successful Initialise presents a fresh visible table and clears
`HiddenColumns` so the setting and the sheet cannot disagree. The width readout uses the stored
setting as its authority and is recomputed with the candidate set; the Gantt tab is restored
after the settings write when the automatic refresh is off.

### D6 — Export follows the visible set, and a fully hidden table exports the chart alone

Every export composition builds its data panel from the same visible-column authority as the
live measurement. A hidden column is not exported; when no column is visible, no panel is
emitted and the export is the chart plot alone, over and above `ExportIncludeDataPanel`. This is
a binding forward contract because the export composition pipeline does not exist in `src/` yet;
the Core `VisibleColumns` helper and its facts are the seam the future builder must consume, so
the rule cannot be invented twice.

### D7 — The existing selection semantics are recorded, not changed

Excel keeps the active cell on the hidden cell when its column is hidden; the probe observed
`B2` before and after. The refresh's application-state scope therefore restores a hidden-cell
selection, which the host accepts. This work item records that behaviour and proves it in the
Office gate; it does not introduce a new selection rule.

## Consequences

- **A fully hidden table is not a broken workbook.** The table still validates, reads, writes,
  repairs, reopens and exports; only its presentation is gone.
- **The chart becomes the only visible object.** Its frame still encloses the plot and the
  leading-gutter strip; its title is hidden until a visible column returns (D4).
- **Selection-gated row commands may remain enabled while the table is invisible.** The active
  cell can still intersect the table's logical range. The Columns control is always available to
  restore the columns, and the Office gate records the observed command state.
- **The `HiddenColumns` value is longer but unchanged in kind.** It is the six display names in
  schema order; the key, the schema version and the catalogue hash do not move.
- **The frame documentation drift is recorded.** The code closes the live frame on `PanelLeftPt`
  (the leading gutter's outer edge, owner ruling 2026-11-16), while ADR-0061 D9 and the entity
  guide still say the frame closes on the table's first visible column. This ADR records the
  code's behaviour for the zero-column case and reports the drift; it does not silently change
  the geometry.
- **Freeze panes and print areas are user-created states.** The probe did not exercise them;
  the Office gate records the observed behaviour rather than promising support.

## Alternatives rejected

- **Keep the last-authoring-column refusal.** Rejected by the owner: it makes the chart-alone
  state impossible, and the table is still authorable through the ribbon or by restoring a
  column.
- **Hide the table by zero-width columns instead of hidden columns.** Rejected by ADR-0061: a
  zero-width column leaves a visible sliver and breaks print layout.
- **Move the user columns or their data to the configuration sheet.** Rejected by the one-sheet
  invariant and ADR-0029 D7.
- **Make the primary a toggleButton with a pressed state.** Rejected: a split button has no
  pressed state, the owner asked for the all/none action on the primary, and the label is the
  accessible channel.
- **Refuse the render when no column is visible.** Rejected: it turns a presentation choice into
  a broken chart, which is the defect this ADR removes.
- **Drop `NonPositiveWidth` so a hidden column's zero is accepted.** Rejected: that would treat a
  caller's hidden column as a real panel column and move the plot.
- **Let export keep the schema panel.** Rejected by the owner: the export must reproduce what is
  visible, including the chart-only state.
