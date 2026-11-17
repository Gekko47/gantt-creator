# ADR-0059 — The visible columns lead the table, and the table starts in column B

- **Status:** Accepted (owner-approved R5.18k, decision D2/D6 recorded 2026-11-13)
- **Date:** 2026-11-13
- **Amends:** the remark on `GanttTableSchema.Default` that put the authoring columns in
  the middle of the engine columns; ADR-0029 D6 (no migration) is reaffirmed, not changed
- **Related:** ADR-0029 (schema contract and versioning), ADR-0030 (reserved rows and the
  plot anchor), ADR-0052 (host-owned geometry), R5.18k, R5.18l

## Context

Two rules had grown into each other.

`GanttTableSchema.Default` listed `Id`, `LaneId` and `StackIndex` first and put the five
columns a user reads and edits (`Type`, `Description`, `Start`, `Finish`, `Duration`)
behind them. The remark justified that order: moving the visible columns to the front
"would have been a second, unrecorded column-order change in the same schema bump". The
engine columns are hidden, so the first column a user *sees* was `Type` anyway — the
order cost nothing on screen.

It cost something else. Every positional test literal and every positional array in the
Office layer had to know that `Id` was column 0, and the first repair path that wrote a
header from a schema position would have been right only by coincidence.

The second rule is the frame. `SceneBuilder` draws the chart frame around the table, and
the table's top and bottom padding rows give the frame a margin above the first row and
below the last. There is no equivalent margin on the left: the frame starts at the table's
first column, so that column sits flush against the frame while the rows above and below
it do not. The owner asked for the same margin on all four sides (R5.18k, D6).

## Decision

1. **The five visible columns lead the table.** The order becomes `Type`, `Description`,
   `Start`, `Finish`, `Duration`, then `Id`, `LaneId`, `StackIndex`, `ParentId`,
   `SiblingOrder`, `StyleKey`, `LabelPosition`, `DateLabelPosition`, `FillColour`,
   `StrokeColour`, `Visible`, `SortOrder`. No column is added or removed, so the counts
   are unchanged and the order is not.
2. **The table starts in column B.** `GanttSheetLayout.TableOriginColumnIndex` is 2, and
   column A is `GanttSheetLayout.LeadingGutterColumnIndex` — the add-in's spacer gutter.
   It is **not** a column of `tblGanttData`.
3. **The leading gutter is add-in owned.** It must never be reported as an "extra user
   column" by the integrity checker, and it is not a place a user's data belongs.
4. **Every derived column index is offset by the origin.** `TitleColumnIndex` is the
   schema position plus the origin; `PlotAnchorColumnIndex` is the origin plus the column
   count. A bare schema position names a worksheet column one place to the left of the
   table it belongs to.
5. **The schema version goes 22 → 23.** A version-22 workbook reports a mismatch.

## No migration

There is no migration path and none is written. `Initialise` is the only creation path
and there are no active users (ADR-0029 D6, reaffirmed). A workbook created under schema
22 has its columns in the old order and its table in column A; opening it reports the
version mismatch and the remedy is Initialise, exactly as every previous schema bump.

## Consequences

- **The catalogue hash moves.** `GanttCatalogues.ComputeCatalogueHash()` covers the schema
  version, so the pinned first-release hash changes with this ADR. That is why R5.18k is
  the commit that re-pins it, and why the two cannot be separated: a re-pin without the
  version bump would hide the mismatch, and a bump without the re-pin would fail the gate
  with no explanation.
- **Positional literals became lookups.** Tests that indexed the schema by number were
  rewritten to resolve a column by name. A test that restates the schema order cannot
  tell a deliberate reorder from a bug in it.
- **The gutter's width and the frame's enclosure are the next slice, not this one.** This
  ADR fixes *where* the gutter is and who owns it. Making it the same width as the
  framing rows, measuring that width back, and extending `LiveDataPanelBounds` so the
  frame encloses it are a separate change with its own golden-scene commit, because the
  measured width is what moves the scene.
- **The plot anchor moves right by one.** Every saved `PlotAnchor` defined name from schema
  22 names the wrong column; the mismatch above is what surfaces that rather than a silently
  misplaced chart.
