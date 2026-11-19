# ADR-0068 — The editable copy composes with the live chart's geometry rules

- **Status:** Accepted (owner ruling 2026-11-19, from a live comparison of a copy against its
  chart). Amends ADR-0030 D6, ADR-0031 D1/D2, ADR-0038 D1, ADR-0048 and guide §3/§10.
  **Amended 2026-11-19** after a second live comparison: D4's anchor-row carry is superseded
  by D9, and D10 records the copied panel's alignment and child indentation.
- **Date:** 2026-11-19
- **Related:** R6.1–R6.8 (the editable export), [`plans/R6.8-export-live-geometry-parity-plan.md`](../../plans/R6.8-export-live-geometry-parity-plan.md)

## Context

The editable copy composes through the same `SceneBuilder` as Refresh, but the
`EditableExport` profile overrides several values that were written before the data panel
became part of the export (R6.1/R6.2). The owner compared a real copy with its chart and
reported four faults; a setting-by-setting audit against the refresh engine found three more.
The export was producing its own geometry rather than the chart's.

## Decision

### D1 — The splitter fill spans the plot

§10's band and its two borders span the plot's left edge to its right edge. The band used to
start at the data panel's left edge, which is invisible on the live sheet — the panel there is
the worksheet's own cells — and became a full-width stripe across the table as soon as the
export drew its own panel. `PanelLeftPt` remains the field the `DataPanelLeft` label position
is measured in.

### D2 — The drawn title is the live title

For a panel-drawing composition the title box is the data panel's left edge and width on the
row the year band occupies, centred within that box — exactly the box the live `ChartTitle`
textbox occupies (`GanttRefreshOrchestrator`). It is no longer a full-width band above every
band, and it therefore reserves no height above the plot. A composition with no panel uses the
plot's columns, the only table-width field it has.

### D3 — Every panel-drawing composition anchors its lanes

Lane anchoring becomes a function of the composition, not of the caller: a composition that
draws the data panel reproduces the sheet row for row, so its lanes are anchored to the same
measured grid the panel is laid out from; a panel-less composition keeps the R3.4 stacking.
`ExportBoundsResolver` refuses an unanchored request instead of composing a chart that cannot
line up, and `EditableExportRequest` anchors. The reported fault — every lane one row above its
panel row from the first row that owns no lane — disappears because `LaneRowAnchorResolver`
already returns origin-relative row tops and `SceneBuilder` adds the plot top exactly once, so
the zero-origin export grid aligns lane 1 with panel row 1 by construction. A hidden row that
*owned* a lane would be refused as a degenerate height; that cannot occur, because a collapsed
row is a child drawing on its parent's lane.

### D4 — The anchor row is carried into the copy (superseded by D9)

The export grid is rebuilt with the measured `AnchorRowHeightPt`, and the frame's bottom
margin gains it, so plot-spanning bands overhang that row and the closing line is drawn as on
the sheet. A zero-height anchor row remains a legitimate configuration and reproduces the
pre-ADR-0038 geometry. **Superseded:** a second live copy showed the chart ending a few
points below the table, because the exported panel draws no anchor row for the overhang to
sit on. D9 trims it.

### D5 — The copied panel takes the workbook's grid-line width

`ExportPanelTheme` resolves the border width from the workbook's `GridLinePt` metric, falling
back to the catalogue default only when the value is missing or unusable. A user who changed
"Grid line width" must not get a chart and a copy of two different weights.

### D6 — Geometry parity is a rule, not a coincidence

Any future difference between the two profiles must be one of: the worksheet rows the live
chart stands on, the cells that are the live panel, or the sheet coordinate origin. A
difference in a *drawing* rule is a defect in the copy.

### D7 — The plot spans the rows the panel reproduces

The export plot's height is the sum of the measured heights of the rows the panel actually
reproduces (one per event), not the whole measured table. A trailing table row that yields no
event — blank, or a blocked row dropped from a partial render — made the plot one row longer
than the panel, and the two bottoms stopped aligning (owner report 2026-11-19: the panel is
fixed, so the plot follows it).

### D8 — A composition with no panel draws no title

The drawn title sits over the data panel, so a chart-only composition has no table to title and
draws none. The rule lives in `ExportBoundsResolver`, where the panel decision is made, and
mirrors the live rule that disables the title when no column is visible (ADR-0065 D4). It is
not a `SceneBuilder` rule: a caller that composes without the resolver keeps its title, which
the equivalence table requires.

### D9 — The copy ends where the panel ends (supersedes D4)

The exported panel and the exported plot are **one row set**. The factory truncates the export
grid's measured row heights to `events.Max(RowNumber)` — the rows the panel reproduces — and
derives both the panel's bottom and the plot's height from that same grid, so their bottom
edges coincide by construction rather than by a second sum. The export carries **no anchor
row**: `ChartAnchorRowHeightPt` is zero, so the plot-spanning bands stop on the plot's own
bottom edge, no closing line is drawn, and the frame's bottom margin is the same uniform outer
padding as its other three sides. The live arrangement is unchanged — it still overhangs the
reserved anchor row because a worksheet insert needs that cell anchor.

### D10 — The copy reproduces the table's alignment and child indentation

The exported panel reads `GanttTableAppearance.SceneAlignmentFor` for every measured column,
so `Start`, `Finish` and `Duration` are centred in the copy exactly as they are on the sheet.
A grouped child's `Description` cell reproduces the live `IndentLevel`: `PanelRowProjection`
derives the same outline plan the worksheet uses and carries each row's indent on
`PanelRow.IndentLevel`; `PanelBuilder` writes it as `SceneText.LeftInsetPt` on the
Description text only; the Office renderer consumes it as `TextFrame2.MarginLeft`. The
one-step width is `HierarchyIndentation.ChildIndentWidthPt`, a code constant rather than a
catalogue metric so the catalogue hash does not move.

### D11 — The plot's bottom border is drawn in every composition

`chart:plot:bottom` is emitted for the live and the export compositions alike. D9 removed the
copy's anchor row, and therefore its closing line, so the copy's plot was left open at the
bottom (owner report 2026-11-19: the copied chart's bottom rule disappeared). On the live sheet
the edge is inset by half its stroke **inside** the last body row, because that row's own fill
clips the outer half. An export composition draws its own panel and has no cells to clip it, so
the line lands exactly on `PlotBounds.Bottom` — the same edge as the panel's last row, closing
the copied table and plot with one shared rule. The closing line keeps its own meaning: it
terminates the band overhang and exists only while `ChartAnchorRowHeightPt` is non-zero.

## Not yet in force

**Label-position parity** (the fourth reported fault: the delineator's description label
flipping above the line in the copy). D9 removed the anchor-row overhang, so it no longer
supplies the below-line room D4 expected; the corner is re-measured on the live copy before
any further rule is added.

**One measured consequence of D3:** the reference fixture numbered its rows by worksheet row
(2 upwards) while `GanttEvent.RowNumber` is the reader's one-based body index, so the export
test fixture re-bases them before composing. Production readers already send body indices, and
`PanelRowProjection`/`LaneRowAnchorResolver` both require that spelling.

## Consequences

- The copy's chart bounds lose the title band's height above the plot and no longer add the
  anchor row below it; the staged group and the live paste-size assertions move with them.
- `SplitterBuilder.BandBounds` is now the plot's rectangle, so a caller that read it for a
  panel-spanning extent must use `PanelLeftPt` through the label field instead.
- The panel's border weight follows the workbook, so a themed chart and its copy stay equal.

## Alternatives rejected

- **A panel-spanning splitter fill.** It is what the code did and what the guide's wording
  implied; the owner's ruling is that the section fill marks the lane, and the table is not
  the lane.
- **Clamping the title to a full-width band.** It reads as a separate heading rather than the
  table's title, which is the fault being fixed.
- **Leaving the anchor row at zero for the export.** It was a deliberate R6.3 choice ("no
  worksheet rows to anchor to"); with the panel drawn, the copy has rows and must show the
  overhang the chart shows.
