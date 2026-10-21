# ADR-0048 - The live chart's border encloses the visible data table

- **Status**: Proposed (owner reported the symptom 2026-10-21)
- **Date**: 2026-10-21
- **Work item**: R5.6e (the chart frame encloses the table)
- **Amends**: ADR-0043 D1 (the live title's placement); ADR-0030 D4/D5 (the live sheet's reserved rows)

## Context

The owner reported that the outside border still did not wrap around the chart title and the table, and asked for the borders to be extended to the known start cells of the visible table horizontally, with a vertical border at the start of the visible table.

The cause is a split between two halves of the same chart.

A **live** chart's title is an owned textbox shape placed in the **table's own columns** (ADR-0043): the title and the table therefore sit to the **LEFT** of the plot. The frame, however, was drawn on the chart's own content rectangle, and for a live chart that rectangle is the **plot**. `SceneBuilder` unions the panel's bounds into the content only when a panel was actually drawn, and a live request supplies no panel - the worksheet's own cells *are* the panel (R4.8A D4) - so the frame fell back to `plotBounds` alone.

The result was a border drawn from the plot's left edge to its right edge. It enclosed neither the title nor the table, and the two objects read as unrelated.

## Decision

### D1 - The border is drawn on the union of the chart and the visible table

The live frame is drawn on the union of the chart's rectangle and the **measured** visible-table rectangle. The union is **one** rectangle: the chart's own top and bottom (its margins), the plot's right, and the table's first column on the left. The top and bottom rules therefore extend across the columns between the table and the plot, and a vertical rule closes the rectangle at the table's start.

The owner chose the straight extension over a stepped border, and accepted the top rule crossing the columns between the table and the plot - it has to, to reach the table.

### D2 - Only the BORDER uses the union

The background, the alternate bands, the grid lines and the closing line keep the plot's own rectangle. A background over the table would **paint the user's own cells**, which is a data-visibility defect, not a styling one. The border is the only thing that grows.

### D3 - The table's rectangle is MEASURED, and it is the same measurement the title uses

The rectangle is built from the measured `PanelCellGrid`: the left edge is the panel's origin, the right edge is the origin plus the measured column widths, the top is the header row's top and the bottom is the last body row's bottom. The left edge is **the same number the live title textbox is placed at**, so the border's vertical rule and the title's left edge cannot drift apart.

The measured origin is the table's first **visible** column: the engine columns before it are hidden and contribute no width, so their worksheet positions do not move it.

### D4 - The export compositions are unchanged

They draw their own panel, so they pass no table rectangle and their border stays exactly where it was. The **golden scene snapshot does not change**, which is the proof that this is a live-composition change and not a change to the chart model.

## Consequences

- The title, the table and the plot are enclosed by one border, so the chart reads as one object.
- The frame's **top and bottom are unchanged**: the chart's own margins, two cells below the Gantt (ADR-0047 D2).
- The visible table's cells are not painted over.
- The table's **own** borders (Excel's table style) are still drawn, and the frame's left rule sits on the table's left edge. Removing the table's own visuals is a **separate, later step** the owner has scheduled; until then the two overlap deliberately.
- The frame follows the table wherever it is, because the origin is measured. A table that is not at the sheet's left edge is enclosed correctly.

## Alternatives considered

- **Extend the background as well.** Rejected: it would paint the user's own cells.
- **Move the live title into the plot's columns.** Rejected: the owner asked for the title in the table's heading row (ADR-0043), and the title is what made the gap visible.
- **A stepped border** - the bottom edge stopping at the table's last row and stepping up to the chart's bottom at the plot. Rejected by the owner: one straight rectangle is what "a uniform frame" means.
- **A code constant for the table's first column, with the measurement refused on a mismatch.** Rejected: the measured origin is the truth of where the table is, and a constant would be a **second** authority for the same fact. A check would also refuse a legitimately placed table, failing the whole refresh for a cosmetic reason. The frame following the measurement is the stronger guarantee, not the weaker one.
- **Writing the border into the table's own cell borders.** Rejected: the add-in would then rewrite the user's table style on every Refresh. The table's own visuals are a later, separate decision.
