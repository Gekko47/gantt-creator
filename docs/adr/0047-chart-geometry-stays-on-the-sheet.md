# ADR-0047 - The live chart's geometry must stay on the sheet

- **Status**: Proposed (owner reported the symptom 2026-10-21)
- **Date**: 2026-10-21
- **Work item**: R5.6d (theme dialog and chrome fixes)
- **Amends**: ADR-0030 D4/D5 (the live sheet's reserved rows); ADR-0043 (the live title's placement)

## Context

The owner reported that the chart's background and its far-right border "drop down too far" - they used to be locked two cells below the Gantt, and after a change they sat lower than that. The same report said the outside border still did not wrap around the chart title and the table.

The change that caused it reserved the **title band** in the frame geometry, so the frame's top edge moved up by one title band (24pt) above the row the frame had started on. On a live sheet that put the frame's top edge at **-24pt**: above row 1.

That is not drawable. Excel **clamps a shape's negative top to the top of the sheet**, and a clamped top with an unchanged height moves the shape's **bottom** down by the same amount. The frame's background and its right-hand edge therefore ended 24pt lower than the geometry said - which is exactly the "drop down too far" the owner saw, and the frame stopped being two cells below the Gantt.

Two further facts belong in the same record, because they are the same class of defect - **two authorities for one row**:

- The live title is an owned textbox shape drawn in the **reserved row** (ADR-0043). Its height came from the `TitleBandHeightPt` token (24pt by default) while the reserved row it is drawn in is normalised to `YearBandHeightPt` (18pt). The two disagreed, so the textbox began **6pt above** the reserved row, in the top padding row, and the frame's top edge - the reserved row's top - cut straight through the title instead of enclosing it. That is why the border did not wrap the title.
- The row-height normalisation passed `GanttCatalogues.MetricDefault(...)` for all seven heights, so the sheet was normalised to the catalogue's defaults whatever the workbook said. The Theme dialog's "Row height", "Splitter height" and "Spacer height" were written to the metrics table and then ignored on every Refresh, while the scene factory read the workbook's metrics - so the sheet and the chart disagreed about how tall a row is (the exact disagreement row-height normalisation exists to prevent, ADR-0026 D2/D3).

## Decision

### D1 - The title occupies the reserved row, and the reserved row is its only authority

The live title's bounds are derived from the **measured** reserved row: its top is the top padding row's height and its height is `OriginTopPt - HeaderHeightPt - TopPaddingHeightPt`. The `TitleBandHeightPt` token is no longer a second authority for the live title's height; it remains the export composition's title strip, which is its own row in that composition.

The title and the year band then share one height and one top edge, because they share one row - which is what the owner asked for when the title was given "the heading space".

### D2 - The live frame is held to the sheet, and the rule is two-sided

A live chart's coordinates are the sheet's own, so:

- **The content must be on the sheet.** A band above row 1 cannot be drawn where the model says, and Excel would clamp it - moving the band down over the row it belongs to. The build is **refused** (`SceneBuilderRefusal.ChartAboveSheet`), not silently misplaced.
- **The frame's own top margin may reach above the sheet, and is then clamped.** The frame is shortened **at the top** and its height is reduced by the same amount, so its **bottom edge does not move**. This is the fix for the reported symptom: the frame now ends where the geometry says it ends, whatever its top margin reaches.

The export compositions build in their own coordinate space and are translated to a zero origin by the renderer, so a negative origin is legitimate there and neither rule applies.

### D3 - The workbook's metrics are the authority for the sheet's row heights

The normalisation reads the **workbook's** metrics for all seven heights, falling back to the catalogue's default for a metric the workbook does not carry. A metric the user has set is therefore applied; a metric the workbook predates still normalises to a defined height.

## Consequences

- The chart's background and its far-right border end where the geometry says, two cells below the Gantt, whatever the frame's top margin is.
- The chart's outside border encloses the title: the title sits in the reserved row, which is the row the frame's top edge is on.
- The Theme dialog's row, splitter and spacer heights take effect on the sheet, and the sheet and the chart agree about them.
- A live chart whose measurement would put content above row 1 is refused with a named reason rather than drawn in the wrong place. A live sheet never does this: the reserved row and the header row are the band rows, so the content's top is the top padding row's height, which is on the sheet.
- The live profile now refuses a chart it could previously compose, so a malformed sheet reports "the chart could not be composed: ChartAboveSheet" instead of drawing a displaced chart.

## Alternatives considered

- **Clamp the whole chart down by the overflow.** Rejected: the chart is aligned to the table's rows, so shifting it down moves the plot off the rows it belongs to. Refusing is honest; shifting is a different chart.
- **Keep the title band reservation and let Excel clamp.** Rejected: that is the defect. Excel's clamp is silent, and the model and the sheet then disagree about where the chart ends.
- **Make the reserved row's height follow the `TitleBandHeightPt` token instead.** Rejected: the reserved row also carries the year band, whose height is the year band's token. One row, one height, and the height is the row's own.
- **Migrate the stored metrics silently on read.** Rejected: a user-editable setting is never replaced behind the user's back (ADR-0029 D6).