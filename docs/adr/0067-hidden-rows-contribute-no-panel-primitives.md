# ADR-0067 — A hidden worksheet row contributes no data-panel primitives, while its plot shapes stay

- **Status:** Accepted (owner ruling 2026-11-19, during the R6.8 live defect)
- **Date:** 2026-11-19
- **Related:** R6.8 (the one-click copy), R6.2 (composition pre-flight), R6.1 (export
  bounds), guide §3/§4, ADR-0023 (the panel measures per row), ADR-0034 (zero-height rows
  are legal), ADR-0065 (zero visible columns)

## Context

Copy editable refused on every attempt for a workbook whose table has an outline-collapsed
section (Excel 16.0 build 20527 x64), while Refresh chart rendered the same workbook
normally. The add-in log recorded the reason once the writer traced it:

```text
gantt.shape create-refused id=G-<token>:panel-cell:Description kind=Rectangle refusal=InvalidGeometry step=validate b=518.362,89,45,0
gantt.shape.create-steps n=24   (identical on all four attempts)
```

`PanelRowProjection` projects every source row; Excel reports an outline-hidden row's height
as zero, and `PanelCellGrid` accepts that value deliberately (ADR-0034). `PanelBuilder.AddBody`
therefore emitted a `panel-cell:{column}` rectangle — and, for a non-blank cell, a
`panel-text:{column}` box — at height 0. `ExcelShapeWriter.Validate` refuses a non-positive
extent before any host call (`IsUsableBounds`), the composer collapsed that refusal to
`HostRejected`, and the command printed one generic sentence advising a Refresh that cannot
help. The 24 successful creates were the visible rows' cells; the refusal was the first
hidden row's.

## Decision

### D1 — A hidden row describes nothing

`PanelBuilder` emits no `panel-cell:`, `panel-text:` or `panel-border-h:` primitive for a row
whose measured height is zero. The row keeps its positional slot and its zero height in
`PanelCellGrid.RowHeightsPt`, so `Rows.Count == RowHeightsPt.Count` still holds, every later
row's running-sum top is unchanged, and `PanelBounds` is unchanged — a zero height already
contributed nothing.

### D2 — The border is dropped with the row

A hidden row's top and bottom edges coincide with the border already drawn for the row above,
so drawing it again would put two coincident lines in the editable copy.

### D3 — The rule keys on the measurement, not on `Visible`

`Visible=false` still reproduces its data-panel row (guide §4): it suppresses the entity and
its label, not the table row. A collapsed child is *hidden on the worksheet* instead — there
is no visible row to reproduce — and the measured height is the only visibility signal the
panel request carries.

### D4 — Plot-area shapes are unchanged

The owner's ruling: shapes for the collapsed row must still show from the plot area, but not
the data panel row when it is collapsed. The scene's spans, overlays, milestones,
delineators and labels are untouched; only the panel's zero-height primitives are gone.

### D5 — A refused composition names its typed reason

`CopyEditableCommand` maps `StagingCompositionOutcome.WriteRefusal` to a sentence that names
the cause, and no copy refusal advises a Refresh: Refresh cannot delete a staging shape (no
production path proves a staging tag outside `StagingLifecycle`'s own registry) and cannot
make a zero-area primitive creatable. An `ExportRefused` composition reuses the existing
export-bounds message instead of restating it generically.

## Consequences

- An editable copy of a workbook with a collapsed section composes; the pasted group's child
  count equals the number of visible panel primitives plus the chart's own.
- `ExcelShapeWriter.Create` records every refusal with the primitive id, kind, typed reason,
  step and geometry — the record that made this defect diagnosable rather than guessable.
- The scene contract tightens: no `SceneRect` or `SceneText` in an export scene has a
  non-positive extent, pinned by `PanelBuilderTests` and `ExportBoundsResolverTests`.
- A pinned primitive count or golden image that includes a collapsed fixture moves; such a
  change is reviewed, not accepted silently.

## Alternatives rejected

- **Clamp a zero extent to a minimum (for example 0.25pt).** It invents a hidden constant and
  puts an invisible, selectable artefact into the copy.
- **Filter zero-area requests in the composer.** It would leave the scene describing pixels
  that cannot exist, and every later consumer of an export scene would need the same filter.
- **Stop projecting hidden rows into `PanelRow`s.** The projection is positional against the
  measured grid; removing a row would shift every row below it into the wrong slot.
