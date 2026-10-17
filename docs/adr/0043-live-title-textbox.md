# ADR-0043 — The live chart title is an owned textbox shape, not a reserved-row cell

- **Status**: Proposed (owner approved the direction 2026-10-17; awaiting implementation)
- **Date**: 2026-10-17
- **Context**: R5.6b ships the Chart title dialog writing the `ChartTitle`/`ShowTitle`
  settings. The owner further requires the live title to render in a textbox shape in
  the row above the table (using the heading space), with the current cell placeholder
  removed. The binding contract today says the opposite: ADR-0030 D4/D6 and R4.7I
  write the title into reserved row 1 as a cell (`ExcelWorkbookInitialiser.WriteTitleRow`
  + `RollBackTitleRow`), and `SceneBuilder` emits no title primitive under
  `SceneCompositionProfile.LiveExcel` — the drawn title band exists only for export
  compositions (entity guide §2, `FrameBandsBuilder` ellipsis + warning).
- **Decision**:
  - **D1 — Live title is one owned shape.** Refresh creates/updates/deletes a single
    textbox in the reserved heading row, carrying the trimmed `ChartTitle` text,
    centred/middle-aligned, `Title` style 12pt bold `#404040`, width spanning the
    table/plot width, height `TitleBandHeightPt`. `ShowTitle=false` deletes (or hides)
    the shape. Ownership follows ADR-0019 (`AlternativeText` tag, `ZLayer.Frame`).
  - **D2 — The cell placeholder is removed.** `WriteTitleRow` no longer writes title
    text into reserved row 1 and `RollBackTitleRow` no longer clears a title cell.
    Legacy workbooks carrying a title cell migrate through the explicit Repair path
    (ADR-0011): clear the cell, create the shape on the next Refresh/repair — never
    silently on open.
  - **D3 — Settings contract unchanged.** `ChartTitle` stays trimmed non-blank
    (even when hidden, so `ShowTitle=true` restores the retained title) with no
    length cap; available space is the constraint and over-long titles truncate with
    the §2 ellipsis + warning, never by shrinking the font. `ShowTitle` stays the
    only title-band visibility control (ADR-0014).
  - **D4 — Export unchanged.** Export compositions keep the §2 title band
    (`FrameBandsBuilder`); the live shape and the export band share the same
    truncation measurement (the shared 8pt seam; the 12pt-vs-8pt residual already
    accepted for the title).
- **Consequences**: `ExcelWorkbookInitialiser`, the refresh shape pipeline (new live
  title port over `ExcelShapeWriter` patterns), `SceneBuilder`'s LiveExcel no-title
  rule, `ChartFrameGeometry`/`PanelCellGrid` anchoring, the equivalence field table,
  the golden fixture, and `ExcelSceneBuildRequestFactory` title plumbing all change
  in the implementing commit. COM runs on the Excel main STA thread, one tested
  ownership helper, no `Task.Run` around COM, five-setting save/restore in
  `try/finally`. The workbook schema version bumps per ADR-0029 D5 (last change in
  the row's commit sequence) because the reserved-row contract moves.
- **Alternatives considered**:
  - Keep the live title as a cell (rejected by the owner 2026-10-17: the title must
    live in a textbox using the heading space).
  - Dual-write cell + shape (rejected: two authorities for one title; cell and shape
    would disagree after any partial failure).
