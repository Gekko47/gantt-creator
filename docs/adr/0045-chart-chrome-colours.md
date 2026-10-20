# ADR-0045 — The chart chrome colours are persisted, editable settings

- **Status**: Proposed (owner approved direction 2026-10-19)
- **Date**: 2026-10-19
- **Work item**: R5.6a (named style presets) / R5.6 (style and theme settings)
- **Supersedes/amends**: nothing; extends ADR-0007 D7 (catalogue token authority) and ADR-0014 (chart title settings)

## Context

The entity guide's §1, §5, §6, and §7 name the tokens that colour the chart's own frame: `ChartBackground` for the background, `AlternateBandFill` for the alternating band, `MinorGridStroke` and `MajorGridStroke` for the grid, `YearHeaderFill` for the year band. The catalogue publishes all five (`GanttCatalogues.Colours`), and the **reference scene fixture already used the token values** for the band and the minor grid.

The live path did not. `ExcelSceneBuildRequestFactory.FrameTheme()` **hard-coded seven colour literals** and read none of the tokens:

| Frame member | Factory literal | Catalogue token | Token value |
|---|---|---|---|
| Background | `#FFFFFF` | `ChartBackground` | `#FFFFFF` |
| AlternateBand | `#F5F5F5` | `AlternateBandFill` | `#F2F2F2` |
| MinorGrid | `#E0E0E0` | `MinorGridStroke` | `#D9D9D9` |
| MajorGrid | `#BFBFBF` | `MajorGridStroke` | `#000000` |
| YearHeader | `#D9D9D9` | `YearHeaderFill` | `#D9D9D9` |
| PeriodHeader | `#E6E6E6` | *(no token of its own)* | — |

So the live chart's frame was a **third** set of values, disagreeing with both the guide and the reference fixture. Nothing in the product could change them, and the two "grid" widths the Chart tab exposed (`GridLinePt`, `MajorBoundaryPt`) were **inert**: the frame theme carried its own literal widths, so the controls named metrics that nothing consumed.

## Decision

### D1 — The chrome colours are settings, not literals

The six chart-chrome colours are **settings** on `tblGanttSettings`, named for the role they play: `ChartBackgroundColour`, `AlternateBandColour`, `MinorGridColour`, `MajorGridColour`, `YearHeaderColour`, `PeriodHeaderColour`. They are settings rather than style rows because they belong to **the chart as a whole**, not to any entity's style — the same reason the chart title and the banding and grid switches are settings.

### D2 — The catalogue token is the single authority for each default

Each setting's default **is** its approved token value, read from `GanttCatalogues.Colours` rather than restated. A test pins the two equal, so a token edited without its setting (or the reverse) cannot ship. The factory's literals are gone: the frame resolves from the workbook's settings, falling back to the token value.

The period header band has no token of its own — the guide's §6 says its "alternating neutral fills may follow plot bands", and the reference scene gives it the band fill — so it borrows `AlternateBandFill` while keeping its own setting key, so a user can still give it a colour of its own.

### D3 — The grid widths reach the render

`GridLinePt` and `MajorBoundaryPt` now supply the frame theme's minor- and major-grid outline widths. They were already the catalogue's authority for those widths and already editable; only the frame theme's literals were bypassing them.

### D4 — A malformed colour is refused, never substituted

The reader refuses a chrome colour that is not uppercase `#RRGGBB` (`BadColourFormat`), and the writer's validation refuses the same. The entity guide's rule is that a colour is never replaced behind the user's back, and a substituted chart colour would be exactly that.

### D5 — Schema version 20, and the config table's anchor moves

The settings key set grows from 19 to 25, so a version-19 workbook reports a mismatch and the remedy is Initialise (ADR-0029 D6, no migration). The settings table's extent grows to row 95, so the config table's anchor moves from **A92 to A98** to keep the blank row that Excel's table auto-expand needs (the 2026-10-09 defect). The blank-row invariant test ties the anchors to the catalogue sizes, so the next addition fails there first.

## Consequences

- The chart's frame, bands, grid, and headers are the user's to change, from the **Chart appearance** group on the Chart tab.
- The "Grid line width" and "Major boundary width" controls now do something.
- The live chart's default frame matches the guide's approved tokens rather than the old literals, so a workbook re-initialised after this change renders the token colours.
- A version-19 workbook must be re-initialised.

## Alternatives considered

- **A new `tblGanttColours` table.** Rejected: a sixth contract table needs its own anchor, headers, reader and writer path, and preservation rules, for six values the settings table already carries in the same key/value shape.
- **Leaving the frame's literals and adding the colours beside them.** Rejected: that would keep two authorities for the same colour, which is the defect this item exists to remove.
- **Persisting the colours but not the widths.** Rejected: the widths were already editable and already inert; fixing only the colours would leave a control that does nothing.
