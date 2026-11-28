# ADR-0075 — Retire the plot grid and the controls that fed it

- **Status:** Accepted (owner ruling 2026-11-27)
- **Related:** entity guide §2, §5, §6, §7, §8, §24; ADR-0037 D1; ADR-0038 D2; ADR-0046; ADR-0061 D8; ADR-0068 D5; ADR-0029 D6; ADR-0073; R11.1

## Context

The plot drew a vertical grid line at every period boundary, plus a major line at each December
boundary, with two workbook switches (`ShowMinorGrid`, `ShowMajorGrid`) and a chrome colour
(`MinorGridColour`) and a width metric (`GridLinePt`). The owner's 2026-11-27 ruling is that the
vertical alternating bands are the period separation and the plot should carry no interior lines.
Because the plot's left and right edges were gated by `ShowMajorGrid`, the same ruling also
required those edges to become unconditional (ADR-0076).

## Decision

- **D1 — No interior vertical lines.** `AddGrid` emits only `chart:grid:edge:left`,
  `chart:grid:edge:right`, `chart:plot:top` and `chart:plot:bottom`. `AddGridLine` and `PeriodKey`
  are deleted; the identifier-stability work they existed for has no lines left to stabilise.
- **D2 — The bounds are structural, not settings.** `FrameBandsRequest` loses `GridLinePt`,
  `ShowMinorGrid` and `ShowMajorGrid`; `SceneBuilder` loses the same three properties. The request
  is positional and carries two adjacent widths, so the three are removed in one edit — a partial
  edit can rebind `MajorBoundaryPt` silently.
- **D3 — The orphaned controls retire with their consumers.** `FrameBandsTheme.MinorGrid` is
  removed (nothing drew it), and with it the `MinorGridColour` setting and the
  `GanttChartChromeRole.MinorGrid` role, and the `GridLinePt` metric. `ThemeEditRequest` and the
  theme dialog lose the "Grid line width" control; the "Grid" metric group becomes "Borders" with
  `MajorBoundaryPt` alone. Leaving either control in place would repeat the ADR-0045 defect of a
  user-visible control with no effect.
- **D4 — The export panel border follows the major boundary.** `ExportPanelTheme` reads
  `MajorBoundaryPt` instead of the retired `GridLinePt`, superseding ADR-0068 D5. That is the same
  weight every border on the chart takes.
- **D5 — The colour token survives under a narrower meaning.** `MinorGridStroke` is NOT retired:
  `GanttTableAppearance.BorderColour` reads it for the live table and the export panel. Its name now
  describes a role it no longer plays; renaming it is a separate schema change and is out of scope.
- **D6 — Schema version 30, no migration.** `tblGanttSettings` loses three rows and
  `tblGanttMetrics` loses one. `AdditiveSchemaUpgrade` inspects settings keys only and refuses a
  longer stored list as `ExtraKey`, so a version-29 workbook reports the confirmation-required
  `SchemaVersionOlder`; the metrics removal shows as `CatalogueHashMismatch`. Repair rewrites both
  catalogues (ADR-0029 D6).

## Consequences

- The plot's period structure is `AlternateBandFill` alone. The owner confirmed that is intended.
- The bands, the plot's edge lines and the delineators are the plot-spanning shape set; ADR-0037
  D1 and ADR-0038 D2 lose grid lines from their lists, and ADR-0046's year-boundary weight
  hierarchy has no line left to express.
- Every version-29 workbook must be repaired once. This is the same class as version 9's
  `MaximumExternalLabelWidthPt` retirement.
