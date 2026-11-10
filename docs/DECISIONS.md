# Decisions log (ADR index)

> Lightweight Architecture Decision Records. One decision per file in `docs/adr/`.
> This file is the index. The file body in each ADR is the durable record.

| ID | Title | Status | Date | File |
| --- | --- | --- | --- | --- |
| ADR-0000 | Adopt the Gantt Creator development kit as the source of truth | Accepted | 2026-09-04 | [`adr/0000-adopt-kit.md`](adr/0000-adopt-kit.md) |
| ADR-0001 | Self-hosted Windows runner as the Office-integration host | Accepted | 2026-09-04 | [`adr/0001-self-hosted-office-runner.md`](adr/0001-self-hosted-office-runner.md) |
| ADR-0002 | Commit `.clinerules/*` and `AGENTS.md` as canonical sources | Accepted | 2026-09-04 | [`adr/0002-commit-clinerules-and-agents.md`](adr/0002-commit-clinerules-and-agents.md) |
| ADR-0003 | Visual Studio 2026 is the verification IDE; Visual Studio 2022 baseline is relaxed | Accepted | 2026-09-04 | [`adr/0003-visual-studio-2026.md`](adr/0003-visual-studio-2026.md) |
| ADR-0004 | Pin Excel-DNA packages to AddIn 1.9.0, Integration 1.9.0, Interop 16.0.0 | Accepted | 2026-09-04 | [`adr/0004-exceldna-package-versions.md`](adr/0004-exceldna-package-versions.md) |
| ADR-0005 | Document the actual host as Windows 11 25H2 (build 26200) | Accepted | 2026-09-04 | [`adr/0005-windows-host-build-number.md`](adr/0005-windows-host-build-number.md) |
| ADR-0006 | Reject the 1904 workbook date system on read | Accepted | 2026-09-19 | [`adr/0006-reject-1904-date-system.md`](adr/0006-reject-1904-date-system.md) |
| ADR-0007 | Configuration-sheet catalogues: Excel Tables, `tblGanttSettings`, and the catalogue hash | Accepted | 2026-09-21 | [`adr/0007-config-sheet-catalogues.md`](adr/0007-config-sheet-catalogues.md) |
| ADR-0008 | Destructive-command policy: no undo, warned confirmation | Accepted | 2026-09-22 | [`adr/0008-destructive-command-policy.md`](adr/0008-destructive-command-policy.md) |
| ADR-0009 | Named-style capability columns | Accepted | 2026-09-23 | [`adr/0009-style-capability-schema.md`](adr/0009-style-capability-schema.md) |
| ADR-0010 | First-live-slice sequencing exception | Accepted | 2026-09-23 | [`adr/0010-first-live-slice-sequencing.md`](adr/0010-first-live-slice-sequencing.md) |
| ADR-0011 | Safe configuration and identity repair | Accepted | 2026-09-24 | [`adr/0011-config-safe-repair-policy.md`](adr/0011-config-safe-repair-policy.md) |
| ADR-0012 | Derive effective stack index in Core | Accepted | 2026-09-25 | [`adr/0012-derived-effective-stack-index.md`](adr/0012-derived-effective-stack-index.md) |
| ADR-0013 | Chart-level scene ownership | Accepted | 2026-09-25 | [`adr/0013-chart-level-scene-ownership.md`](adr/0013-chart-level-scene-ownership.md) |
| ADR-0014 | Versioned frame and title settings | Accepted | 2026-09-25 | [`adr/0014-frame-band-settings.md`](adr/0014-frame-band-settings.md) |
| ADR-0015 | Label `Auto` cascade order and the widest-gap truncation fallback | Accepted | 2026-09-26 | [`adr/0015-label-cascade-and-widest-gap-fallback.md`](adr/0015-label-cascade-and-widest-gap-fallback.md) |
| ADR-0016 | Approved date display format for event dates | Accepted | 2026-09-26 | [`adr/0016-date-display-format-setting.md`](adr/0016-date-display-format-setting.md) |
| ADR-0017 | Shared scene ownership for deduplicated entities | Accepted | 2026-09-26 | [`adr/0017-shared-scene-ownership.md`](adr/0017-shared-scene-ownership.md) |
| ADR-0018 | Text alignment is its own type, not a label position | Accepted | 2026-09-26 | [`adr/0018-text-alignment-is-not-a-label-position.md`](adr/0018-text-alignment-is-not-a-label-position.md) |
| ADR-0019 | Shape ownership carrier is `AlternativeText`, with a bounded versioned tag | Accepted | 2026-09-27 | [`adr/0019-shape-ownership-carrier.md`](adr/0019-shape-ownership-carrier.md) |
| ADR-0020 | The application-state scope covers five settings, not six | Accepted | 2026-09-27 | [`adr/0020-application-state-scope-five-settings.md`](adr/0020-application-state-scope-five-settings.md) |
| ADR-0021 | Structural rows reach lane layout; "empty" means no visible entity | Accepted | 2026-09-28 | [`adr/0021-structural-rows-and-empty-scene.md`](adr/0021-structural-rows-and-empty-scene.md) |
| ADR-0022 | The §10 splitter band is a new scene builder on the existing style pipeline | Accepted | 2026-09-28 | [`adr/0022-splitter-band-builder.md`](adr/0022-splitter-band-builder.md) |
| ADR-0023 | The panel measures per row, and the panel's bounds are derived | Accepted | 2026-09-28 | [`adr/0023-panel-measures-per-row.md`](adr/0023-panel-measures-per-row.md) |
| ADR-0024 | Measure each row; a read-only measurement does not refuse a protected sheet | Accepted | 2026-09-28 | [`adr/0024-per-row-measurement-and-read-only-protection.md`](adr/0024-per-row-measurement-and-read-only-protection.md) |
| ADR-0025 | REV5 and REV6 enter the source-of-truth order as tracked inputs | Accepted | 2026-09-29 | [`adr/0025-rev5-rev6-as-tracked-governance-inputs.md`](adr/0025-rev5-rev6-as-tracked-governance-inputs.md) |
| ADR-0026 | Fixed managed row heights; live lane auto-growth is removed | Accepted | 2026-09-29 | [`adr/0026-fixed-managed-row-heights.md`](adr/0026-fixed-managed-row-heights.md) |
| ADR-0027 | The critical interval is a filled rectangle at half the predetermined activity height | Accepted | 2026-09-29 | [`adr/0027-critical-interval-filled-rectangle.md`](adr/0027-critical-interval-filled-rectangle.md) |
| ADR-0028 | Horizontal-only label positions; the existing truncation is retained | Accepted | 2026-09-29 | [`adr/0028-horizontal-only-label-positions.md`](adr/0028-horizontal-only-label-positions.md) |
| ADR-0029 | Workbook schema versions 4, 5 and 6: hierarchy, derived Duration, fixed row geometry, horizontal-only labels | Accepted | 2026-09-29 | [`adr/0029-workbook-schema-version-4.md`](adr/0029-workbook-schema-version-4.md) |
| ADR-0030 | The live chart is anchored to the worksheet rows, not to page coordinates | Accepted | 2026-10-01 | [`adr/0030-live-chart-vertical-anchoring.md`](adr/0030-live-chart-vertical-anchoring.md) |
| ADR-0033 | Labels: no fill, no border, no margin, one row tall, and stacks are not collisions | Accepted | 2026-10-02 | [`adr/0033-label-appearance-and-stack-collision.md`](adr/0033-label-appearance-and-stack-collision.md) |
| ADR-0034 | Lanes are anchored to their worksheet rows, the content block is centred, and the critical interval is top-aligned | Accepted | 2026-10-02 | [`adr/0034-lane-anchoring-centring-critical-top-align.md`](adr/0034-lane-anchoring-centring-critical-top-align.md) |
| ADR-0035 | Labels are bounded only by available space, the bottom margin is a reserved row, and an insert is conditional on the active cell | Accepted | 2026-10-02 | [`adr/0035-reserved-bottom-padding-row-and-unbounded-label-width.md`](adr/0035-reserved-bottom-padding-row-and-unbounded-label-width.md) |
| ADR-0036 | Every insert shifts the sheet: a worksheet row, never `ListRows.Add(position)` | Accepted | 2026-10-03 | [`adr/0036-append-inserts-the-worksheet-row-before-claiming-it.md`](adr/0036-append-inserts-the-worksheet-row-before-claiming-it.md) |
| ADR-0037 | The plot-spanning shapes are anchored into the header row so they stretch | Accepted | 2026-10-03 | [`adr/0037-plot-bands-anchored-into-the-header-row.md`](adr/0037-plot-bands-anchored-into-the-header-row.md) |
| ADR-0038 | An anchor row below the body lets the plot's BOTTOM stretch, closed by its own line | Accepted (amended 2026-10-03) | 2026-10-03 | [`adr/0038-anchor-row-below-the-body-for-the-plot-bottom.md`](adr/0038-anchor-row-below-the-body-for-the-plot-bottom.md) |
| ADR-0039 | The month period label steps down MMM → MM when the band is too narrow | Superseded by ADR-0040 | 2026-10-09 | [`adr/0039-period-label-fit-ladder.md`](adr/0039-period-label-fit-ladder.md) |
| ADR-0040 | The Month period-label format is the user's explicit choice (MM or MMM) | Accepted (amended 2026-10-10 by ADR-0041) | 2026-10-09 | [`adr/0040-period-label-format-selection.md`](adr/0040-period-label-format-selection.md) |
| ADR-0041 | Period-label visibility is measured-fit, so MM shows where MMM does not | Accepted | 2026-10-10 | [`adr/0041-period-label-measured-fit.md`](adr/0041-period-label-measured-fit.md) |
| ADR-0042 | Destructive commands offer a session-scoped "don't ask again" opt-out, amending ADR-0008 D3 | Accepted | 2026-10-13 | [`adr/0042-destructive-confirmation-opt-out.md`](adr/0042-destructive-confirmation-opt-out.md) |
| ADR-0043 | The live chart title is an owned textbox shape, not a reserved-row cell | Accepted | 2026-10-17 | [`adr/0043-live-title-textbox.md`](adr/0043-live-title-textbox.md) |
| ADR-0044 | A named style carries a stroke dash pattern, in a new appended column | Accepted | 2026-10-19 | [`adr/0044-style-stroke-dash.md`](adr/0044-style-stroke-dash.md) |
| ADR-0045 | The chart chrome colours are persisted, editable settings | Accepted | 2026-10-19 | [`adr/0045-chart-chrome-colours.md`](adr/0045-chart-chrome-colours.md) |
| ADR-0046 | Every chart border is one uniform soft grey (`MajorGridStroke` = `#D9D9D9`) | Accepted | 2026-10-20 | [`adr/0046-uniform-chart-borders.md`](adr/0046-uniform-chart-borders.md) |
| ADR-0047 | The live chart's geometry must stay on the sheet: the title occupies the reserved row, the frame is clamped at its top, and the workbook's metrics drive the row heights | Accepted | 2026-10-21 | [`adr/0047-chart-geometry-stays-on-the-sheet.md`](adr/0047-chart-geometry-stays-on-the-sheet.md) |
| ADR-0048 | The live chart's border encloses the visible data table, drawn on the union of the chart and the measured table | Accepted | 2026-10-21 | [`adr/0048-chart-frame-encloses-the-table.md`](adr/0048-chart-frame-encloses-the-table.md) |
| ADR-0049 | A delineator carries a user-authored named style whose width overrides the chart-wide metric | Accepted | 2026-10-28 | [`adr/0049-delineator-named-style.md`](adr/0049-delineator-named-style.md) |
| ADR-0050 | Empty live chart and coalesced automatic refresh | Accepted | 2026-10-31 | [`adr/0050-empty-live-chart-and-coalesced-refresh.md`](adr/0050-empty-live-chart-and-coalesced-refresh.md) |
| ADR-0051 | Retain the unprotected worksheet; add-in-owned protection is not adopted | Accepted | 2026-11-01 | [`adr/0051-retain-unprotected-worksheet.md`](adr/0051-retain-unprotected-worksheet.md) |
| ADR-0056 | Ribbon first-run surface, Critical Interval attachment, and the append-to-end insert | Accepted | 2026-11-10 | [`adr/0056-ribbon-first-run-critical-interval-and-append-to-end.md`](adr/0056-ribbon-first-run-critical-interval-and-append-to-end.md) |

**ADR-0052 through ADR-0055 have files but no index row.** They are implemented
and cited by other records (verified shape skip, command-scoped shape updates,
the refresh round-trip budget, and add-in-owned sheet protection). The gap is
reported rather than filled in this change: adding four rows would mix an
unrelated index repair into the R5.16 governance commit. The files remain
readable at `docs/adr/0052-*.md` … `0055-*.md`.

**ADR-0041 amends ADR-0040's D2 to make period-label visibility measured-fit.** ADR-0040 let the user choose `MM` or `MMM`, but both forms were gated by the same fixed `MinimumHeaderLabelWidthPt` floor, so `MM` showed on identical intervals to `MMM` — cosmetic, not functional. ADR-0041 replaces that floor **for the period band** with a measured-fit test (show iff the interval is at least as wide as the label measures), so `MM` — measuring shorter — survives on narrower intervals. The band stays uniform: the format is still one verbatim, user-chosen form, only the show/hide decision varies, and this is measured-fit, not ADR-0039's rejected per-interval ladder. The year band keeps the floor; measurement failure falls back to it.

**ADR-0040 supersedes ADR-0039's automatic MMM→MM fit ladder.** The owner rejected the per-interval step-down on review — it makes the band non-uniform (some months at `MMM`, some at `MM`) and is effectively unobservable in live use, because a three-character month at the 8pt measuring seam is narrower than the 18pt `MinimumHeaderLabelWidthPt` suppression floor. The selected `PeriodLabelFormat` is instead surfaced directly as a Plot Time Scale dropdown (Month scale: `MM` or `MMM`) and emitted verbatim, so the format is a uniform, explicit choice. ADR-0039 is retained as history.

**ADR-0034 D3 supersedes ADR-0027's slot-centring.** ADR-0027 records the critical
interval as centred on its own visual slot; ADR-0034 D3 (owner ruling 2026-10-02)
changes that to top-alignment, and `CriticalOverlayBuilder` implements the D3 form.
Where ADR-0027 and ADR-0034 differ on vertical placement, **ADR-0034 wins**. ADR-0034
is itself amended in part by ADR-0035 (its bottom padding row becomes reserved, and
its insert rule becomes conditional on the active cell).

**ADR-0036 and ADR-0037 rows were added 2026-10-03** alongside ADR-0038's. Both
records already existed and were implemented; only the index table had been left
stale.

**ADR-0031 and ADR-0032 are cited by ID but have no file in `docs/adr/`.** They are
referenced from ADR-0030, ADR-0033, ADR-0034 and ADR-0035, and from the roadmap, but
no `adr/0031-*.md` or `adr/0032-*.md` exists. They are therefore absent from this
index: an index row must link to a readable record, and inventing one would fabricate
a decision. This gap is reported, not filled.

**Six index rows said "Proposed" for decisions the code had already shipped (corrected 2026-11-03).** ADR-0043 through ADR-0048 were indexed as `Proposed (owner approved direction …)` / `Proposed (owner reported the symptom …)`, which recorded the governance state when they were authored, not the state of the tree. Each is implemented and exercised: 0043 by `GanttRefreshOrchestrator`'s owned title textbox and `GanttSchemaVersion` 17; 0044 by `GanttStrokeDash` through `ExcelShapeWriter.MapDashStyle`, the style column 14 and the theme dialog; 0045 by `GanttChartChrome`'s settings keys; 0046 by `GanttCatalogues`' `MajorGridStroke`/`MinorGridStroke` both `#D9D9D9`; 0047 by `FrameBandsBuilder`/`SceneBuilder` refusing a band above row 1; 0048 by `ChartFrameGeometry`'s union of chart and measured table. Their statuses are now `Accepted`. Two rows were also **missing entirely** and are added: **ADR-0042** (session-scoped destructive-confirmation opt-out — cited in `CommandBoundary.Dialog`, `ConfirmationSession` and `RemoveCommand`) and **ADR-0049** (the R5.7 delineator named style). Both files existed on disk and were implemented; only the index was stale, which is the same class of drift ADR-0036/0037 corrected here on 2026-10-03.

## ADR template

```markdown
# ADR-<NNNN> — <title>

- **Status**: Proposed | Accepted | Superseded by ADR-<NNNN>
- **Date**: YYYY-MM-DD
- **Context**: the forces at play, the constraint, the question.
- **Decision**: what we decided.
- **Consequences**: trade-offs, what becomes easier, what becomes harder.
- **Alternatives considered**: what we rejected, and why.
```
