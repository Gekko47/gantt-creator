# Chart framing cleanup — splitter rules, plot gridlines, seams and header edges

Status: revision 3. All decisions are taken; no product code has been changed.

Revision 3: O1, O2 and O3 are answered (owner, 2026-11-27). `GridLinePt` and `MinorGridColour` are retired with their consumers; the bands are confirmed as the intended period separation; the horizontal header overlaps are deduplicated. A review claim is corrected: the panel's bottom border and `chart:plot:bottom` abut, they do not overlap.

## Outcome

1. A `Splitter` row renders as a filled `SplitterFill` rectangle spanning the plot lane and **nothing else** — no top or bottom rule.
2. The plot area contains **no interior vertical lines**: no minor period lines and no December/year major lines. The vertical alternating bands are the period separation, and `ShowMinorGrid` / `ShowMajorGrid` are retired from `tblGanttSettings` (schema v30).
3. The plot's own left and right edge lines are **unconditional** — the plot's border is a complete rectangle on all four sides in every composition.
4. In a table-plus-chart composition, the table/plot seam is **one** rule, not two coincident ones.
5. Every header cell draws its **left** edge; only the **last** cell in each band row draws a right edge.
6. No two header primitives share a line: the period row's top edge is the year row's bottom, and the period row's bottom edge is `chart:plot:top`. Each overlap is removed once (D8).
7. The orphaned controls are gone with their consumers, so no theme setting or metric survives with no effect.

## Decisions (all answered)

| Id | Decision | Basis |
|---|---|---|
| D1 | Remove every interior plot line, including December boundaries; retire `ShowMinorGrid` and `ShowMajorGrid`; schema v30 | owner |
| D2 | Header vertical edges: left per cell, right only on the last cell | owner |
| D3 | Splitter band: fill only, no border rules | owner |
| D4 | Plot left/right edges: emitted unconditionally | owner |
| D5 | No live inset for header cell edges | owner, after review |
| D6 | **Retire the orphaned controls in v30:** `GridLinePt` metric, `MinorGridColour` setting and `GanttChartChromeRole.MinorGrid`, and the `FrameBandsTheme.MinorGrid` style. The export panel's border width moves to `MajorBoundaryPt`, superseding ADR-0068 D5. The `MinorGridStroke` colour token stays: [`GanttTableAppearance.BorderColour`](src/GanttCreator.Core/GanttTableAppearance.cs:208) still reads it for the live table and panel borders. | owner approved the recommendation |
| D7 | **Bands are the intended period separation.** No change to `AlternateBandFill` or any other token value. | owner |
| D8 | **Remove one line from every horizontal header overlap:** the period cells lose their own top and bottom edges — the year row's bottom edge is the period row's top, and `chart:plot:top` is the period row's bottom. Year cells keep top and bottom, plus left each and right on the last. | owner: "remove one of the overlapping lines" |

### Correction recorded from the review

Revision 2's review claimed the panel's bottom border and `chart:plot:bottom` are coincident. They are not: the panel's horizontal borders span `panelLeft → panelRight`, and `panelRight == PlotBounds.Left` ([`PanelBuilder.cs`](src/GanttCreator.Core/Scene/PanelBuilder.cs:104)), while `chart:plot:bottom` spans `PlotBounds.Left → Right` ([`FrameBandsBuilder.cs`](src/GanttCreator.Core/Scene/FrameBandsBuilder.cs:583)). The two rules **abut** end to end and form one continuous boundary; there is nothing to deduplicate there. The only horizontal duplicates are inside the header (D8).

## Work item and contract documents (do this first)

- New work item `docs/work-items/R11.2-chart-framing-and-gridline-cleanup.md`. R11.1 is already allocated to the core-refactor remediation, so the identifier moved to R11.2 (the scope-change note and roadmap row are in place).
- New ADRs (the `00xx-` placeholders below shipped as 0074–0076):
  - `docs/adr/0074-splitter-band-is-a-fill-only.md` — supersedes ADR-0022 D3/D4; amends guide §10 and ADR-0027's `SceneLine` list.
  - `docs/adr/0075-retire-the-plot-grid.md` — schema v30 (both tables, see below), `ShowMinorGrid`/`ShowMajorGrid`/`GridLinePt`/`MinorGridColour` retirement, the bands as the period separation, and the amendments to guide §2/§5/§6/§7/§8/§24, ADR-0037 D1, ADR-0038, ADR-0046, ADR-0061 D8, ADR-0068 D5.
  - `docs/adr/0076-plot-edges-and-header-cell-edges.md` — unconditional edges, the single vertical seam, header left/last-right, the D8 horizontal rule, and the statement that the panel/plot bottom rules abut by design.
- Update `docs/DECISIONS.md`, `docs/STATUS.md`, and ADR-0052 D11's member list (it names the splitter's two rules, [`0052-verified-shape-skip.md`](docs/adr/0052-verified-shape-skip.md:160)).
- After any `docs/0N-*.md` change run `pwsh -NoProfile -File scripts/sync-cline-skills.ps1` and commit `.cline/skills/` with it, then hand-refresh the `.roo/skills/07-entity-guide` mirror.

## Workstream 1 — the splitter band loses its two rules

Code:

- [`SplitterBuilder.cs`](src/GanttCreator.Core/Scene/SplitterBuilder.cs:21): delete `BorderStyle` and `BorderWidthPt` from `SplitterRequest`. Their uses are the null guard ([`:131`](src/GanttCreator.Core/Scene/SplitterBuilder.cs:131)), the geometry guard ([`:153`](src/GanttCreator.Core/Scene/SplitterBuilder.cs:153)), the synthesised `borderStyle` ([`:195`](src/GanttCreator.Core/Scene/SplitterBuilder.cs:195)) and the two `SceneLine` entries ([`:208`](src/GanttCreator.Core/Scene/SplitterBuilder.cs:208)). Dead inputs would make their validators unreachable, which `AGENTS.md` treats as a defect.
- Keep `PanelLeftPt` (the `DataPanelLeft`/`Both` label field) and every refusal member; all remain reachable.
- Fix the XML docs that describe the removed lines: the `<remarks>` ([`:78`](src/GanttCreator.Core/Scene/SplitterBuilder.cs:78)) and `SplitterResult.Primitives` ([`:35`](src/GanttCreator.Core/Scene/SplitterBuilder.cs:35)).
- [`SceneBuilder.cs`](src/GanttCreator.Core/Scene/SceneBuilder.cs:1790): drop the `SceneStyle borderStyle` parameter, the `request.MajorBoundaryPt` argument ([`:1844`](src/GanttCreator.Core/Scene/SceneBuilder.cs:1844)) and the comment that justifies it ([`:1817`](src/GanttCreator.Core/Scene/SceneBuilder.cs:1817)).
- [`SceneBuilder.cs`](src/GanttCreator.Core/Scene/SceneBuilder.cs:1114): the call-site comment still says "§10's band spans the data panel and the plot", false since ADR-0068 D1. Fix it and drop the `frameTheme.MajorGrid` argument ([`:1124`](src/GanttCreator.Core/Scene/SceneBuilder.cs:1124)).

Tests:

- [`SplitterBuilderTests.cs`](tests/GanttCreator.Core.Tests/Scene/SplitterBuilderTests.cs:64): rewrite the layer test — the band is the only `SceneRect` at `ZLayer.Section`, and the list contains **no** `SceneLine`.
- Delete the two border tests ([`:80`](tests/GanttCreator.Core.Tests/Scene/SplitterBuilderTests.cs:80), [`:106`](tests/GanttCreator.Core.Tests/Scene/SplitterBuilderTests.cs:106)) and the `_borderStyle` fixture / `borderWidthPt:` argument ([`:15`](tests/GanttCreator.Core.Tests/Scene/SplitterBuilderTests.cs:15), [`:40`](tests/GanttCreator.Core.Tests/Scene/SplitterBuilderTests.cs:40)). Delete, do not weaken.
- Refusal test ([`:206`](tests/GanttCreator.Core.Tests/Scene/SplitterBuilderTests.cs:206)): drop the `BorderStyle` null and `borderWidthPt` cases; keep NaN `PanelLeftPt`, zero-height lane and label-position cases.
- [`SceneBuilderTests.cs`](tests/GanttCreator.Core.Tests/Scene/SceneBuilderTests.cs:1325): one splitter primitive, not three.

Docs: guide §10 Geometry/Style ([`docs/07-GANTT-ENTITY-GUIDE.md`](docs/07-GANTT-ENTITY-GUIDE.md:534) — "`SplitterFill`, fill only; no border, no outline"); ADR-0022 D3/D4 marked superseded in place; ADR-0027's `SceneLine` note ([`docs/adr/0027-critical-interval-filled-rectangle.md`](docs/adr/0027-critical-interval-filled-rectangle.md:113)) loses both grid lines and splitter bands.

## Workstream 2 — remove the plot grid and every control that fed it

Code:

- [`FrameBandsBuilder.cs`](src/GanttCreator.Core/Scene/FrameBandsBuilder.cs:488) `AddGrid`: delete the `lines` dictionary, the `ShowMajorGrid` gate, the period loop, `AddGridLine` ([`:662`](src/GanttCreator.Core/Scene/FrameBandsBuilder.cs:662)), `PeriodKey` ([`:686`](src/GanttCreator.Core/Scene/FrameBandsBuilder.cs:686)) and the ordering loop. **Also drop the now-unused `sequence` parameter and its argument at the call site ([`:286`](src/GanttCreator.Core/Scene/FrameBandsBuilder.cs:286)).**
- Emit, unconditionally: `chart:grid:edge:left` at `PlotBounds.Left + halfStroke`, `chart:grid:edge:right` at `PlotBounds.Right - halfStroke`, `ZLayer.Frame`, `Theme.MajorGrid`, spanning `plotSpan` (keeping the live inset at [`:533`](src/GanttCreator.Core/Scene/FrameBandsBuilder.cs:533) and the header-lift/anchor-row span). `chart:plot:top` and `chart:plot:bottom` unchanged.
- **Dead carriers — remove in the same edit:**
  - `GridLinePt` on [`SceneBuilder`](src/GanttCreator.Core/Scene/SceneBuilder.cs:231) and [`FrameBandsRequest`](src/GanttCreator.Core/Scene/ChartFrameGeometry.cs:143); its only remaining use is the positivity check at [`FrameBandsBuilder.cs`](src/GanttCreator.Core/Scene/FrameBandsBuilder.cs:783).
  - The factory assignment ([`ExcelSceneBuildRequestFactory.cs`](src/GanttCreator.Office/ExcelSceneBuildRequestFactory.cs:687)) and `_gridLineToken` ([`:189`](src/GanttCreator.Office/ExcelSceneBuildRequestFactory.cs:189)).
  - `FrameBandsTheme.MinorGrid` ([`ChartFrameGeometry.cs`](src/GanttCreator.Core/Scene/ChartFrameGeometry.cs:35)) and its `ValidGeometry` member check ([`FrameBandsBuilder.cs`](src/GanttCreator.Core/Scene/FrameBandsBuilder.cs:805)).
- **Positional-argument hazard:** `FrameBandsRequest` is positional with two adjacent `double` widths ([`ChartFrameGeometry.cs`](src/GanttCreator.Core/Scene/ChartFrameGeometry.cs:142)). Remove `GridLinePt`, `ShowMinorGrid` and `ShowMajorGrid` in one edit, then build; a partial edit can rebind `MajorBoundaryPt` silently.
- Catalogue ([`GanttCatalogues.cs`](src/GanttCreator.Core/GanttCatalogues.cs:344)): delete `ShowMinorGrid`, `ShowMajorGrid` ([`:344`](src/GanttCreator.Core/GanttCatalogues.cs:344)), `MinorGridColour` ([`:385`](src/GanttCreator.Core/GanttCatalogues.cs:385)) and the `GridLinePt` metric ([`:257`](src/GanttCreator.Core/GanttCatalogues.cs:257)).
- [`ExcelConfigCatalogueReader.cs`](src/GanttCreator.Office/ExcelConfigCatalogueReader.cs:991): delete the retired names from the boolean-setting classifier.
- Chrome role (D6): remove `GanttChartChromeRole.MinorGrid` and its key/label/default/order entries ([`GanttChartChrome.cs`](src/GanttCreator.Core/GanttChartChrome.cs:70), [`:105`](src/GanttCreator.Core/GanttChartChrome.cs:162)); remove the theme dialog's "Minor grid" colour row.
- Theme dialog metrics (D6): remove `GridLinePt` from `ThemeEditRequest.EditableMetrics` ([`ThemeEditRequest.cs`](src/GanttCreator.Core/ThemeEditRequest.cs:193)), from the dialog's metric group ([`ThemeSettingsForm.cs`](src/GanttCreator.AddIn/ThemeSettingsForm.cs:1408) — the "Grid" group becomes "Borders" with `MajorBoundaryPt` alone) and from `DescribeMetric` ([`:1593`](src/GanttCreator.AddIn/ThemeSettingsForm.cs:1593)).
- Panel border width (D6): [`ExportPanelTheme.cs`](src/GanttCreator.Office/ExportPanelTheme.cs:62) reads `MajorBoundaryPt` instead of `GridLinePt`; update its remarks, which currently explain the opposite choice (ADR-0068 D5 superseded).
- [`GanttSchemaVersion.cs`](src/GanttCreator.Core/GanttSchemaVersion.cs:298): `CurrentSchemaVersion = 30`, with a note in the style of the v9 retirement ([`:117`](src/GanttCreator.Core/GanttSchemaVersion.cs:117)) naming **both** tables: `tblGanttSettings` loses three rows (`ShowMinorGrid`, `ShowMajorGrid`, `MinorGridColour`); `tblGanttMetrics` loses `GridLinePt`. `AdditiveSchemaUpgrade` inspects settings keys only ([`AdditiveSchemaUpgrade.cs`](src/GanttCreator.Core/ConfigIntegrity/AdditiveSchemaUpgrade.cs:109)), so settings removal reports confirmation-required `SchemaVersionOlder` and the metrics removal reports `CatalogueHashMismatch` (auto-repair). Say both.

Tests:

- [`FrameBandsBuilderTests.cs`](tests/GanttCreator.Core.Tests/Scene/FrameBandsBuilderTests.cs:41), [`:95`](tests/GanttCreator.Core.Tests/Scene/FrameBandsBuilderTests.cs:95), [`:136`](tests/GanttCreator.Core.Tests/Scene/FrameBandsBuilderTests.cs:136): keep the shared-span assertions, retarget them at the bands and edge lines, and fix the comments claiming interior grid lines share the span.
- Rewrite `Chart_owner_snapshot_round_trips_and_deduplicates_grid_boundaries` ([`:863`](tests/GanttCreator.Core.Tests/Scene/FrameBandsBuilderTests.cs:863)): keep the round-trip; assert "exactly two `chart:grid:` primitives, both edges".
- Delete `A_width_change_moves_the_grid_lines_without_renaming_them` and its helpers ([`:963`](tests/GanttCreator.Core.Tests/Scene/FrameBandsBuilderTests.cs:963)); record the retirement in `docs/STATUS.md` ([`STATUS`](docs/STATUS.md:152) becomes historical).
- New `[Fact]`: the plot's rectangle is always closed — one edge line each side, one `chart:plot:top`, one `chart:plot:bottom`, no interior `chart:grid:{date}`, in both coordinate modes.
- Remove `new SceneStyle("MinorGrid", …)` from the six theme fixtures: [`FrameBandsBuilderTests.cs`](tests/GanttCreator.Core.Tests/Scene/FrameBandsBuilderTests.cs:11), [`ReferenceSceneBuilder.cs`](tests/GanttCreator.Core.Tests/Scene/ReferenceSceneBuilder.cs:45), [`RepresentativeAcceptanceFixtureTests.cs`](tests/GanttCreator.Core.Tests/Scene/RepresentativeAcceptanceFixtureTests.cs:94), [`SceneBuilderTests.cs`](tests/GanttCreator.Core.Tests/Scene/SceneBuilderTests.cs:122), [`SceneProjectionTests.cs`](tests/GanttCreator.Core.Tests/Scene/SceneProjectionTests.cs:101), [`UserStyleBarHeightTests.cs`](tests/GanttCreator.Core.Tests/Scene/UserStyleBarHeightTests.cs:84).
- [`ExcelSceneBuildRequestFactoryTests.cs`](tests/GanttCreator.Office.ContractTests/ExcelSceneBuildRequestFactoryTests.cs:1310): the four `theme.MinorGrid` assertions ([`:1337`](tests/GanttCreator.Office.ContractTests/ExcelSceneBuildRequestFactoryTests.cs:1349), [`:1389`](tests/GanttCreator.Office.ContractTests/ExcelSceneBuildRequestFactoryTests.cs:1389)) are the only tests that a workbook metric reaches the theme — move them onto `theme.MajorGrid` / the major-width metric. The role-to-role "one uniform grey" comparison ([`:1349`](tests/GanttCreator.Office.ContractTests/ExcelSceneBuildRequestFactoryTests.cs:1349)) has no second role left: replace it with a token-to-token comparison of `MinorGridStroke` and `MajorGridStroke` so the ADR-0046 property is still asserted. Remove the retired names from the settings-key list ([`:1569`](tests/GanttCreator.Office.ContractTests/ExcelSceneBuildRequestFactoryTests.cs:1569)).
- [`ReferenceFixtureDocument.cs`](tests/GanttCreator.Core.Tests/Scene/ReferenceFixtureDocument.cs:42): delete both flags — the fixture JSON never carried them.
- [`GanttConfigCatalogueTests.cs`](tests/GanttCreator.Core.Tests/GanttConfigCatalogueTests.cs:1012): move `PinnedFirstReleaseHash`, the settings count pin and its comment; drop the retired names from the boolean/colour lists ([`:392`](tests/GanttCreator.Core.Tests/GanttConfigCatalogueTests.cs:392), [`:401`](tests/GanttCreator.Core.Tests/GanttConfigCatalogueTests.cs:401)).
- [`GanttSchemaVersionTests.cs`](tests/GanttCreator.Core.Tests/GanttSchemaVersionTests.cs:105): `Assert.Equal(30, …)` plus the v30 paragraph.
- New `[Fact]`: a stored settings key list containing the retired keys makes `AdditiveSchemaUpgrade.TryPlan` return `ExtraKey` — the positive test the retirement relies on.
- Optional but recommended: extend the inert-metric list at [`GanttStyleFamiliesTests.cs`](tests/GanttCreator.Core.Tests/GanttStyleFamiliesTests.cs:159) so a re-added `GridLinePt` cannot silently sit outside `EditableMetrics`. Nothing in CI catches an unconsumed metric today ([`MetricTokenAuthorityTests.cs`](tests/GanttCreator.Architecture.Tests/MetricTokenAuthorityTests.cs:149) checks only that read keys exist).

Docs:

- Guide §8 ([`docs/07-GANTT-ENTITY-GUIDE.md`](docs/07-GANTT-ENTITY-GUIDE.md:482)): rewrite Geometry/Layer/Validation to the plot's rectangle only; delete the "no duplicate coincident minor/major lines" and "year line wins at shared X" obligations, which have no subject.
- Guide §5/§6 ([`:442`](docs/07-GANTT-ENTITY-GUIDE.md:442), [`:444`](docs/07-GANTT-ENTITY-GUIDE.md:444)): "period boundaries continue into the plot as grid lines" is false in two places.
- Guide §7 ([`:460`](docs/07-GANTT-ENTITY-GUIDE.md:460)): the shared vertical span is the bands and the delineators.
- Guide §24 ([`:763`](docs/07-GANTT-ENTITY-GUIDE.md:763)): a delineator shares the bands' span, not the grid lines'.
- Guide §2: remove the two settings rows, the `GridLinePt` metric row ([`:234`](docs/07-GANTT-ENTITY-GUIDE.md:234)), and reword the `MinorGridStroke` description ([`:265`](docs/07-GANTT-ENTITY-GUIDE.md:265)) to the table/panel border colour.
- `FrameBandsRequest.DataPanelBounds` remark ([`ChartFrameGeometry.cs`](src/GanttCreator.Core/Scene/ChartFrameGeometry.cs:126)): drop grid lines from the list of shapes that keep the plot's rectangle.
- Builder comments explaining deleted behaviour: [`:547`](src/GanttCreator.Core/Scene/FrameBandsBuilder.cs:547), [`:565`](src/GanttCreator.Core/Scene/FrameBandsBuilder.cs:565), [`:589`](src/GanttCreator.Core/Scene/FrameBandsBuilder.cs:589), [`:633`](src/GanttCreator.Core/Scene/FrameBandsBuilder.cs:633).
- ADR-0046 ([`docs/adr/0046-uniform-chart-borders.md`](docs/adr/0046-uniform-chart-borders.md:27)): the year-boundary-vs-grid weight hierarchy has **no line left to express**.
- ADR-0037 D1 / ADR-0038 D2 / guide §7: the plot-spanning set is bands, edge lines and delineators.

## Workstream 3 — one rule at the table/plot seam (vertical)

Code:

- [`PanelBuilder.cs`](src/GanttCreator.Core/Scene/PanelBuilder.cs:362) `AddBorders`: stop emitting the rightmost vertical border, and change the method remark ([`:357`](src/GanttCreator.Core/Scene/PanelBuilder.cs:357)) to "every shared edge once; the outer right edge is owned by the plot", so the rule is stated where the old one lives.
- Rejected alternative: keep the panel's border and drop the plot's left edge when a panel exists — that makes one side of the plot's rectangle depend on another entity.
- No horizontal change here: the panel's bottom border abuts `chart:plot:bottom` and the two form one continuous rule (see "Correction recorded from the review").

Tests:

- [`PanelBuilderTests.cs`](tests/GanttCreator.Core.Tests/Scene/PanelBuilderTests.cs:249): N vertical edges for N columns, with the test comment and its doc remark updated to name the plot edge as owner.
- New composed-scene `[Fact]` in `SceneBuilderTests`: on a panel-bearing build exactly one vertical `SceneLine` has `From.X == PlotBounds.Left`, and no `panel-border-v:` line shares that x.

Docs: guide §3/§4 panel borders and §8 seam ownership; the ADR records that the vertical seam is single and the horizontal seam abuts by design.

## Workstream 4 — header vertical edges: left per cell, right only on the last

Code:

- [`FrameBandsBuilder.cs`](src/GanttCreator.Core/Scene/FrameBandsBuilder.cs:441) `AddHeaderBorder`: split the four sides into explicit flags. Years pass left, top, bottom and `last`; periods pass left and `last`. Callers compute `last` from `index == Count - 1` ([`:373`](src/GanttCreator.Core/Scene/FrameBandsBuilder.cs:373), [`:408`](src/GanttCreator.Core/Scene/FrameBandsBuilder.cs:408)).
- Result: N cells produce N+1 vertical lines and no coincident pair.

Tests:

- New `[Fact]`: for N periods, exactly N+1 period vertical lines exist, their x values are distinct, and the set equals the N left edges plus one right edge. The same for N years.
- [`:1090`](tests/GanttCreator.Core.Tests/Scene/FrameBandsBuilderTests.cs:1090) ("each cell is boxed") is narrowed to "every cell draws a left edge at the major weight".

Docs: guide §6 ([`docs/07-GANTT-ENTITY-GUIDE.md`](docs/07-GANTT-ENTITY-GUIDE.md:444)) — left per cell, right on the last — and §8's no-duplicate obligation becomes true for vertical header edges.

## Workstream 5 — header horizontal edges: one line per seam (D8)

Code:

- In the same `AddHeaderBorder` rewrite: **period cells emit no top and no bottom edge.** Their top is the year row's bottom edge ([`FrameBandsBuilder.cs`](src/GanttCreator.Core/Scene/FrameBandsBuilder.cs:451) today draws both); their bottom is `chart:plot:top` ([`FrameBandsBuilder.cs`](src/GanttCreator.Core/Scene/FrameBandsBuilder.cs:594)).
- Year cells keep their top and bottom: the year bottom is the period row's top. The year top is the chart's header top and has no counterpart.
- No change to `chart:plot:top` or `chart:plot:bottom`, so the plot's rectangle stays closed (D4) and the live inset behaviour is untouched (D5).
- Update `AddHeaderBorder`'s remark: "Draws the sides a header band owns; a shared edge is drawn by exactly one band."

Tests:

- New `[Fact]`: no two horizontal header lines share a y — the y set of header top/bottom lines equals the distinct seam set (year top, year/period seam, `chart:plot:top`).
- New `[Fact]`: the period row emits no `:border:top` and no `:border:bottom`, and `chart:plot:top` sits on the period band's bottom edge in the export composition.
- Keep the existing header assertion ([`:1090`](tests/GanttCreator.Core.Tests/Scene/FrameBandsBuilderTests.cs:1090)) after narrowing it as in Workstream 4.

Docs: guide §6 must say which band owns which horizontal edge; §8's no-duplicate obligation covers horizontal header edges too.

## Dropped from this plan

The live half-stroke inset for header cell edges. The live header row **is** the table's header row, so the worksheet's own cell borders already draw at every period boundary; insetting the drawn shape would move it 0.125pt at catalogue defaults and 0.5pt at the test theme's 1pt, creating two near-parallel hairlines where they coincide today. The frame and plot-edge insets were justified by a measured half-stroke loss ([`FrameBandsBuilder.cs`](src/GanttCreator.Core/Scene/FrameBandsBuilder.cs:520)); the header boxes have no such measurement. Header edges stay coincident with the cell borders, recorded as deliberate in the ADR.

## Commit sequence

1. `docs: add R11.2 work item and the three ADRs`.
2. `feat(scene)!: the splitter band is a fill only` — Workstream 1.
3. `feat(scene)!: retire the plot grid and its settings` — Workstream 2, including the dead carriers and schema v30.
4. `fix(scene): the plot edge owns the table-to-chart seam` — Workstream 3.
5. `fix(scene): header cells own their left edge and the last cell its right` — Workstream 4.
6. `fix(scene): one horizontal rule per header seam` — Workstream 5.
7. `test(golden): regenerate the reference scene for the framing cleanup` — dedicated human-reviewed commit with the primitive delta printed (D4 policy; precedent [`docs/STATUS.md`](docs/STATUS.md:1406)).
8. `docs: update entity guide, STATUS and skill mirrors` — plus `sync-cline-skills.ps1` output if not folded earlier.

Each code commit is green apart from the expected scene-golden failure, presented with its delta and regenerated in commit 7 — the repository's documented flow, not a silently red gate.

## Gate matrix

| Gate | Why |
|---|---|
| `dotnet test tests/GanttCreator.Core.Tests` | splitter, frame/bands, panel, scene, catalogue, schema, `AdditiveSchemaUpgrade` |
| `dotnet test tests/GanttCreator.Office.ContractTests` | request factory, theme mapping, renderer translation, host-owned classification |
| `pwsh -NoProfile -File scripts/verify-quick.ps1` | staged gate per commit |
| `pwsh -NoProfile -File scripts/verify.ps1` | branch-final, clean tree |
| `pwsh -NoProfile -File scripts/verify-office.ps1` | live host: splitter band has no rules; no interior lines; four-sided plot; single vertical seam; header left edges, last right edge, no doubled horizontal header line; header edges still land on the table's cell borders; panel bottom and plot bottom read as one rule; bands unchanged |
| `pwsh -NoProfile -File scripts/check-cline-skills.ps1` | guide edits synced into `.cline/skills` |
| `scripts/check-md-links.ps1`, `scripts/check-status.ps1` | doc gates |
| Repair path | open a v29 workbook on the v30 add-in: `SchemaVersionOlder`, confirmation-required Repair, settings rewritten without the three keys, metrics without `GridLinePt`, version and hash stamped |

## Risks

- **Repair becomes mandatory for every existing workbook** (v29→v30 is a removal, so no additive upgrade). Same class as the v9 `MaximumExternalLabelWidthPt` retirement; state it in `STATUS.md` and the ADR.
- **No guard detects an unconsumed metric or chrome colour.** Nothing in CI fails for one ([`MetricTokenAuthorityTests.cs`](tests/GanttCreator.Architecture.Tests/MetricTokenAuthorityTests.cs:149)). Completeness rests on this plan plus the recommended inert-list assertion.
- **Seam weight** (Workstream 3): the surviving vertical rule is the plot edge (`MajorGridStroke`, `MajorBoundaryPt`) rather than the panel border (`MinorGridStroke`, `GridLinePt`). Identical at catalogue defaults; a workbook that raised one metric and not the other will see the seam follow the major weight — consistent with ADR-0060.
- **Panel border width now follows the major boundary** (D6), so a user who raised "Major boundary width" thickens the copied panel's grid. Intended, and it removes the second authority ADR-0068 D5 created.
- **Export shape count** drops (2 splitter lines, N−1 interior grid lines, 1 panel seam line, ~3 header lines per cell, the period row's two horizontal edges); tests pinning total primitive counts move with it.
- **Bands carry the period structure alone** (D7). On the live host, check that #F2F2F2 against #FFFFFF still reads at the user's zoom; nothing in code changes, but the owner's judgement is the gate.

## Out of scope

- The live inset for header cell edges (Dropped above).
- Any token value change (`AlternateBandFill`, `SplitterFill`, `MinorGridStroke`, `MajorGridStroke`).
- Renaming `MinorGridStroke` and its remaining role as the table/panel border colour.
- Horizontal lane separators, which guide §8 mentions but which have no setting and no primitive.
