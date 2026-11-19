using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// R6.1: the export-selection and bounds model. Both selections are pinned against the
/// scene builder's own chart bounds, so the model and the composition R6.3 will stage
/// cannot drift apart.
/// </summary>
/// <remarks>
/// The fixture geometry is the same as <see cref="SceneBuilderTests.PanelSceneRequest"/>'s,
/// with two deliberate differences: the composition padding is the export rule
/// (<see cref="ChartPaddingPt.Uniform"/>) and the profile is never fixed here, so a test
/// can prove the model pins it. Every refusal in
/// <see cref="ExportBoundsRefusal"/> has a positive test that constructs the bad input.
/// </remarks>
public sealed class ExportBoundsResolverTests
{
    /// <summary>The export composition's uniform outer padding.</summary>
    private const double Padding = 4;

    /// <summary>The measured panel header height.</summary>
    private const double HeaderHeight = 10;

    /// <summary>The default measured body-row height.</summary>
    private const double BodyRowHeight = 10;

    /// <summary>The fixture title band height (also a <see cref="SceneBuildRequest"/> input).</summary>
    private const double TitleBandHeight = 14;

    /// <summary>The fixture year band height.</summary>
    private const double YearBandHeight = 16;

    /// <summary>The fixture period band height.</summary>
    private const double PeriodBandHeight = 20;

    private static readonly IReadOnlyList<PanelColumn> _columns =
    [
        new PanelColumn("Id", 80),
        new PanelColumn("Type", 120),
        new PanelColumn("Description", 180),
        new PanelColumn("Start", 70),
        new PanelColumn("Finish", 70),
    ];

    private static readonly PanelTheme _panelTheme = new(
        new SceneStyle("DataPanelFill", fillColour: ColourHex.Parse("#F2F2F2")),
        new SceneStyle("DefaultText", fillColour: ColourHex.Parse("#000000")),
        new SceneStyle("HeaderFill", fillColour: ColourHex.Parse("#D9D9D9")),
        new SceneStyle("HeaderFontSizePt", fillColour: ColourHex.Parse("#000000"), bold: true),
        new SceneStyle("Border", strokeColour: ColourHex.Parse("#7F7F7F"), outlineWidthPt: 0.5));

    /// <summary>
    /// The composition request the model is asked to resolve. The plot starts at the
    /// panel's width (520) so the panel sits at x = 0, which is what makes the panel's
    /// contribution to the union unambiguous.
    /// </summary>
    private static SceneBuildRequest Request(
        IReadOnlyList<GanttEvent>? events = null,
        IReadOnlyList<double>? rowHeights = null,
        IReadOnlyList<PanelColumn>? columns = null,
        string? title = ReferenceSceneFixture.Title)
    {
        GanttValidationOutcome validated = ReferenceSceneFixture.LoadValidated();
        IReadOnlyList<GanttEvent> eventList = events ?? [.. validated.Events];

        // The fixture file numbers its rows by WORKSHEET row (2 upwards, row 1 being the
        // header); a `GanttEvent.RowNumber` is the reader's ONE-BASED BODY index, which is
        // what `PanelRowProjection` orders by and what `LaneRowAnchorResolver` indexes the
        // measured grid with (`RowNumber - 1`). Re-basing here is the fixture telling the
        // truth in the contract's own terms, and it is what makes an anchored request
        // resolvable at all (2026-11-19).
        int firstRowNumber = eventList.Count == 0 ? 1 : eventList.Min(atEvent => atEvent.RowNumber);
        if (firstRowNumber != 1)
        {
            eventList = [.. eventList.Select(atEvent => atEvent with { RowNumber = atEvent.RowNumber - (firstRowNumber - 1) })];
        }

        IReadOnlyList<PanelColumn> gridColumns = columns ?? _columns;

        PanelCellGridCreationOutcome grid = PanelCellGrid.TryCreate(
            gridColumns,
            rowHeights ?? [.. Enumerable.Repeat(BodyRowHeight, eventList.Count)],
            HeaderHeight,
            [.. gridColumns.Select(column => column.Name)],

            // A table with no body rows is a legal empty export state for the
            // refusal paths below; the grid refuses it unless the caller opts in.
            allowEmptyRows: eventList.Count == 0);
        Assert.True(grid.Succeeded, "Panel grid refused: " + grid.Refusal);

        return new SceneBuildRequest
        {
            Events = eventList,
            Registry = ReferenceSceneBuilder.StyleRegistry,
            Grid = grid.Grid,

            // Owner ruling 2026-11-19: the export draws the data panel, so its lanes are
            // anchored to the same grid the panel is laid out from.
            AnchorLanesToRows = true,
            PlotBounds = new RectD(520, 110, 600, 290),
            Metrics = new FakeTextMetrics(static _ => 4.0, 10.0),
            LaneMetrics = new LaneLayoutMetrics(18, 3, 3, 2, 18, 9),
            FrameTheme = ReferenceSceneBuilder.FrameTheme,
            PlotStart = ReferenceSceneFixture.PlotStart,
            PlotFinish = ReferenceSceneFixture.PlotFinish,
            Scale = GanttTimeScale.Month,
            PeriodLabelFormat = GanttPeriodLabelFormat.MMM,
            DateFormat = GanttDateDisplayFormat.DdMMyyyy,
            Title = title,
            ChartPadding = ChartPaddingPt.Uniform(Padding),
            GridLinePt = 0.5,
            MajorBoundaryPt = 1,
            MilestoneSizePt = 8,
            TitleBandHeightPt = TitleBandHeight,
            YearBandHeightPt = YearBandHeight,
            PeriodBandHeightPt = PeriodBandHeight,
            DelineatorLinePt = 1,
            DelineatorStackGapPt = 10,
            LabelGapPt = 2,
            RowHeightPt = 8,
            LabelStyle = new SceneStyle("DefaultText", fillColour: ColourHex.Parse("#000000")),
        };
    }

    [Fact]
    public void Chart_only_bounds_are_the_padded_chart_without_the_panel()
    {
        ExportBoundsOutcome outcome = ExportBoundsResolver.TryResolve(
            Request(),
            _panelTheme,
            ExportSelection.ChartOnly);

        Assert.True(outcome.Succeeded, "Resolution refused: " + outcome.Refusal);
        ExportBounds bounds = outcome.Bounds!;

        // The plot spans x 520..1120 and every band shares that extent, so the
        // content union is the plot's own 600pt width. With no panel composed the
        // union starts at the plot, and the padding is applied once on each side.
        Assert.Equal(0, bounds.Bounds.X);
        Assert.Equal(0, bounds.Bounds.Y);
        Assert.Equal(600 + (2 * Padding), bounds.Bounds.Width, 6);

        // Vertically the union runs from the year band's top (110 - 20 - 16) to the
        // plot's bottom (400): 326pt, padded once to 334. The drawn title shares the year
        // band's row, so it no longer adds a band above it (owner ruling 2026-11-19).
        Assert.Equal(326 + (2 * Padding), bounds.Bounds.Height, 6);

        // The delta is the untranslated top-left corner, and it is carried rather
        // than discarded (entity guide section 1): the padding sits inside the
        // bounds, so a crop that ignored it would lose one margin on two sides.
        Assert.Equal(520 - Padding, bounds.PlacementDelta.X, 6);
        Assert.Equal(74 - Padding, bounds.PlacementDelta.Y, 6);
        Assert.Equal(bounds.PlacementDelta.X, bounds.ChartBounds.X, 6);
        Assert.Equal(bounds.PlacementDelta.Y, bounds.ChartBounds.Y, 6);
    }

    /// <summary>
    /// A composition with no data panel draws no title: the title sits over the table, so a
    /// chart-only copy has nothing to title (owner ruling 2026-11-19, mirroring the live rule
    /// that disables the title when no column is visible — ADR-0065 D4).
    /// </summary>
    [Fact]
    public void A_chart_only_composition_draws_no_title_while_table_plus_chart_does()
    {
        ExportBoundsOutcome chartOnly = ExportBoundsResolver.TryResolve(
            Request(),
            _panelTheme,
            ExportSelection.ChartOnly);
        ExportBoundsOutcome withTable = ExportBoundsResolver.TryResolve(
            Request(),
            _panelTheme,
            ExportSelection.TablePlusChart);

        Assert.True(chartOnly.Succeeded, "Resolution refused: " + chartOnly.Refusal);
        Assert.True(withTable.Succeeded, "Resolution refused: " + withTable.Refusal);

        Assert.DoesNotContain(
            chartOnly.Scene!.Primitives,
            primitive => primitive.PrimitiveId is "chart:title-band" or "chart:title-text");
        Assert.Contains(
            withTable.Scene!.Primitives,
            primitive => primitive.PrimitiveId == "chart:title-text");
    }

    [Fact]
    public void The_title_shares_the_year_band_row_and_extends_no_bounds()
    {
        ExportBoundsOutcome withTitle = ExportBoundsResolver.TryResolve(
            Request(),
            _panelTheme,
            ExportSelection.ChartOnly);
        ExportBoundsOutcome withoutTitle = ExportBoundsResolver.TryResolve(
            Request(title: null),
            _panelTheme,
            ExportSelection.ChartOnly);

        Assert.True(withTitle.Succeeded, "Resolution refused: " + withTitle.Refusal);
        Assert.True(withoutTitle.Succeeded, "Resolution refused: " + withoutTitle.Refusal);

        // Owner ruling 2026-11-19: the drawn title occupies the year band's row over the
        // data panel, exactly where the live title sits, so it adds no height of its own.
        // It used to be a band above every band, which pushed the composition one title
        // height taller and shifted the placement delta by the same amount.
        Assert.Equal(withTitle.Bounds!.ChartBounds, withoutTitle.Bounds!.ChartBounds);
        Assert.Equal(withTitle.Bounds.PlacementDelta, withoutTitle.Bounds.PlacementDelta);
        Assert.Equal(326 + (2 * Padding), withoutTitle.Bounds.Bounds.Height, 6);
    }

    [Fact]
    public void Table_plus_chart_adds_exactly_the_measured_panel_width()
    {
        ExportBoundsOutcome chartOnly = ExportBoundsResolver.TryResolve(
            Request(),
            _panelTheme,
            ExportSelection.ChartOnly);
        ExportBoundsOutcome withTable = ExportBoundsResolver.TryResolve(
            Request(),
            _panelTheme,
            ExportSelection.TablePlusChart);

        Assert.True(chartOnly.Succeeded, "Resolution refused: " + chartOnly.Refusal);
        Assert.True(withTable.Succeeded, "Resolution refused: " + withTable.Refusal);

        double panelWidth = _columns.Sum(column => column.WidthPt);

        // The panel contributes its measured width and nothing else: the union's
        // vertical extent is unchanged because the panel sits within the plot's
        // own top-to-bottom extent in this fixture.
        Assert.Equal(chartOnly.Bounds!.Bounds.Width + panelWidth, withTable.Bounds!.Bounds.Width, 6);
        Assert.Equal(chartOnly.Bounds.Bounds.Height, withTable.Bounds.Bounds.Height, 6);

        // The panel extends the union to the left, so the translated delta moves
        // left by exactly its width.
        Assert.Equal(chartOnly.Bounds.PlacementDelta.X - panelWidth, withTable.Bounds.PlacementDelta.X, 6);
    }

    [Fact]
    public void Table_plus_chart_refuses_an_empty_visible_table_with_the_actionable_message()
    {
        ExportBoundsOutcome outcome = ExportBoundsResolver.TryResolve(
            Request(events: []),
            _panelTheme,
            ExportSelection.TablePlusChart);

        Assert.False(outcome.Succeeded);
        Assert.Equal(ExportBoundsRefusal.EmptyVisibleTable, outcome.Refusal);
        Assert.Null(outcome.SceneRefusal);
        Assert.Equal(
            "A table-plus-chart copy needs at least one visible row. Make a row visible, or copy the chart only.",
            ExportBoundsRefusalMessages.EmptyVisibleTable);
    }

    [Fact]
    public void Table_plus_chart_refuses_a_table_whose_rows_are_all_hidden()
    {
        GanttValidationOutcome validated = ReferenceSceneFixture.LoadValidated();
        IReadOnlyList<GanttEvent> hidden =
        [
            .. validated.Events.Select(@event => @event with { Visible = false }),
        ];

        ExportBoundsOutcome outcome = ExportBoundsResolver.TryResolve(
            Request(events: hidden),
            _panelTheme,
            ExportSelection.TablePlusChart);

        Assert.Equal(ExportBoundsRefusal.EmptyVisibleTable, outcome.Refusal);
    }

    [Fact]
    public void Chart_only_does_not_apply_the_table_refusal_to_an_empty_table()
    {
        ExportBoundsOutcome outcome = ExportBoundsResolver.TryResolve(
            Request(events: []),
            null,
            ExportSelection.ChartOnly);

        // The export still refuses to compose an empty chart, but through the scene
        // builder's own empty-events refusal - not the table rule. This pins the D2
        // asymmetry: an empty table is invalid as an export either way, and the
        // export refusal must never be "fixed" to match the live calendar-year
        // fallback (ADR-0050, R5.18j).
        Assert.Equal(ExportBoundsRefusal.SceneBuildRefused, outcome.Refusal);
        Assert.Equal(SceneBuilderRefusal.EmptyEvents, outcome.SceneRefusal);
    }

    [Theory]
    [InlineData(0, 290)]
    [InlineData(600, 0)]
    public void A_plot_with_no_area_is_refused_as_invalid_bounds(double width, double height)
    {
        SceneBuildRequest request = Request() with
        {
            PlotBounds = new RectD(520, 110, width, height),
        };

        ExportBoundsOutcome outcome = ExportBoundsResolver.TryResolve(
            request,
            _panelTheme,
            ExportSelection.TablePlusChart);

        // Raised before any composition, so no scene refusal accompanies it.
        Assert.Equal(ExportBoundsRefusal.InvalidBounds, outcome.Refusal);
        Assert.Null(outcome.SceneRefusal);
    }

    [Fact]
    public void Table_plus_chart_without_a_panel_theme_is_refused_rather_than_exported_chart_only()
    {
        ExportBoundsOutcome outcome = ExportBoundsResolver.TryResolve(
            Request(),
            null,
            ExportSelection.TablePlusChart);

        Assert.Equal(ExportBoundsRefusal.PanelThemeMissing, outcome.Refusal);
    }

    [Fact]
    public void An_undefined_selection_fails_closed()
    {
        ExportBoundsOutcome outcome = ExportBoundsResolver.TryResolve(
            Request(),
            _panelTheme,
            (ExportSelection)99);

        Assert.Equal(ExportBoundsRefusal.InvalidSelection, outcome.Refusal);
    }

    [Fact]
    public void A_null_request_is_refused()
    {
        ExportBoundsOutcome outcome = ExportBoundsResolver.TryResolve(
            null,
            _panelTheme,
            ExportSelection.ChartOnly);

        Assert.Equal(ExportBoundsRefusal.NullRequest, outcome.Refusal);
    }

    [Fact]
    public void Zero_visible_columns_degrade_to_the_chart_only_bounds()
    {
        ExportBoundsOutcome chartOnly = ExportBoundsResolver.TryResolve(
            Request(),
            _panelTheme,
            ExportSelection.ChartOnly);
        ExportBoundsOutcome zeroColumns = ExportBoundsResolver.TryResolve(
            Request(columns: []),
            _panelTheme,
            ExportSelection.TablePlusChart);

        // ADR-0065 D4: a table with every user column hidden is a supported state.
        // The grid sums to zero width, the panel contributes no bounds, and the
        // table-plus-chart export is exactly the chart-only export - never a
        // refusal and never a phantom-width panel.
        Assert.True(
            zeroColumns.Succeeded,
            "Resolution refused: " + zeroColumns.Refusal + "/" + zeroColumns.SceneRefusal);
        Assert.Equal(chartOnly.Bounds!.Bounds, zeroColumns.Bounds!.Bounds);
        Assert.Equal(chartOnly.Bounds.PlacementDelta, zeroColumns.Bounds.PlacementDelta);
    }

    [Fact]
    public void Collapsed_rows_shrink_the_panel_but_never_the_chart()
    {
        GanttValidationOutcome validated = ReferenceSceneFixture.LoadValidated();
        double tallRow = 24;

        // Tall enough that the panel, not the plot, governs the union's bottom, so a
        // collapsed group's zero-height rows are observable in the export bounds.
        int firstRowNumber = validated.Events.Min(atEvent => atEvent.RowNumber);
        double[] expanded = [.. Enumerable.Repeat(tallRow, validated.Events.Count)];
        double[] collapsed = [.. expanded];

        // A collapsed row is a CHILD hidden by its parent's outline, so it draws on its
        // parent's lane and anchoring still resolves. (A hidden row that OWNED a lane would
        // be refused as a degenerate height, and that cannot happen on a sheet.) The
        // fixture's two critical intervals are exactly such children.
        foreach (GanttEvent child in validated.Events.Where(
            atEvent => atEvent.Type == GanttEntityType.CriticalInterval))
        {
            collapsed[child.RowNumber - firstRowNumber] = 0;
        }

        ExportBoundsOutcome expandedTable = ExportBoundsResolver.TryResolve(
            Request(rowHeights: expanded),
            _panelTheme,
            ExportSelection.TablePlusChart);
        ExportBoundsOutcome collapsedTable = ExportBoundsResolver.TryResolve(
            Request(rowHeights: collapsed),
            _panelTheme,
            ExportSelection.TablePlusChart);
        ExportBoundsOutcome expandedChart = ExportBoundsResolver.TryResolve(
            Request(rowHeights: expanded),
            _panelTheme,
            ExportSelection.ChartOnly);
        ExportBoundsOutcome collapsedChart = ExportBoundsResolver.TryResolve(
            Request(rowHeights: collapsed),
            _panelTheme,
            ExportSelection.ChartOnly);

        Assert.True(expandedTable.Succeeded, "Resolution refused: " + expandedTable.Refusal);
        Assert.True(collapsedTable.Succeeded, "Resolution refused: " + collapsedTable.Refusal);
        Assert.True(expandedChart.Succeeded, "Resolution refused: " + expandedChart.Refusal);
        Assert.True(collapsedChart.Succeeded, "Resolution refused: " + collapsedChart.Refusal);

        // The panel shrinks by exactly the collapsed rows' measured height (ADR-0034
        // accepts zero-height rows by design).
        Assert.Equal(
            2 * tallRow,
            expandedTable.Bounds!.Bounds.Height - collapsedTable.Bounds!.Bounds.Height,
            6);

        // The defect this guard closes (2026-11-19): the collapsed rows' panel cells
        // were emitted as zero-height rectangles, and the shape writer refuses a
        // non-positive extent before any host call, so every editable copy failed on
        // the first hidden row. The scene is the only geometry authority, so nothing in
        // the export scene may be a request no renderer can create.
        Assert.All(
            collapsedTable.Scene!.Primitives.OfType<SceneRect>(),
            rect => Assert.True(rect.Bounds.Width > 0 && rect.Bounds.Height > 0, rect.PrimitiveId));
        Assert.All(
            collapsedTable.Scene!.Primitives.OfType<SceneText>(),
            text => Assert.True(text.TextBounds.Width > 0 && text.TextBounds.Height > 0, text.PrimitiveId));

        // ...and the Gantt itself never shrinks: the chart-only bounds do not read
        // row heights at all. The collapsed children still compose on their parent's
        // lane (R4.7B), which the successful resolution proves.
        Assert.Equal(expandedChart.Bounds!.Bounds, collapsedChart.Bounds!.Bounds);
        Assert.Equal(expandedChart.Bounds.PlacementDelta, collapsedChart.Bounds.PlacementDelta);
    }

    [Fact]
    public void The_profile_is_the_export_contract_not_the_callers_choice()
    {
        // A live profile omits the drawn title band (ADR-0030 D6). The model pins
        // EditableExport, so a request that still carried a live profile produces the
        // same bounds as an explicit export request.
        ExportBoundsOutcome liveTemplate = ExportBoundsResolver.TryResolve(
            Request() with { Profile = SceneCompositionProfile.LiveExcel },
            _panelTheme,
            ExportSelection.ChartOnly);
        ExportBoundsOutcome exportTemplate = ExportBoundsResolver.TryResolve(
            Request() with { Profile = SceneCompositionProfile.EditableExport },
            _panelTheme,
            ExportSelection.ChartOnly);

        Assert.True(liveTemplate.Succeeded, "Resolution refused: " + liveTemplate.Refusal);
        Assert.True(exportTemplate.Succeeded, "Resolution refused: " + exportTemplate.Refusal);
        Assert.Equal(exportTemplate.Bounds!.Bounds, liveTemplate.Bounds!.Bounds);
        Assert.Equal(exportTemplate.Bounds.PlacementDelta, liveTemplate.Bounds.PlacementDelta);
    }

    [Fact]
    public void Chart_only_drops_a_panel_the_template_carried()
    {
        // D4: the selection decides whether the panel is supplied. There is no second
        // authority, so a template that already carries a panel is reduced to the
        // chart-only composition rather than being composed as table-plus-chart.
        ExportBoundsOutcome fromPanelTemplate = ExportBoundsResolver.TryResolve(
            Request() with { Panel = _panelTheme },
            _panelTheme,
            ExportSelection.ChartOnly);
        ExportBoundsOutcome fromPlainTemplate = ExportBoundsResolver.TryResolve(
            Request(),
            _panelTheme,
            ExportSelection.ChartOnly);

        Assert.True(fromPanelTemplate.Succeeded, "Resolution refused: " + fromPanelTemplate.Refusal);
        Assert.Equal(fromPlainTemplate.Bounds!.Bounds, fromPanelTemplate.Bounds!.Bounds);
        Assert.Equal(fromPlainTemplate.Bounds.PlacementDelta, fromPanelTemplate.Bounds.PlacementDelta);
    }

    [Fact]
    public void Repeated_resolution_is_deterministic()
    {
        ExportBoundsOutcome first = ExportBoundsResolver.TryResolve(
            Request(),
            _panelTheme,
            ExportSelection.TablePlusChart);
        ExportBoundsOutcome second = ExportBoundsResolver.TryResolve(
            Request(),
            _panelTheme,
            ExportSelection.TablePlusChart);

        Assert.Equal(first.Bounds, second.Bounds);
    }

    [Fact]
    public void The_resolved_scene_is_the_scene_the_bounds_were_read_from()
    {
        // R6.2 amendment: the bounds describe a scene, so the scene is carried on the
        // outcome rather than discarded. The composition consumes this scene instead of
        // building a second one, which is what keeps the bounds and the staged shapes
        // from being two authorities over the same composition.
        ExportBoundsOutcome outcome = ExportBoundsResolver.TryResolve(
            Request(),
            _panelTheme,
            ExportSelection.TablePlusChart);

        Assert.True(outcome.Succeeded, "Resolution refused: " + outcome.Refusal);
        Assert.NotNull(outcome.Scene);

        // The bounds are exactly the carried scene's chart bounds - the same object
        // graph, not a recomputation that could drift.
        Assert.Equal(outcome.Scene!.ChartBounds, outcome.Bounds!.ChartBounds);

        // And the carried scene is byte-for-byte the export scene a direct build
        // produces: same primitive identifiers in the same order, same plot.
        SceneBuildOutcome direct = SceneBuilder.TryBuild(
            Request() with
            {
                Profile = SceneCompositionProfile.EditableExport,
                Panel = _panelTheme,
            });
        Assert.NotNull(direct.Result);
        Assert.Equal(direct.Result!.Scene!.PlotBounds, outcome.Scene.PlotBounds);
        Assert.Equal(
            direct.Result.Scene.Primitives.Select(primitive => primitive.PrimitiveId),
            outcome.Scene.Primitives.Select(primitive => primitive.PrimitiveId));
    }

    [Fact]
    public void An_unanchored_request_is_refused_before_any_composition()
    {
        // Owner ruling 2026-11-19: the export draws the data panel and reproduces the sheet
        // row for row, so its lanes must be anchored to the same grid. Stacked lanes left
        // every lane one row above its panel row (the live defect). The model refuses an
        // unanchored request rather than composing a chart that cannot line up.
        ExportBoundsOutcome outcome = ExportBoundsResolver.TryResolve(
            Request() with { AnchorLanesToRows = false },
            _panelTheme,
            ExportSelection.ChartOnly);

        Assert.False(outcome.Succeeded);
        Assert.Equal(ExportBoundsRefusal.UnanchoredRequest, outcome.Refusal);
        Assert.Null(outcome.Scene);
        Assert.Null(outcome.Bounds);
    }

    [Fact]
    public void A_refusal_carries_no_scene()
    {
        // The scene member is success-only, so a caller cannot stage a partial
        // composition from a refused resolution.
        ExportBoundsOutcome outcome = ExportBoundsResolver.TryResolve(
            Request(events: []),
            _panelTheme,
            ExportSelection.TablePlusChart);

        Assert.False(outcome.Succeeded);
        Assert.Null(outcome.Scene);
    }
}
