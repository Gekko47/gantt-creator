using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// Tests for the R3.12 orchestrator. Every refusal in
/// <see cref="SceneBuilderRefusal"/> has a positive test here that constructs the
/// bad input and asserts the error path fires, because a validator without a
/// positive test is how a refusal silently stops working.
/// </summary>
public sealed class SceneBuilderTests
{
    private static readonly DateOnly _plotStart = new(2024, 1, 1);
    private static readonly DateOnly _plotFinish = new(2024, 1, 31);
    private static readonly ITextMetrics _metrics = new FakeTextMetrics(_ => 8.0, 10.0);
    private static readonly LaneLayoutMetrics _laneMetrics = new(18, 3, 3, 2, 18, 9);

    // A panel at 0..200 and a plot at 200..500 inside a 0..520 chart, so the plot
    // is inside the chart by construction and every default is already coherent.
    // The plot top is 60 because the three header bands total 14 + 16 + 20 = 50pt
    // and are placed above the plot: with a smaller top the period band would be
    // emitted above the chart's own top edge (see the R3.5 finding in the work item).
    private static readonly RectD _panelBounds = new(0, 0, 200, 200);
    private static readonly RectD _plotBounds = new(200, 60, 300, 140);

    private static GanttStyleDefinition Style(string key, double height) =>
        new(
            key,
            new HashSet<GanttLabelPosition>
            {
                GanttLabelPosition.Inside,
                GanttLabelPosition.Left,
                GanttLabelPosition.Right,
                GanttLabelPosition.Auto,
            },
            EntityColourCapability.Fill | EntityColourCapability.Stroke,
            GanttLabelPosition.Inside,
            "#92D050",
            "#404040",
            "#000000",
            GanttHatchPattern.None,
            0,
            0,
            0.5,
            height,
            8);

    private static readonly IReadOnlyList<GanttStyleDefinition> SharedStyles =
    [
        Style("AsPlannedActivity", 8),
        Style("AsBuiltActivity", 8),
        Style("CriticalInterval", 8),
        Style("AsPlannedMilestone", 8),
    ];

    private static readonly GanttStyleRegistry _registry = new(SharedStyles);

    private static PanelCellGrid Grid() =>
        PanelCellGrid.TryCreate(
            [new PanelColumn("Id", 40), new PanelColumn("Description", 160)],
            10,
            ["Id", "Description"]).Grid!;

    private static FrameBandsTheme FrameTheme() => new(
        new SceneStyle("Background"),
        new SceneStyle("AlternateBand"),
        new SceneStyle("MinorGrid"),
        new SceneStyle("MajorGrid"),
        new SceneStyle("YearHeader"),
        new SceneStyle("PeriodHeader"),
        new SceneStyle("Title"));

    private static SceneBuildRequest Request(params GanttEvent[] events) =>
        new()
        {
            Events = events,
            Registry = _registry,
            Grid = Grid(),
            PanelBounds = _panelBounds,
            PlotBounds = _plotBounds,
            Metrics = _metrics,
            LaneMetrics = _laneMetrics,
            FrameTheme = FrameTheme(),
            PlotStart = _plotStart,
            PlotFinish = _plotFinish,
            GridLinePt = 0.5,
            MajorBoundaryPt = 1,
            MilestoneSizePt = 8,
            CriticalLinePt = 1,
            // Every band height must be strictly positive, and the three bands must
            // fit above the plot: 14 + 16 + 20 = 50, leaving 110 of plot height.
            TitleBandHeightPt = 14,
            YearBandHeightPt = 16,
            PeriodBandHeightPt = 20,
            // §24 delineator tokens, positive so the line builder has a width to use.
            // A request that omits them has no line width, which the builder refuses
            // rather than defaulting.
            DelineatorLinePt = 1,
            DelineatorStackGapPt = 10,
            LabelGapPt = 2,
            LabelHeightPt = 8,
            ChartOuterPaddingPt = 0,
            MinimumHeaderLabelWidthPt = 0,
        };

    private static GanttEvent Event(
        int row,
        GanttEntityType type = GanttEntityType.AsPlannedActivity,
        DateOnly? start = null,
        DateOnly? finish = null,
        string? styleKey = "AsPlannedActivity",
        bool visible = true,
        GanttRowId? parentId = null) =>
        new(
            row,
            GanttRowId.New(),
            GanttRowId.New(),
            0,
            type,
            $"Row {row}",
            start ?? new DateOnly(2024, 1, 5),
            finish ?? (type is GanttEntityType.AsPlannedMilestone ? null : new DateOnly(2024, 1, 10)),
            parentId,
            styleKey,
            null,
            null,
            null,
            visible,
            null);


    [Fact]
    public void A_valid_request_builds_a_scene_with_the_frame_and_the_bar()
    {
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(Request(Event(1)));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        // The frame contributes a chart background at Background and the bar sits
        // at ActivityBody, so both layers must be present for a complete scene.
        Assert.Contains(outcome.Result!.Scene.Primitives, p => p.ZLayer == ZLayer.Background);
        Assert.Contains(outcome.Result!.Scene.Primitives, p => p.ZLayer == ZLayer.ActivityBody);
        Assert.Equal(_plotBounds, outcome.Result!.Scene.PlotBounds);
    }

    [Fact]
    public void A_null_request_is_refused()
    {
        Assert.Equal(SceneBuilderRefusal.NullRequest, SceneBuilder.TryBuild(null).Refusal);
    }

    [Fact]
    public void An_empty_event_list_is_refused()
    {
        Assert.Equal(SceneBuilderRefusal.EmptyEvents, SceneBuilder.TryBuild(Request()).Refusal);
    }

    [Fact]
    public void A_null_text_metrics_seam_is_refused()
    {
        Assert.Equal(
            SceneBuilderRefusal.NullMetrics,
            SceneBuilder.TryBuild(Request(Event(1)) with { Metrics = null }).Refusal);
    }

    [Fact]
    public void A_null_measured_grid_is_refused()
    {
        Assert.Equal(
            SceneBuilderRefusal.NullPanelGrid,
            SceneBuilder.TryBuild(Request(Event(1)) with { Grid = null }).Refusal);
    }

    [Fact]
    public void Null_bounds_are_refused()
    {
        Assert.Equal(SceneBuilderRefusal.NullBounds, SceneBuilder.TryBuild(Request(Event(1)) with { PlotBounds = null }).Refusal);
        Assert.Equal(SceneBuilderRefusal.NullBounds, SceneBuilder.TryBuild(Request(Event(1)) with { PanelBounds = null }).Refusal);
    }

    [Fact]
    public void The_declared_chart_bounds_equal_the_emitted_chart_background()
    {
        // The regression this item fixes. Entity guide section 2 defines ChartBounds
        // as the union of the title, panel, headers, and plot, so the frame builder
        // derives it; the scene must then be created with that same value. It was
        // previously created with a caller-supplied chart bounds instead, so the
        // scene could declare bounds the chart:background it contained did not have.
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(Request(Event(1)));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        SceneRect background = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneRect>(),
            rect => rect.PrimitiveId == "chart:background");
        Assert.Equal(background.Bounds, outcome.Result!.Scene.ChartBounds);
    }

    [Fact]
    public void An_awkward_measured_geometry_still_produces_a_self_consistent_scene()
    {
        // The R3.12 fixture had to place the panel at y=20 and the plot at y=110 to
        // keep the header bands inside its declared chart. With chart bounds derived
        // rather than supplied, that fudge is no longer needed: the bands may extend
        // above the plot and the scene stays coherent, because the chart bounds
        // follow the content instead of constraining it.
        SceneBuildRequest request = Request(Event(1)) with
        {
            PanelBounds = new RectD(0, 0, 200, 200),
            PlotBounds = new RectD(200, 20, 300, 140),
        };

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(request);

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        SceneValidationReport report = SceneValidator.Validate(outcome.Result!.Scene);
        Assert.True(report.IsClean, "Findings: " + string.Join("; ", report.Findings));
    }

    [Fact]
    public void A_degenerate_plot_range_is_refused()
    {
        // A finish before the start cannot produce a monotonic scale.
        Assert.Equal(
            SceneBuilderRefusal.InvalidPlotRange,
            SceneBuilder.TryBuild(Request(Event(1)) with { PlotFinish = new DateOnly(2023, 12, 1) }).Refusal);
    }

    [Fact]
    public void Null_lane_metrics_are_refused()
    {
        Assert.Equal(
            SceneBuilderRefusal.InvalidLayoutSettings,
            SceneBuilder.TryBuild(Request(Event(1)) with { LaneMetrics = null }).Refusal);
    }

    [Fact]
    public void Null_frame_theme_is_refused()
    {
        Assert.Equal(
            SceneBuilderRefusal.InvalidLayoutSettings,
            SceneBuilder.TryBuild(Request(Event(1)) with { FrameTheme = null }).Refusal);
    }


    [Fact]
    public void An_unresolvable_style_is_refused_rather_than_defaulted_silently()
    {
        // Custom Activity has no Type default, so a blank style key has nothing to
        // fall back to. The build must refuse instead of rendering an unstyled bar.
        // Note that an *empty registry* is deliberately not the trigger: the
        // resolver falls back to the built-in preset for Types that have one.
        Assert.Equal(
            SceneBuilderRefusal.UnresolvableStyle,
            SceneBuilder.TryBuild(
                Request(Event(1, GanttEntityType.CustomActivity, styleKey: null))).Refusal);
    }

    [Fact]
    public void A_hidden_only_event_list_is_refused_because_nothing_can_render()
    {
        Assert.Equal(
            SceneBuilderRefusal.EmptyEvents,
            SceneBuilder.TryBuild(Request(Event(1, visible: false))).Refusal);
    }

    [Fact]
    public void A_splitter_and_spacer_alone_are_refused_because_neither_has_an_entity_primitive()
    {
        Assert.Equal(
            SceneBuilderRefusal.EmptyEvents,
            SceneBuilder.TryBuild(Request(
                Event(1, GanttEntityType.Splitter, styleKey: null),
                Event(2, GanttEntityType.Spacer, styleKey: null))).Refusal);
    }

    [Fact]
    public void A_milestone_emits_its_diamond_and_reads_only_the_start_date()
    {
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(Event(1, GanttEntityType.AsPlannedMilestone, styleKey: "AsPlannedMilestone")));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        Assert.Contains(outcome.Result!.Scene.Primitives, p => p.ZLayer == ZLayer.Milestone);
    }

    [Fact]
    public void A_critical_interval_clips_to_its_parents_visible_span()
    {
        GanttEvent parent = Event(1);
        GanttEvent child = Event(
            2,
            GanttEntityType.CriticalInterval,
            new DateOnly(2024, 1, 6),
            new DateOnly(2024, 1, 8),
            "CriticalInterval",
            parentId: parent.Id);

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(Request(parent, child));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        // The overlay must be narrower than the plot, proving it was clipped to the
        // parent rather than spanning the whole plot.
        SceneRect overlay = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneRect>(),
            rect => rect.ZLayer == ZLayer.CriticalOverlay);
        Assert.True(
            overlay.Bounds.Width < _plotBounds.Width,
            "A critical overlay must clip to its parent, not span the full plot.");
    }

    [Fact]
    public void A_critical_interval_whose_parent_is_invisible_emits_no_overlay()
    {
        // A hidden parent is not rendered, so no visible span exists to clip
        // against; an overlay must not appear without one.
        GanttEvent parent = Event(1, visible: false);
        GanttEvent child = Event(
            2,
            GanttEntityType.CriticalInterval,
            new DateOnly(2024, 1, 6),
            new DateOnly(2024, 1, 8),
            "CriticalInterval",
            parentId: parent.Id);

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(Request(parent, child));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        Assert.DoesNotContain(
            outcome.Result!.Scene.Primitives,
            primitive => primitive.ZLayer == ZLayer.CriticalOverlay);
    }

    [Fact]
    public void The_build_never_remeasures_the_supplied_grid()
    {
        SceneBuildRequest request = Request(Event(1));
        PanelCellGrid before = request.Grid!;

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(request);

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        // Core measures nothing: the grid the caller supplied must reach the panel
        // builder unchanged, with the same columns and row height.
        Assert.Equal(before.RowHeightPt, request.Grid!.RowHeightPt);
        Assert.Equal(
            before.Columns.Select(column => column.Name),
            request.Grid.Columns.Select(column => column.Name));
    }


    [Fact]
    public void The_panel_header_band_bottom_equals_the_R3_5_derived_period_header_bottom()
    {
        // Section 4 fixes the header bottom to PlotBounds.Y - YearBandHeightPt. The
        // panel builder cannot know the plot bounds, so SceneBuilder supplies it;
        // this test proves the value arrives intact rather than being assumed.
        SceneBuildRequest request = Request(Event(1)) with
        {
            Panel = new PanelTheme(
                new SceneStyle("BodyFill", fillColour: ColourHex.Parse("#FFFFFF")),
                new SceneStyle("BodyText"),
                new SceneStyle("HeaderFill", fillColour: ColourHex.Parse("#D9D9D9")),
                new SceneStyle("HeaderText", bold: true),
                new SceneStyle("Border")),
        };

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(request);

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        double expectedBottom = _plotBounds.Y - request.YearBandHeightPt;

        // Filtered to the header cells specifically. The period band is also a
        // SceneRect whose bottom happens to equal expectedBottom, so an unfiltered
        // Assert.Contains would pass on the frame alone and prove nothing about the
        // panel's header.
        SceneRect[] headers =
        [
            .. outcome.Result!.Scene.Primitives
                .OfType<SceneRect>()
                .Where(rect => rect.PrimitiveId.Contains("header-cell:", StringComparison.Ordinal)),
        ];
        Assert.NotEmpty(headers);
        Assert.All(
            headers,
            header => Assert.Equal(expectedBottom, header.Bounds.Top + header.Bounds.Height, precision: 9));
    }

    [Fact]
    public void A_span_bar_that_cannot_be_placed_warns_instead_of_vanishing_silently()
    {
        // SpanBarBuilder refuses an event that is not a span (no Finish, or a Finish
        // before its Start). The span pass used to `continue` with no warning, so a
        // row whose bar could not be placed was indistinguishable from a row whose
        // bar was clipped away. The other three placement passes already warn.
        //
        // The build still succeeds: a refused bar is a per-entity warning, not a
        // whole-scene refusal, exactly as the milestone and overlay passes behave.
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(Event(1, start: new DateOnly(2024, 1, 10), finish: new DateOnly(2024, 1, 5)))
            with { LabelStyle = new SceneStyle("DefaultText") });

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);

        // Positive control: the reversed range really is refused, so the warning
        // below is not passing on some unrelated row's own warning.
        Assert.DoesNotContain(
            outcome.Result!.Scene.Primitives,
            primitive => primitive.PrimitiveId.EndsWith(":bar", StringComparison.Ordinal));

        SceneWarning warning = Assert.Single(
            outcome.Result!.Scene.Warnings,
            candidate => candidate.Code == "SpanBarRefused");
        Assert.Equal(SceneOwnerKind.Row, warning.OwnerId.Kind);
        Assert.False(string.IsNullOrWhiteSpace(warning.Message));
    }

    [Fact]
    public void An_external_label_is_capped_by_the_configured_maximum_width_not_the_plot_width()
    {
        // The regression the request property fixes. LabelMetrics' maximum external
        // width was the plot width, so the cap moved with the time axis rather than
        // following the approved MaximumExternalLabelWidthPt token. Here the plot is
        // 300pt wide but the configured cap is 36pt, so a long description must be
        // truncated to the cap rather than allowed the whole 300pt gap.
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(
                Event(1, start: new DateOnly(2024, 1, 8), finish: new DateOnly(2024, 1, 12))
                with
                {
                    LabelPosition = GanttLabelPosition.Right,
                    Description = new string('W', 60),
                })
            with
            {
                LabelStyle = new SceneStyle("DefaultText"),
                MaximumExternalLabelWidthPt = 36,
            });

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        Assert.True(_plotBounds.Width > 36, "The plot must be wider than the cap for this test to discriminate.");

        // Filtered to row-owned text: the frame's own period/year labels also end in
        // ":label", so an unfiltered match would pick a chart label instead.
        SceneText description = Assert.Single(
            outcome.Result!.Scene.Primitives
                .OfType<SceneText>(),
            text => text.OwnerId.Kind == SceneOwnerKind.Row
                && text.PrimitiveId.EndsWith(":label", StringComparison.Ordinal));
        Assert.True(
            description.TextBounds.Width <= 36 + 1e-9,
            "An external label must respect the configured cap, not the plot width.");
        Assert.Contains(
            outcome.Result!.Scene.Warnings,
            warning => warning.Code == LabelPlanner.TruncatedToFitCode);
    }

    [Fact]
    public void Every_bar_and_marker_lies_inside_the_plot_bounds()
    {
        // The regression the plot-top offset fixes. The lane layout is lane-relative
        // and starts at y=0, so a placement used as a chart Y put every bar above the
        // plot, overlapping the header bands. Bars and markers are placed in chart
        // coordinates, so each must sit within the plot rectangle.
        SceneBuildRequest request = Request(
            Event(1),
            Event(2, GanttEntityType.AsPlannedMilestone, styleKey: "AsPlannedMilestone"),
            Event(3, GanttEntityType.AsBuiltActivity, styleKey: "AsBuiltActivity"));

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(request);

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        RectD plot = outcome.Result!.Scene.PlotBounds;

        SceneRect[] bars =
        [
            .. outcome.Result.Scene.Primitives
                .OfType<SceneRect>()
                .Where(rect => rect.PrimitiveId.EndsWith(":bar", StringComparison.Ordinal)),
        ];
        Assert.NotEmpty(bars);
        Assert.All(bars, AssertInsidePlot);

        // Milestones are four-point polygons, so the same containment rule is
        // asserted over their bounding box.
        ScenePolygon[] markers = [.. outcome.Result.Scene.Primitives.OfType<ScenePolygon>()];
        Assert.NotEmpty(markers);
        Assert.All(markers, marker => Assert.True(
            marker.Points.Min(point => point.Y) >= plot.Y
            && marker.Points.Max(point => point.Y) <= plot.Bottom,
            "A milestone marker must lie inside the plot bounds."));

        void AssertInsidePlot(SceneRect bar) => Assert.True(
            bar.Bounds.Y >= plot.Y && bar.Bounds.Bottom <= plot.Bottom,
            $"Bar {bar.PrimitiveId} at y={bar.Bounds.Y}..{bar.Bounds.Bottom} is outside the plot y={plot.Y}..{plot.Bottom}.");
    }

    [Fact]
    public void A_delineator_is_a_full_height_line_and_takes_no_lane()
    {
        GanttEvent first = Event(
            1,
            GanttEntityType.Delineator,
            new DateOnly(2024, 1, 8),
            null,
            styleKey: null);
        GanttEvent second = Event(2);

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(Request(first, second));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);

        // §24: the line spans PlotBounds.Top to PlotBounds.Bottom exactly.
        SceneLine line = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneLine>(),
            candidate => candidate.PrimitiveId.EndsWith(":delineator", StringComparison.Ordinal));
        Assert.Equal(_plotBounds.Y, line.From.Y, precision: 9);
        Assert.Equal(_plotBounds.Bottom, line.To.Y, precision: 9);

        // A delineator is not lane-bound, so the one activity bar must still be the
        // only bar: the line must not have consumed a lane.
        Assert.Single(outcome.Result.Scene.Primitives.OfType<SceneRect>(), rect => rect.PrimitiveId.EndsWith(":bar", StringComparison.Ordinal));
    }

    [Fact]
    public void Same_date_delineators_with_different_stroke_overrides_each_keep_their_own_line()
    {
        // §24 draws "same-date lines ... once per resolved line style", and a
        // per-row StrokeColour override is part of the resolved line style, not a
        // separate concern. Grouping by the style *name* put both rows in one
        // group, DelineatorLayout then refused it as InconsistentLineStyle, and
        // SceneBuilder reported one group refusal -- so both rows lost their line
        // entirely and neither was drawn. Two lines must now survive, one per
        // resolved colour.
        //
        // The registry carries DefaultDelineator because that is how a real
        // workbook reaches this path: the per-row override is applied during style
        // resolution, so a registry that cannot resolve the style never sees the
        // override at all and the two rows would resolve identically -- making the
        // assertion below vacuous for the wrong reason.
        GanttEvent plain = Event(1, GanttEntityType.Delineator, new DateOnly(2024, 1, 8), null, styleKey: null);
        GanttEvent overridden = Event(
            2,
            GanttEntityType.Delineator,
            new DateOnly(2024, 1, 8),
            null,
            styleKey: null) with { StrokeColour = "#FF0000" };

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(Event(3), plain, overridden) with { Registry = RegistryWithDelineator() });

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);

        SceneLine[] lines =
        [
            .. outcome.Result!.Scene.Primitives
                .OfType<SceneLine>()
                .Where(line => line.PrimitiveId.EndsWith(":delineator", StringComparison.Ordinal)),
        ];

        Assert.Equal(2, lines.Length);
        Assert.DoesNotContain(
            outcome.Result.Scene.Warnings,
            warning => warning.Code == "DelineatorGroupRefused");

        // Both lines stand at the same X (one date) and carry the two resolved
        // stroke colours, which is exactly "one line per resolved line style".
        Assert.Equal(lines[0].From.X, lines[1].From.X);
        Assert.Equal(2, lines.Select(line => line.Style.StrokeColour).Distinct().Count());
        Assert.Contains(lines, line => line.Style.StrokeColour == ColourHex.Parse("#FF0000"));
    }

    /// <summary>The shared test styles plus a resolvable <c>DefaultDelineator</c> style.</summary>
    private static GanttStyleRegistry RegistryWithDelineator() => new(
    [
        .. SharedStyles,
        Style("DefaultDelineator", 8),
    ]);

    [Fact]
    public void A_rendered_panel_emits_a_cell_per_grid_column_with_invariant_dates()
    {
        // The panel pass populates cells in grid-column order and formats dates through
        // the approved invariant format. A null cell list would emit a panel with no
        // text at all, which is what this guards.
        // The test's default grid carries only Id and Description, so a date cell
        // needs a grid that includes the date columns: the cells are projected in
        // grid-column order, so a column that is not in the grid is not emitted.
        SceneBuildRequest request = Request(Event(1, start: new DateOnly(2024, 1, 5), finish: new DateOnly(2024, 1, 10)))
        with
        {
            Grid = PanelCellGrid.TryCreate(
                [new PanelColumn("Id", 40), new PanelColumn("Start", 80), new PanelColumn("Finish", 80)],
                10,
                ["Id", "Start", "Finish"]).Grid,
            Panel = new PanelTheme(
                new SceneStyle("BodyFill"),
                new SceneStyle("BodyText"),
                new SceneStyle("HeaderFill"),
                new SceneStyle("HeaderText"),
                new SceneStyle("Border")),
        };

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(request);

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        SceneText dateCell = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneText>(),
            text => text.PrimitiveId.EndsWith(":panel-text:Start", StringComparison.Ordinal));

        // ADR-0016 D1/D4: dd/mm/yyyy, culture-invariant, never a re-derived pattern.
        Assert.Equal("05/01/2024", dateCell.Text);
    }

    [Fact]
    public void Shuffled_events_produce_an_identical_scene_serialization()
    {
        GanttEvent first = Event(1, start: new DateOnly(2024, 1, 3), finish: new DateOnly(2024, 1, 8));
        GanttEvent second = Event(2, start: new DateOnly(2024, 1, 12), finish: new DateOnly(2024, 1, 18));
        GanttEvent third = Event(3, GanttEntityType.AsPlannedMilestone, styleKey: "AsPlannedMilestone");

        // Three orderings, including a reversal, must all serialize identically:
        // the scene is ordered by the R3.3 z-order keys, not by input order.
        string baseline = Serialize(Request(first, second, third));
        Assert.Equal(baseline, Serialize(Request(third, first, second)));
        Assert.Equal(baseline, Serialize(Request(second, third, first)));
    }

    [Fact]
    public void A_panel_cell_reproduces_the_worksheet_Id_and_Type_text()
    {
        // The panel is the visible worksheet restated on the chart, so an Id or Type
        // cell that shows something other than the worksheet value misreports the
        // data beside the shapes built from it. Type in particular must be the
        // catalogue DisplayName ('As-Planned Activity'), not the durable enum member
        // name ('AsPlannedActivity'), which the workbook never shows.
        SceneBuildRequest request = Request(Event(1)) with
        {
            Grid = PanelCellGrid.TryCreate(
                [new PanelColumn("Id", 120), new PanelColumn("Type", 120)],
                10,
                ["Id", "Type"]).Grid,
            Panel = new PanelTheme(
                new SceneStyle("BodyFill"),
                new SceneStyle("BodyText"),
                new SceneStyle("HeaderFill"),
                new SceneStyle("HeaderText"),
                new SceneStyle("Border")),
        };
        GanttEvent @event = request.Events[0];

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(request);

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        SceneText idCell = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneText>(),
            text => text.PrimitiveId.EndsWith(":panel-text:Id", StringComparison.Ordinal));
        SceneText typeCell = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneText>(),
            text => text.PrimitiveId.EndsWith(":panel-text:Type", StringComparison.Ordinal));

        Assert.Equal(@event.Id.Value, idCell.Text);
        Assert.Equal(EntityTypeCatalog.GetDefinition(@event.Type)!.DisplayName, typeCell.Text);
        // Non-vacuous: the enum member name is the exact failure this guards, and for
        // this type the two strings differ.
        Assert.NotEqual(@event.Type.ToString(), typeCell.Text);
    }

    [Fact]
    public void A_resolved_bar_style_carries_the_definition_hatch_and_outline_width()
    {
        // The renderer-equivalence rule makes the scene the single source of the
        // appearance, so a hatch or outline the resolved style carries but the scene
        // drops can never be drawn. The previous construction left both at their
        // SceneStyle defaults, silently discarding two resolved tokens.
        GanttStyleRegistry hatched = new(
        [
            new GanttStyleDefinition(
                "Hatched",
                new HashSet<GanttLabelPosition> { GanttLabelPosition.Inside, GanttLabelPosition.Auto },
                EntityColourCapability.Fill | EntityColourCapability.Stroke,
                GanttLabelPosition.Inside,
                "#92D050",
                "#404040",
                "#000000",
                GanttHatchPattern.Cross,
                0,
                0,
                1.25,
                8,
                8),
        ]);

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(Event(1, styleKey: "Hatched")) with { Registry = hatched });

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        SceneRect bar = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneRect>(),
            rect => rect.PrimitiveId.EndsWith(":bar", StringComparison.Ordinal));

        Assert.Equal(GanttHatchPattern.Cross, bar.Style.HatchPattern);
        Assert.Equal(1.25, bar.Style.OutlineWidthPt);
    }

    [Fact]
    public void A_milestone_emits_its_description_label()
    {
        // Milestone bounds must reach the label pass. Without them the planner finds
        // no entry for the row and skips the description, which left section 22's
        // "milestone description label" unreachable.
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(Event(
                1,
                GanttEntityType.AsPlannedMilestone,
                start: new DateOnly(2024, 1, 15),
                styleKey: "AsPlannedMilestone"))
            with { LabelStyle = new SceneStyle("DefaultText") });

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        SceneText description = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneText>(),
            text => text.OwnerId.Kind == SceneOwnerKind.Row
                && text.PrimitiveId.EndsWith(":label", StringComparison.Ordinal));

        Assert.Equal("Row 1", description.Text);
    }

    [Fact]
    public void A_label_in_another_lane_does_not_constrain_this_rows_free_space()
    {
        // The free-space measure is one-dimensional, so before the vertical filter a
        // label anywhere in the scene shrank every other row's measured gap. Both rows
        // here take an explicit Right position, which is the section 22 path that
        // measures FreeRight -- Auto would pick Inside and never reach that code,
        // leaving the test vacuous. The dates are chosen so row 1's label begins
        // about 10pt to the right of row 2's bar, which is far less than the 40pt
        // "Row 2" text needs: with row 1's label counted as an obstruction the free
        // gap is ~10pt and the label is truncated, and with it correctly ignored the
        // gap runs to the plot edge.
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(
                RightLabelled(1, new DateOnly(2024, 1, 7), new DateOnly(2024, 1, 7)),
                RightLabelled(2, new DateOnly(2024, 1, 5), new DateOnly(2024, 1, 6)))
            with { LabelStyle = new SceneStyle("DefaultText") });

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        SceneRect[] bars =
        [
            .. outcome.Result!.Scene.Primitives
                .OfType<SceneRect>()
                .Where(rect => rect.PrimitiveId.EndsWith(":bar", StringComparison.Ordinal)),
        ];
        // The rows really are in separate lanes, so their vertical bands cannot
        // overlap -- proven here rather than assumed, so the assertion below cannot be
        // vacuous if the lane layout ever changes.
        Assert.Equal(2, bars.Length);
        Assert.True(bars[0].Bounds.Bottom <= bars[1].Bounds.Top);

        SceneText[] labels =
        [
            .. outcome.Result.Scene.Primitives
                .OfType<SceneText>()
                .Where(text => text.OwnerId.Kind == SceneOwnerKind.Row
                    && text.PrimitiveId.EndsWith(":label", StringComparison.Ordinal)),
        ];
        Assert.Contains(labels, text => text.Text == "Row 1");
        SceneText second = Assert.Single(labels, text => text.Text == "Row 2");

        // With the vertical filter the two labels cannot collide, so row 2's
        // description is emitted whole; without it the measured gap is cut short and
        // the label is truncated with an ellipsis.
        Assert.DoesNotContain('…', second.Text);

        // Non-vacuity on the obstruction itself: row 1's label really does begin
        // inside the gap row 2 measures, so the filter is doing work here rather than
        // discarding an occupant that could never have limited anything.
        SceneText first = Assert.Single(labels, text => text.Text == "Row 1");
        Assert.True(first.TextBounds.Left > bars[1].Bounds.Right);
    }

    /// <summary>Builds an event whose description label is forced to the §22 Right position.</summary>
    /// <param name="row">The one-based row number.</param>
    /// <param name="start">The inclusive start date.</param>
    /// <param name="finish">The inclusive finish date.</param>
    /// <returns>The event with an explicit Right label position.</returns>
    private static GanttEvent RightLabelled(int row, DateOnly start, DateOnly finish) =>
        Event(row, start: start, finish: finish) with
        {
            Description = $"Row {row}",
            LabelPosition = GanttLabelPosition.Right,
        };

    [Fact]
    public void A_critical_milestone_is_planned_before_a_planned_activity()
    {
        // Section 22 ranks a critical milestone's description above a planned
        // activity's, so the higher-priority label must be offered its box first. The
        // loop is ordered by that rank, which a placement-order loop would invert.
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(
                Event(1, start: new DateOnly(2024, 1, 10), finish: new DateOnly(2024, 1, 20)),
                Event(2, GanttEntityType.CriticalMilestone, start: new DateOnly(2024, 1, 10)))
            with { LabelStyle = new SceneStyle("DefaultText") });

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        var activityId = outcome.Result!.Scene.Primitives
            .OfType<SceneRect>()
            .First(rect => rect.PrimitiveId.EndsWith(":bar", StringComparison.Ordinal))
            .PrimitiveId.Split(':')[0];
        var milestoneId = outcome.Result.Scene.Primitives
            .OfType<ScenePolygon>()
            .Single(marker => marker.PrimitiveId.EndsWith(":marker", StringComparison.Ordinal))
            .PrimitiveId.Split(':')[0];

        // The milestone is planned first, so it is never the row whose label was
        // displaced by the higher-ranked one. Both rows are accounted for, so a loop
        // that planned neither -- or mislabelled the owner -- would fail here.
        SceneText[] labels =
        [
            .. outcome.Result.Scene.Primitives
                .OfType<SceneText>()
                .Where(text => text.OwnerId.Kind == SceneOwnerKind.Row
                    && text.PrimitiveId.EndsWith(":label", StringComparison.Ordinal)),
        ];
        Assert.NotEmpty(labels);
        Assert.Contains(labels, text => text.PrimitiveId.StartsWith(milestoneId, StringComparison.Ordinal));
        Assert.All(
            labels,
            text => Assert.True(
                text.PrimitiveId.StartsWith(milestoneId, StringComparison.Ordinal)
                    || text.PrimitiveId.StartsWith(activityId, StringComparison.Ordinal),
                "Unexpected label owner: " + text.PrimitiveId));
    }

    [Fact]
    public void A_procurement_label_is_ordered_after_every_activity_label()
    {
        // §22's priority list reads "actual labels, planned labels, baseline labels,
        // procurement/custom labels" as four consecutive groups, so procurement sits
        // after all three activity groups. Ranking procurement with the activity
        // family it belongs to (as-built procurement at rank 3, tied with as-built
        // *activity*) let a procurement label outrank a planned or baseline activity
        // label, which is the reverse of the stated order.
        //
        // The order is asserted on the published rank rather than through a built
        // scene. The lane layout gives every row in a lane its own slot, so two
        // rows' description labels sit in separate vertical bands and never compete
        // for one box; a scene-level test would therefore pass under either ranking
        // and prove nothing.
        Assert.Equal(3, LabelPlacementPriority.For(GanttEntityType.AsBuiltActivity));
        Assert.Equal(4, LabelPlacementPriority.For(GanttEntityType.AsPlannedActivity));
        Assert.Equal(5, LabelPlacementPriority.For(GanttEntityType.BaselineActivity));

        // Every procurement type shares the custom rank and every activity rank is
        // strictly lower, which is what makes the two groups ordered rather than
        // interleaved.
        foreach (GanttEntityType procurement in new[]
        {
            GanttEntityType.AsBuiltProcurement,
            GanttEntityType.AsPlannedProcurement,
            GanttEntityType.BaselineProcurement,
        })
        {
            Assert.Equal(LabelPlacementPriority.For(GanttEntityType.CustomActivity), LabelPlacementPriority.For(procurement));
            Assert.True(
                LabelPlacementPriority.For(procurement) > LabelPlacementPriority.For(GanttEntityType.BaselineActivity),
                $"{procurement} must rank after every activity label.");
        }
    }

    [Fact]
    public void The_section_22_label_priority_covers_every_entity_type_in_order()
    {
        // The whole §22 list, pinned by name rather than by enum declaration order,
        // because the enum's numeric order is a workbook-schema contract that does
        // not follow §22. Asserting the complete map means an unranked or
        // mis-ordered type fails here rather than sorting silently to the end.
        (GanttEntityType Type, int Rank)[] expected =
        [
            (GanttEntityType.CriticalMilestone, 0),
            (GanttEntityType.AsPlannedMilestone, 1),
            (GanttEntityType.AsBuiltMilestone, 1),
            (GanttEntityType.BaselineMilestone, 1),
            (GanttEntityType.DelayEvent, 2),
            (GanttEntityType.AsBuiltActivity, 3),
            (GanttEntityType.AsPlannedActivity, 4),
            (GanttEntityType.BaselineActivity, 5),
            (GanttEntityType.AsBuiltProcurement, 6),
            (GanttEntityType.AsPlannedProcurement, 6),
            (GanttEntityType.BaselineProcurement, 6),
            (GanttEntityType.CustomActivity, 6),
            (GanttEntityType.CriticalInterval, 7),
            (GanttEntityType.Delineator, 8),
            (GanttEntityType.Splitter, 9),
            (GanttEntityType.Spacer, 9),
        ];

        // Every enum member is covered, so a type added later fails this test
        // until its rank is decided rather than inheriting int.MaxValue.
        Assert.Equal(
            expected.Select(entry => entry.Type).Order(),
            Enum.GetValues<GanttEntityType>().Order());

        foreach ((GanttEntityType type, int rank) in expected)
        {
            Assert.Equal(rank, LabelPlacementPriority.For(type));
        }

        // §22 order is non-decreasing, so no later group can outrank an earlier one.
        Assert.Equal(expected.Select(entry => entry.Rank), expected.Select(entry => entry.Rank).Order());
    }

    [Fact]
    public void An_unclipped_bar_does_not_take_the_clipped_date_fallback()
    {
        // Section 23's never-suppress fallback fires only for a bar the plot actually
        // cut. FullBounds used to be the whole plot width, so every bar narrower than
        // the plot read as clipped and every date label took the fallback box. A bar
        // wholly inside the plot must now be planned normally.
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(Event(1, start: new DateOnly(2024, 1, 8), finish: new DateOnly(2024, 1, 18)))
            with { LabelStyle = new SceneStyle("DefaultText") });

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        SceneRect bar = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneRect>(),
            rect => rect.PrimitiveId.EndsWith(":bar", StringComparison.Ordinal));

        // Positive control on the input geometry: this bar is wholly inside the plot,
        // so nothing about it could make it genuinely clipped.
        Assert.True(bar.Bounds.Left > _plotBounds.Left);
        Assert.True(bar.Bounds.Right < _plotBounds.Right);

        // A planned date label is never a plausible-looking wrong date (D-G11), so an
        // ellipsis would mean the clipped path was taken for a bar the plot never cut.
        Assert.DoesNotContain(
            outcome.Result!.Scene.Primitives.OfType<SceneText>().Where(
                text => text.PrimitiveId.Contains(":date-", StringComparison.Ordinal)),
            text => text.Text.Contains('…', StringComparison.Ordinal));
    }

    [Fact]
    public void An_unclipped_bar_with_no_room_suppresses_its_date_label()
    {
        // The discriminating case for the clipped derivation. Section 23 anchors a
        // finish date label to the Right of the bar, and the event's own Right
        // description label claims that same box first. The description is long enough
        // (160pt against a measured 80pt date) that the finish date label cannot be
        // placed beside it and the planner declines. For an UNCLIPPED bar the correct
        // outcome is suppression: the never-suppress rule applies only to a bar the
        // plot cut.
        //
        // With the old plot-width FullBounds every bar read as clipped, so this same
        // declined label was emitted as an unclamped fallback box instead -- which is
        // why the golden scene carried date labels for bars the plot never touched.
        // Asserting suppression therefore fails on the old code.
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(Event(1, start: new DateOnly(2024, 1, 8), finish: new DateOnly(2024, 1, 18))
                with
                {
                    LabelPosition = GanttLabelPosition.Right,
                    Description = "A description long enough to fill the whole right-hand gap",
                })
            with { LabelStyle = new SceneStyle("DefaultText") });

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        SceneRect bar = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneRect>(),
            rect => rect.PrimitiveId.EndsWith(":bar", StringComparison.Ordinal));

        // Positive control: the bar is wholly inside the plot, so nothing about the
        // input could make it genuinely clipped.
        Assert.True(bar.Bounds.Left > _plotBounds.Left);
        Assert.True(bar.Bounds.Right < _plotBounds.Right);

        SceneText[] rowTexts =
        [
            .. outcome.Result.Scene.Primitives
                .OfType<SceneText>()
                .Where(text => text.OwnerId.Kind == SceneOwnerKind.Row),
        ];

        // The description label was placed and really does occupy the finish slot,
        // so the date label had a reason to decline rather than being lost for want
        // of a subject.
        SceneText description = Assert.Single(
            rowTexts,
            text => text.PrimitiveId.EndsWith(":label", StringComparison.Ordinal));
        Assert.True(description.TextBounds.Left >= bar.Bounds.Right);

        // Declined and not clipped: suppressed, with no fallback box beside the bar.
        Assert.DoesNotContain(
            rowTexts,
            text => text.PrimitiveId.EndsWith(":date-finish", StringComparison.Ordinal));
    }

    [Fact]
    public void A_point_event_keeps_its_start_date_label_instead_of_being_refused()
    {
        // The regression the FullBoundsOf guard fixes. A milestone has a Start and no
        // Finish, so the unclipped extent was built with width 0. DateLabelBuilder
        // refuses a non-positive FullBounds, so every milestone's start date label was
        // refused and then dropped without a trace. A point event has nothing the
        // plot can shorten, so `visible` is returned and the label is planned.
        // A blank description is legal (R2.5 U2) and emits no description label, so
        // the start date label is not competing with one for the same gap. The
        // milestone sits late in the month so the gap to its Left is inside the
        // plot; a marker at the plot's own left edge has no room for a Left label and
        // would be suppressed for that reason rather than the one under test.
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(
                Event(1, GanttEntityType.AsPlannedMilestone, start: new DateOnly(2024, 1, 20), styleKey: "AsPlannedMilestone")
                with { Description = null })
            with { LabelStyle = new SceneStyle("DefaultText") });

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);

        // Positive control: the marker exists, so the row genuinely reached the
        // label pass and has something to anchor a label to.
        Assert.Contains(
            outcome.Result!.Scene.Primitives,
            primitive => primitive.ZLayer == ZLayer.Milestone);

        SceneText startLabel = Assert.Single(
            outcome.Result!.Scene.Primitives
                .OfType<SceneText>(),
            text => text.PrimitiveId.EndsWith(":date-start", StringComparison.Ordinal));
        Assert.StartsWith(
            GanttDateFormatting.Format(new DateOnly(2024, 1, 20), GanttDateDisplayFormat.DdMMyyyy),
            startLabel.Text,
            StringComparison.Ordinal);
        Assert.DoesNotContain(
            outcome.Result!.Scene.Warnings,
            warning => warning.Code == "DateLabelRefused");
    }

    [Fact]
    public void A_refused_date_label_is_reported_as_a_warning()
    {
        // A date-label refusal is a broken dependency, not a placement decision, and
        // must be visible in the scene rather than dropping the label silently. A
        // failing text-metrics seam makes the planner refuse: it cannot measure the
        // text it is required to place.
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(Event(1))
            with
            {
                LabelStyle = new SceneStyle("DefaultText"),
                Metrics = new FailingTextMetrics(),
            });

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);

        SceneWarning warning = Assert.Single(
            outcome.Result!.Scene.Warnings,
            candidate => candidate.Code == "DateLabelRefused");
        Assert.Equal(SceneOwnerKind.Row, warning.OwnerId.Kind);
        Assert.False(string.IsNullOrWhiteSpace(warning.Message));
        Assert.DoesNotContain(
            outcome.Result!.Scene.Primitives.OfType<SceneText>(),
            text => text.PrimitiveId.Contains(":date-", StringComparison.Ordinal));
    }

    /// <summary>
    /// A text-metrics seam that cannot measure, so the label planner refuses rather
    /// than guessing a width. Used to drive the date-label refusal path.
    /// </summary>
    private sealed class FailingTextMetrics : ITextMetrics
    {
        public bool TryMeasure(string text, out TextMeasurement? measurement)
        {
            measurement = null;
            return false;
        }
    }

    [Fact]
    public void A_bar_the_plot_clipped_still_shows_its_true_off_plot_date()
    {
        // The counterpart that must keep working: a bar running off the plot's left
        // edge is genuinely clipped, and section 23 forbids suppressing its true
        // date. The fallback box is what guarantees the date is emitted at all.
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(Event(1, start: new DateOnly(2023, 12, 20), finish: new DateOnly(2024, 1, 10)))
            with { LabelStyle = new SceneStyle("DefaultText") });

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        SceneText start = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneText>(),
            text => text.PrimitiveId.EndsWith(":date-start", StringComparison.Ordinal));

        // The true, off-plot start date -- not a clamped one and not a truncation.
        Assert.Equal("20/12/2023", start.Text);
    }

    private static string Serialize(SceneBuildRequest request)
    {
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(request);
        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        return SceneSnapshot.Serialize(outcome.Result!.Scene);
    }

    /// <summary>Builds the representative scene from the committed neutral fixture.</summary>
    /// <param name="rows">
    /// The rows to build from, or <see langword="null"/> to load the committed
    /// fixture. Passing rows explicitly is how the golden test proves that a
    /// shuffled input order produces an identical scene.
    /// </param>
    /// <returns>The built scene.</returns>
    internal static GanttScene BuildScene(IReadOnlyList<GanttRowDto>? rows = null)
    {
        GanttValidationOutcome outcome = rows is null
            ? ReferenceSceneFixture.LoadValidated()
            : GanttRowValidator.Validate([.. rows]);
        SceneBuildRequest request = new()
        {
            Events = outcome.Events.ToList(),
            Registry = ReferenceSceneBuilder.StyleRegistry,
            Grid = PanelCellGrid.TryCreate(
                [
                    new PanelColumn("Id", 80),
                    new PanelColumn("Type", 120),
                    new PanelColumn("Description", 180),
                    new PanelColumn("Start", 70),
                    new PanelColumn("Finish", 70),
                ],
                10,
                ["Id", "Type", "Description", "Start", "Finish"]).Grid,
            // The data panel and the plot are unioned by FrameBandsBuilder to find
            // the content origin, and the title band is placed above that origin.
            // The panel therefore cannot start at y=0 or the title would sit above
            // the chart; it starts below the band stack instead.
            PanelBounds = new RectD(0, 20, 520, 200),
            PlotBounds = new RectD(520, 110, 600, 290),
            Metrics = new FakeTextMetrics(static _ => 4.0, 10.0),
            LaneMetrics = new LaneLayoutMetrics(18, 3, 3, 2, 18, 9),
            FrameTheme = ReferenceSceneBuilder.FrameTheme,
            PlotStart = ReferenceSceneFixture.PlotStart,
            PlotFinish = ReferenceSceneFixture.PlotFinish,
            Scale = GanttTimeScale.Month,
            PeriodLabelFormat = GanttPeriodLabelFormat.MMM,
            DateFormat = GanttDateDisplayFormat.DdMMyyyy,
            Title = ReferenceSceneFixture.Title,
            GridLinePt = 0.5,
            MajorBoundaryPt = 1,
            MilestoneSizePt = 8,
            CriticalLinePt = 1,
            // The three header bands total 50pt and stack above the plot, which
            // starts at 110 so they all fit inside the chart. FrameBandsBuilder does
            // not itself check this fit (an R3.5 finding recorded in the work item),
            // so the caller must supply a plot top that accommodates them.
            TitleBandHeightPt = 14,
            YearBandHeightPt = 16,
            PeriodBandHeightPt = 20,
            // §24 delineator tokens. These must be positive: the builder refuses a
            // non-positive line width rather than defaulting one, so a request that
            // omits them has no line to draw.
            DelineatorLinePt = 1,
            DelineatorStackGapPt = 10,
            LabelGapPt = 2,
            LabelHeightPt = 8,
            LabelStyle = new SceneStyle("DefaultText", fillColour: ColourHex.Parse("#000000")),
        };

        SceneBuildOutcome build = SceneBuilder.TryBuild(request);
        Assert.True(build.Succeeded, "Scene build refused: " + build.Refusal);
        return build.Result!.Scene;
    }
}
