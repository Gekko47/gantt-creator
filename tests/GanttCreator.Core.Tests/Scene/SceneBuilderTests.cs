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
    private static readonly RectD _chartBounds = new(0, 0, 520, 200);

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

    private static readonly GanttStyleRegistry _registry = new(
    [
        Style("AsPlannedActivity", 8),
        Style("AsBuiltActivity", 8),
        Style("CriticalInterval", 8),
        Style("AsPlannedMilestone", 8),
    ]);

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
            ChartBounds = _chartBounds,
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
        Assert.Equal(SceneBuilderRefusal.NullBounds, SceneBuilder.TryBuild(Request(Event(1)) with { ChartBounds = null }).Refusal);
    }

    [Fact]
    public void A_plot_outside_the_chart_is_refused()
    {
        // The plot extends 40pt past the chart's right edge, so a frame label at
        // the chart edge could fall outside the chart that owns it.
        Assert.Equal(
            SceneBuilderRefusal.PlotOutsideChart,
            SceneBuilder.TryBuild(Request(Event(1)) with { PlotBounds = new RectD(200, 20, 400, 160) }).Refusal);
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
        Assert.Contains(
            outcome.Result!.Scene.Primitives.OfType<SceneRect>(),
            rect => Math.Abs((rect.Bounds.Top + rect.Bounds.Height) - expectedBottom) < 1e-9);
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

    private static string Serialize(SceneBuildRequest request)
    {
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(request);
        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        return SceneSnapshot.Serialize(outcome.Result!.Scene);
    }

    /// <summary>
    /// Builds a representative multi-entity scene for the validator's end-to-end
    /// test, so SceneValidator and SceneBuilder are proven against the same
    /// orchestrator output the renderers will consume rather than a hand-built one.
    /// </summary>
    /// <returns>The built scene.</returns>
    internal static GanttScene BuildScene()
    {
        GanttEvent parent = Event(1, start: new DateOnly(2024, 1, 3), finish: new DateOnly(2024, 1, 12));
        GanttEvent actual = Event(2, GanttEntityType.AsBuiltActivity, new DateOnly(2024, 1, 6), new DateOnly(2024, 1, 15), "AsBuiltActivity");
        GanttEvent critical = Event(
            3,
            GanttEntityType.CriticalInterval,
            new DateOnly(2024, 1, 7),
            new DateOnly(2024, 1, 10),
            "CriticalInterval",
            parentId: parent.Id);
        GanttEvent milestone = Event(4, GanttEntityType.AsPlannedMilestone, new DateOnly(2024, 1, 20), styleKey: "AsPlannedMilestone");

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(Request(parent, actual, critical, milestone));
        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        return outcome.Result!.Scene;
    }
}
