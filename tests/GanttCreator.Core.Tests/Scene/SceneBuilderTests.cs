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
    private static readonly RectD _plotBounds = new(200, 60, 300, 140);

    /// <summary>
    /// The major-boundary width every request in this file passes (ADR-0038 D2), named
    /// so the closing-boundary arithmetic is asserted from one place.
    /// </summary>
    private const double MajorBoundaryPt = 1;

    /// <summary>
    /// The reserved anchor row's default height (ADR-0038 D1). Used to state the
    /// expected closing boundary as <c>anchor + w/2</c> rather than as a literal, so
    /// the arithmetic is pinned and not the constant that satisfies it today.
    /// </summary>
    private const double AnchorRowPt = 0.25;

    /// <summary>
    /// The header-row overlap the plot-spanning shapes are lifted by (ADR-0037 D1).
    /// </summary>
    /// <remarks>
    /// Named so the delineator's lifted top is asserted as the computed overlap rather
    /// than as a literal that happens to match today. The delineator's top was left
    /// unlifted until a live probe on 2026-10-04 showed it sliding where the bands
    /// beside it stretched.
    /// </remarks>
    private const double TopOverlapPt = 0.5;

    private static GanttStyleDefinition Style(string key, double height, string textColour = "#000000") =>
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
            textColour,
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

    /// <summary>
    /// A uniform measured grid carrying exactly one height per projected panel row.
    /// </summary>
    /// <remarks>
    /// The count is the event count because the panel now reproduces every source
    /// row, and the measured heights are positional. A fixture that supplied one
    /// height for many rows would now be refused, which is the intended behaviour
    /// rather than an obstacle to work around.
    /// </remarks>
    private static PanelCellGrid Grid(int rowCount) =>
        PanelCellGrid.TryCreate(
            [new PanelColumn("Id", 40), new PanelColumn("Description", 160)],
            [.. Enumerable.Repeat(10.0, rowCount)],
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

    [Fact]
    public void A_delay_label_is_white_inside_its_body_and_black_outside_it()
    {
        // The end-to-end scene proof of entity guide section 17. The planner half is
        // covered in LabelPlannerTests; what this adds is that SceneBuilder actually
        // SUPPLIES the outside style. Before this change both LabelPlanRequest call
        // sites omitted OutsideTextStyle, so the section 17 branch was unreachable
        // and every delay label stayed white wherever it was placed - white on the
        // chart background, which is unreadable.
        //
        // A registry whose DelayEvent carries the real DelayText token, so the
        // white asserted here is the catalogue's value and not a literal.
        var delayRegistry = new GanttStyleRegistry(
        [
            Style("AsPlannedActivity", 8),
            Style("DelayEvent", 8, textColour: "#FFFFFF"),
        ]);

        // Two rows: a wide delay bar that keeps its label inside the red body, and a
        // narrow one whose label cannot fit and is pushed outside it.
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(
                Event(1, GanttEntityType.DelayEvent, new DateOnly(2024, 1, 2), new DateOnly(2024, 1, 20), "DelayEvent"),
                Event(2, GanttEntityType.DelayEvent, new DateOnly(2024, 1, 2), new DateOnly(2024, 1, 3), "DelayEvent"))
            with
            {
                Registry = delayRegistry,
                LabelStyle = new SceneStyle("DefaultText"),
            });

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);

        // Both rows contribute a description label, and the two must disagree on
        // colour: the wide bar's label sits inside the red body, the narrow bar's
        // does not. Asserting both colours are present is the section 17 proof - if
        // the outside style were never supplied, both labels would be white.
        string[] labelColours =
        [
            .. outcome.Result!.Scene.Primitives
                .OfType<SceneText>()
                .Where(text => text.PrimitiveId.EndsWith(":label", StringComparison.Ordinal))
                .Select(text => text.Style.TextColour?.ToString() ?? "(null)"),
        ];

        Assert.Contains("#FFFFFF", labelColours);
        Assert.Contains("#000000", labelColours);
    }

    [Fact]
    public void A_non_delay_label_is_not_recoloured_by_the_outside_style()
    {
        // The control for the row above. A planned activity's label must keep its own
        // style's text colour whether it is placed inside or outside, so the section
        // 17 switch cannot be quietly widened to every entity.
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(Event(1))
            with { LabelStyle = new SceneStyle("DefaultText") });

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);

        SceneText label = Assert.Single(
            outcome.Result!.Scene.Primitives
                .OfType<SceneText>(),
            text => text.OwnerId.Kind == SceneOwnerKind.Row
                && text.PrimitiveId.EndsWith(":label", StringComparison.Ordinal));

        // The shared Style helper gives AsPlannedActivity the #000000 DefaultText
        // value, so the label is black - but it got there through its OWN resolved
        // style, not through the outside style.
        Assert.Equal("#000000", label.Style.TextColour?.ToString());
    }

    private static SceneBuildRequest Request(params GanttEvent[] events) =>
        new()
        {
            Events = events,
            Registry = _registry,
            Grid = Grid(events.Length),
            PlotBounds = _plotBounds,
            Metrics = _metrics,
            LaneMetrics = _laneMetrics,
            FrameTheme = FrameTheme(),
            PlotStart = _plotStart,
            PlotFinish = _plotFinish,
            GridLinePt = 0.5,
            MajorBoundaryPt = 1,
            MilestoneSizePt = 8,
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
            RowHeightPt = 8,
            ChartPadding = ChartPaddingPt.Uniform(0),
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
            PlotBounds = new RectD(200, 20, 300, 140),
        };

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(request);

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        SceneValidationReport report = SceneValidator.Validate(outcome.Result!.Scene);
        Assert.True(report.IsClean, "Findings: " + string.Join("; ", report.Findings));
    }
    [Fact]
    public void The_outer_padding_is_included_and_translating_the_chart_to_a_zero_origin_changes_no_relationship()
    {
        // The product decision recorded in entity guide §1: the outer padding is part
        // of the chart bounds, and a renderer places the chart by translating the
        // bounds to a zero origin and carrying one delta. Excel cannot express a
        // negative shape offset, so a negative chart origin is a real placement case,
        // not a hypothetical one - the content union starts at the panel origin here,
        // so the derived bounds start at minus exactly one padding. The panel theme
        // is what supplies that origin: with no panel the union starts at the plot,
        // whose left edge is 200, and the derived bounds are positive.
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(Event(1), Event(2, GanttEntityType.AsPlannedMilestone)) with
            {
                ChartPadding = ChartPaddingPt.Uniform(6),
                Panel = new PanelTheme(
                    new SceneStyle("BodyFill"),
                    new SceneStyle("BodyText"),
                    new SceneStyle("HeaderFill"),
                    new SceneStyle("HeaderText"),
                    new SceneStyle("Border")),
                // A drawn panel belongs to a profile whose destination has no cells of
                // its own. The live profile is refused a panel (R4.8A D4), so a
                // panel-bearing test must name the profile it is really exercising.
                Profile = SceneCompositionProfile.Raster,
            });

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        GanttScene scene = outcome.Result!.Scene;

        // The negative origin is reachable, and negative by exactly one padding, on
        // the axis where the content genuinely extends left of zero: the panel's left
        // edge sits at 0 while the plot's is at 200, so the union starts at 0 and the
        // derived bounds start at minus one padding.
        //
        // The vertical origin is no longer asserted negative. It used to be, but only
        // because the caller supplied a `PanelBounds` whose top happened to be 0 - an
        // arbitrary rectangle, not a measurement. Now that the panel's own bounds are
        // derived, the content's top is the period band and the derived top is
        // positive. The subject of this test is that the translation is lossless, and
        // that holds for a negative origin on either axis.
        Assert.True(scene.ChartBounds.Left < 0, $"Expected a negative origin, got {scene.ChartBounds}.");
        Assert.Equal(-6, scene.ChartBounds.Left, precision: 9);

        // The padding's presence on all four sides is pinned in FrameBandsBuilderTests,
        // where the builder reports the content and title rectangles it derived the
        // bounds from. Here the contract under test is only that the translation is
        // lossless.

        // Now translate the whole scene to a zero origin - the operation a renderer
        // performs - and prove it is lossless. The check is against the ORIGINAL scene,
        // not against arithmetic restated in the test, so a translation that moved
        // only some primitive kinds, or moved a kind's two geometry members by
        // different amounts, would fail here.
        GanttScene placed = Translate(scene, -scene.ChartBounds.Left, -scene.ChartBounds.Top);

        Assert.Equal(0, placed.ChartBounds.Left, precision: 9);
        Assert.Equal(0, placed.ChartBounds.Top, precision: 9);
        Assert.Equal(scene.ChartBounds.Width, placed.ChartBounds.Width, precision: 9);
        Assert.Equal(scene.ChartBounds.Height, placed.ChartBounds.Height, precision: 9);
        Assert.Equal(scene.Primitives.Count, placed.Primitives.Count);

        foreach ((ScenePrimitive before, ScenePrimitive after) in scene.Primitives.Zip(placed.Primitives))
        {
            Assert.Equal(before.PrimitiveId, after.PrimitiveId);
            string which = $"{before.GetType().Name} {before.PrimitiveId}";
            Assert.True(
                GeometryMath.ApproximatelyEqual(LeftOf(before) - scene.ChartBounds.Left, LeftOf(after)),
                $"{which}: left {LeftOf(before)} did not shift by the delta; placed at {LeftOf(after)}.");
            Assert.True(
                GeometryMath.ApproximatelyEqual(TopOf(before) - scene.ChartBounds.Top, TopOf(after)),
                $"{which}: top {TopOf(before)} did not shift by the delta; placed at {TopOf(after)}.");
        }

        // Every kind must actually have been present, or the sweep above would have
        // proved nothing about three of the four geometry shapes.
        Assert.Contains(scene.Primitives, primitive => primitive is SceneRect);
        Assert.Contains(scene.Primitives, primitive => primitive is SceneLine);
        Assert.Contains(scene.Primitives, primitive => primitive is ScenePolygon);
        Assert.Contains(scene.Primitives, primitive => primitive is SceneText);

        // And the decisive one: translating must not change the scene's validity, so a
        // zero origin cannot be what makes a label fit or a plot containment pass.
        Assert.Equal(
            SceneValidator.Validate(scene).IsClean,
            SceneValidator.Validate(placed).IsClean);
    }

    /// <summary>
    /// Gets a primitive's left edge, whichever member holds its geometry.
    /// </summary>
    private static double LeftOf(ScenePrimitive primitive) =>
        primitive switch
        {
            SceneRect rect => rect.Bounds.Left,
            SceneLine line => line.From.X,
            ScenePolygon polygon => polygon.Points.Min(point => point.X),
            SceneText text => text.TextBounds.Left,
            _ => 0,
        };

    /// <summary>Gets a primitive's top edge, whichever member holds its geometry.</summary>
    private static double TopOf(ScenePrimitive primitive) =>
        primitive switch
        {
            SceneRect rect => rect.Bounds.Top,
            SceneLine line => line.From.Y,
            ScenePolygon polygon => polygon.Points.Min(point => point.Y),
            SceneText text => text.TextBounds.Top,
            _ => 0,
        };

    private static ScenePrimitive Shift(ScenePrimitive primitive, double dx, double dy) =>
        primitive switch
        {
            SceneRect rect => new SceneRect(
                rect.PrimitiveId, rect.OwnerId, rect.ZLayer, Shift(rect.Bounds, dx, dy), rect.Style,
                rect.EntityType, rect.LaneOrder, rect.StackIndex, rect.SortOrder),
            SceneLine line => new SceneLine(
                line.PrimitiveId, line.OwnerId, line.ZLayer,
                new PointD(line.From.X + dx, line.From.Y + dy),
                new PointD(line.To.X + dx, line.To.Y + dy),
                line.Style, line.EntityType, line.LaneOrder, line.StackIndex, line.SortOrder),
            ScenePolygon polygon => new ScenePolygon(
                polygon.PrimitiveId, polygon.OwnerId, polygon.ZLayer,
                [.. polygon.Points.Select(point => new PointD(point.X + dx, point.Y + dy))],
                polygon.Style, polygon.EntityType, polygon.LaneOrder, polygon.StackIndex, polygon.SortOrder),
            SceneText text => new SceneText(
                text.PrimitiveId, text.OwnerId, text.ZLayer, text.Text, Shift(text.TextBounds, dx, dy), text.Style, text.Alignment,
                text.EntityType, text.LaneOrder, text.StackIndex, text.SortOrder),
            _ => primitive,
        };

    private static RectD Shift(RectD bounds, double dx, double dy) =>
        new(bounds.X + dx, bounds.Y + dy, bounds.Width, bounds.Height);

    private static GanttScene Translate(GanttScene scene, double dx, double dy)
    {
        List<ScenePrimitive> primitives = [.. scene.Primitives.Select(primitive => Shift(primitive, dx, dy))];
        SceneCreationOutcome outcome = GanttScene.TryCreate(
            Shift(scene.ChartBounds, dx, dy),
            Shift(scene.PlotBounds, dx, dy),
            primitives,
            scene.Warnings);
        Assert.True(outcome.Succeeded, "The translated scene was refused: " + outcome.Refusal);
        return outcome.Scene!;
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
    public void A_delineator_only_scene_is_not_empty_because_a_delineator_renders()
    {
        // §24 makes a Delineator a full-height plot line that consumes no lane. It is
        // therefore absent from the lane participants, and treating "no lane-bound
        // event" as "empty scene" refused a chart that has a line to draw.
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(Event(1, GanttEntityType.Delineator, styleKey: null, start: new DateOnly(2024, 1, 8), finish: null)));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        SceneLine line = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneLine>(),
            candidate => candidate.ZLayer == ZLayer.Delineator);
        Assert.Equal(_plotBounds.Top - TopOverlapPt, line.From.Y);
        Assert.Equal(_plotBounds.Bottom + AnchorRowPt + (MajorBoundaryPt / 2), line.To.Y);
    }

    /// <summary>
    /// The delineator's top lift is INDEPENDENT of the anchor row, and both ends are
    /// asserted together (ADR-0037 D1 / ADR-0038 D2).
    /// </summary>
    /// <remarks>
    /// <para>
    /// This is the test for the asymmetry, which is the finding most likely to be
    /// re-broken by a well-meaning cleanup. The two ends work for DIFFERENT reasons
    /// and are therefore configured by DIFFERENT tokens: the top reaches into an
    /// existing header row and needs no reservation, while the bottom needs the
    /// reserved anchor row because only a reserved row can create the cell anchor
    /// below the insert point.
    /// </para>
    /// <para>
    /// <b>The zero-anchor-row case is the load-bearing half.</b> It sets the anchor row
    /// to zero -- removing the bottom extension entirely -- and asserts the top is
    /// STILL lifted. A "symmetrisation" that tied the top to the anchor row, or a
    /// builder that simply extended the plot box at both ends, would leave the
    /// delineator sliding again the moment the anchor row is configured away, which is
    /// the opposite of what that token is for.
    /// </para>
    /// <para>
    /// The label corners are asserted NOT to move with the line: §22/§24 place them
    /// against <c>PlotBounds</c>, so widening the plot box to carry the overhang would
    /// drag every corner label with it.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_delineator_lifts_into_the_header_independently_of_the_anchor_row()
    {
        GanttEvent only = Event(
            1,
            GanttEntityType.Delineator,
            new DateOnly(2024, 1, 8),
            null,
            styleKey: null);

        // No anchor row: the bottom extension is off entirely.
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(Request(only) with { ChartAnchorRowHeightPt = 0 });
        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);

        SceneLine line = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneLine>(),
            candidate => candidate.PrimitiveId.EndsWith(":delineator", StringComparison.Ordinal));

        // The TOP is lifted even with no anchor row -- this is the assertion that
        // fails if the two ends are ever coupled.
        Assert.Equal(_plotBounds.Y - TopOverlapPt, line.From.Y, precision: 9);

        // ...and the bottom is back on the plot's own edge, because the anchor row
        // that would have carried it past is not there.
        Assert.Equal(_plotBounds.Bottom, line.To.Y, precision: 9);

        // The closing line is not emitted either, so the two switch off as a package.
        Assert.DoesNotContain(
            outcome.Result.Scene.Primitives.OfType<SceneLine>(),
            candidate => string.Equals(candidate.PrimitiveId, "chart:plot-closing", StringComparison.Ordinal));

        // The label corner stays against PlotBounds: a lifted TOP must not drag it up.
        SceneText label = Assert.Single(
            outcome.Result.Scene.Primitives.OfType<SceneText>(),
            candidate => candidate.PrimitiveId.EndsWith(":delineator-label", StringComparison.Ordinal));
        Assert.Equal(_plotBounds.Y, label.TextBounds.Y, precision: 9);
    }


    [Fact]
    public void A_delineator_style_the_registry_supplies_is_preferred_over_the_catalogue_fallback()
    {
        // P1-4: the delineator fallback used to restate the stroke colour as a literal,
        // making the token table and SceneBuilder two sources of truth. Both paths now
        // read the catalogue, so a registry that supplies its own style wins and the
        // fallback agrees with it rather than overriding it.
        GanttStyleDefinition registryDelineator = new(
            "DefaultDelineator",
            new HashSet<GanttLabelPosition> { GanttLabelPosition.Auto, GanttLabelPosition.None },
            EntityColourCapability.Stroke,
            GanttLabelPosition.Auto,
            null,
            "#FF00FF",
            "#000000",
            GanttHatchPattern.None,
            0,
            0,
            0.5,
            0,
            8);
        GanttStyleRegistry registryWithDelineator = new(SharedStyles.Append(registryDelineator));

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(Event(1, GanttEntityType.Delineator, styleKey: null, start: new DateOnly(2024, 1, 8), finish: null))
                with { Registry = registryWithDelineator });

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        SceneLine line = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneLine>(),
            candidate => candidate.ZLayer == ZLayer.Delineator);
        Assert.Equal(ColourHex.Parse("#FF00FF"), line.Style.StrokeColour);
    }

    [Fact]
    public void An_unresolvable_delineator_takes_the_catalogue_preset_not_a_restated_literal()
    {
        // The fallback path is proven reachable and proven to read the catalogue: the
        // expected colour is the DelineatorStroke token, never a value written into
        // this test, so a change to the token moves this assertion with it.
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(Event(1, GanttEntityType.Delineator, styleKey: null, start: new DateOnly(2024, 1, 8), finish: null)));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        SceneLine line = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneLine>(),
            candidate => candidate.ZLayer == ZLayer.Delineator);
        Assert.Equal(
            ColourHex.Parse(GanttCatalogues.GetPreset("DefaultDelineator").StrokeColour),
            line.Style.StrokeColour);
    }

    [Fact]
    public void Several_delineators_only_scene_builds_one_line_per_resolved_style()
    {
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(
                Event(1, GanttEntityType.Delineator, styleKey: null, start: new DateOnly(2024, 1, 8), finish: null),
                Event(2, GanttEntityType.Delineator, styleKey: null, start: new DateOnly(2024, 1, 20), finish: null)));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        // Two different dates are two groups, so two lines; neither consumed a lane.
        Assert.Equal(2, outcome.Result!.Scene.Primitives.OfType<SceneLine>()
            .Count(candidate => candidate.ZLayer == ZLayer.Delineator));
    }

    [Fact]
    public void A_splitter_and_spacer_only_scene_builds_the_splitter_band_and_no_spacer_primitive()
    {
        // This test previously asserted EmptyEvents, which encoded the defect as
        // contractual: both rows were dropped before lane layout, so §10/§11 lane
        // geometry was unreachable and a Splitter never rendered. The result is now
        // defined explicitly — a Splitter emits a band, a Spacer emits nothing.
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(
                Event(1, GanttEntityType.Splitter, styleKey: null),
                Event(2, GanttEntityType.Spacer, styleKey: null)));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        // §10: the band, plus its two major borders.
        Assert.Equal(
            3,
            outcome.Result!.Scene.Primitives.Count(
                candidate => candidate.PrimitiveId.Contains(SplitterBuilder.BandRole, StringComparison.Ordinal)));
        // §11: "no foreground fill, border, or label" — a Spacer owns no primitive.
        Assert.DoesNotContain(
            outcome.Result.Scene.Primitives,
            candidate => candidate.PrimitiveId.Contains(":spacer", StringComparison.Ordinal));
    }

    [Fact]
    public void A_splitter_between_two_activities_displaces_the_second_by_exactly_the_splitter_height()
    {
        GanttEvent first = Event(1, GanttEntityType.AsPlannedActivity, styleKey: "AsPlannedActivity");
        GanttEvent splitter = Event(2, GanttEntityType.Splitter, styleKey: null);
        GanttEvent second = Event(3, GanttEntityType.AsPlannedActivity, styleKey: "AsPlannedActivity");

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(Request(first, splitter, second));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        double firstTop = BarTop(outcome, first);
        double secondTop = BarTop(outcome, second);

        // The lone activity lane is LaneHeightPt tall and the splitter is
        // SplitterHeightPt, so the second bar sits exactly one splitter height lower.
        Assert.Equal(_laneMetrics.LaneHeightPt + _laneMetrics.SplitterHeightPt, secondTop - firstTop);
    }

    [Fact]
    public void A_splitter_and_a_spacer_between_two_activities_each_contribute_exactly_once()
    {
        GanttEvent first = Event(1, GanttEntityType.AsPlannedActivity, styleKey: "AsPlannedActivity");
        GanttEvent splitter = Event(2, GanttEntityType.Splitter, styleKey: null);
        GanttEvent spacer = Event(3, GanttEntityType.Spacer, styleKey: null);
        GanttEvent second = Event(4, GanttEntityType.AsPlannedActivity, styleKey: "AsPlannedActivity");

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(Request(first, splitter, spacer, second));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        Assert.Equal(
            _laneMetrics.LaneHeightPt + _laneMetrics.SplitterHeightPt + _laneMetrics.SpacerHeightPt,
            BarTop(outcome, second) - BarTop(outcome, first));
    }

    [Fact]
    public void A_spacer_only_scene_is_not_empty_because_the_spacer_still_occupies_a_lane()
    {
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(Event(1, GanttEntityType.Spacer, styleKey: null)));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        // A Spacer renders no primitive, so the frame alone is the correct output.
        Assert.DoesNotContain(
            outcome.Result!.Scene.Primitives,
            candidate => candidate.PrimitiveId.Contains(":spacer", StringComparison.Ordinal));
    }

    /// <summary>Reads one event's bar top from a built scene.</summary>
    private static double BarTop(SceneBuildOutcome outcome, GanttEvent @event) =>
        Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneRect>(),
            candidate => candidate.PrimitiveId == ScenePrimitive.CreateId(SceneOwnerId.ForRow(@event.Id), "bar")).Bounds.Y;

    [Fact]
    public void A_milestone_emits_its_diamond_and_reads_only_the_start_date()
    {
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(Event(1, GanttEntityType.AsPlannedMilestone, styleKey: "AsPlannedMilestone")));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        Assert.Contains(outcome.Result!.Scene.Primitives, p => p.ZLayer == ZLayer.Milestone);
    }

    /// <summary>
    /// A critical interval is clipped to the PLOT, not to its parent (owner ruling
    /// 2026-09-30). This test used to be named and commented as proving parent
    /// clipping, but its only assertion — that the overlay is narrower than the
    /// plot — is satisfied by any child that fits inside the plot, so it proved
    /// nothing about the parent at all. It is restated here to discriminate: the
    /// child's OWN dates run past the plot finish and past its parent's finish,
    /// and the overlay must stop exactly at the plot edge.
    /// </summary>
    [Fact]
    public void A_critical_interval_clips_to_the_plot_and_not_to_its_parent()
    {
        GanttEvent parent = Event(1, finish: new DateOnly(2024, 1, 10));
        GanttEvent child = Event(
            2,
            GanttEntityType.CriticalInterval,
            // Starts before the plot opens and finishes well after both the plot
            // and the parent end, so any of the three could be the clip source.
            new DateOnly(2023, 12, 20),
            new DateOnly(2024, 2, 20),
            "CriticalInterval",
            parentId: parent.Id);

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(Request(parent, child));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        SceneRect overlay = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneRect>(),
            rect => rect.ZLayer == ZLayer.CriticalOverlay);

        // Clipped to the plot on both edges, and not to the shorter parent. The
        // exact equality against the plot width IS the "was not clipped to the
        // parent" proof: a parent clip would stop the overlay short of the plot's
        // right edge and fail the second assertion.
        Assert.Equal(_plotBounds.Left, overlay.Bounds.Left, precision: 6);
        Assert.Equal(_plotBounds.Right, overlay.Bounds.Right, precision: 6);
        Assert.Equal(_plotBounds.Width, overlay.Bounds.Width, precision: 6);
    }

    /// <summary>
    /// ADR-0027 D2/D5: the overlay is a FILLED rectangle. A renderer receives the
    /// primitive and paints it, so the scene must actually carry a fill — an
    /// unfilled or outline-only rect here is the defect the ADR removes, not a
    /// variant. This asserts the fill the catalogue preset resolves.
    /// </summary>
    [Fact]
    public void A_critical_interval_overlay_resolves_a_fill_not_just_an_outline()
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
        SceneRect overlay = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneRect>(),
            rect => rect.ZLayer == ZLayer.CriticalOverlay);

        Assert.NotNull(overlay.Style.FillColour);
        // Half the predetermined ActivityHeightPt (8 in this fixture), per D3.
        Assert.Equal(4, overlay.Bounds.Height, precision: 6);
    }

    [Fact]
    public void A_critical_interval_whose_parent_is_invisible_is_still_drawn()
    {
        // This used to assert that NO overlay is emitted, because the builder
        // clipped to the parent's visible bar and a hidden parent emitted none.
        // Owner ruling 2026-09-30 removes that dependency: the interval draws from
        // its OWN dates in its own lane, so a hidden parent does not erase it. The
        // parent still governs LANE membership via R4.7B projection — it just no
        // longer governs whether the entity is drawn.
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
        SceneRect bar = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneRect>(),
            primitive => primitive.ZLayer == ZLayer.CriticalOverlay);
        // Half the parent's activity height, from the child's own dates.
        Assert.Equal(4, bar.Bounds.Height, 10);
    }

    [Fact]
    public void The_build_never_remeasures_the_supplied_grid()
    {
        SceneBuildRequest request = Request(Event(1));
        PanelCellGrid before = request.Grid!;

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(request);

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        // Core measures nothing: the grid the caller supplied must reach the panel
        // builder unchanged, with the same columns and the same measured row heights.
        // With per-row heights this is now a list, so it also proves no height was
        // collapsed into a single sample on the way through.
        Assert.Equal(before.RowHeightsPt, request.Grid!.RowHeightsPt);
        Assert.Equal(before.HeaderHeightPt, request.Grid.HeaderHeightPt);
        Assert.Equal(
            before.Columns.Select(column => column.Name),
            request.Grid.Columns.Select(column => column.Name));
    }


    [Fact]
    public void The_panel_header_band_bottom_equals_the_period_header_bottom()
    {
        // Section 4 fixes the header bottom to the period band's bottom. The
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
            // A drawn panel belongs to a profile whose destination has no cells of
            // its own; the live profile is refused one (R4.8A D4).
            Profile = SceneCompositionProfile.Raster,
        };

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(request);

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);

        // The period band sits DIRECTLY above the plot (entity guide §6), so its
        // bottom - and therefore the panel header's bottom - is the plot's own top
        // edge. This asserted `_plotBounds.Y - request.YearBandHeightPt`, which is
        // only correct while the YEAR band is the one adjacent to the plot. That
        // ordering was the inversion: it made the header align with a band that is
        // not there, so the value was wrong even though the test was green.
        double expectedBottom = _plotBounds.Y;

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

        // And the panel header must align with the period band the scene actually
        // emitted, not merely with a number this test computed. That is the
        // relationship §4 states, and computing the expectation here independently
        // is what let the two drift apart in the first place.
        SceneRect[] periodBands =
        [
            .. outcome.Result!.Scene.Primitives
                .OfType<SceneRect>()
                .Where(rect => rect.PrimitiveId.StartsWith("chart:period:", StringComparison.Ordinal)),
        ];
        Assert.NotEmpty(periodBands);
        Assert.All(
            periodBands,
            band => Assert.Equal(expectedBottom, band.Bounds.Bottom, precision: 9));
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
    public void An_external_label_is_bounded_only_by_the_free_space_beside_its_shape()
    {
        // ADR-0035 D1, INVERTED. This test previously asserted the OPPOSITE -- that a
        // label is capped by an absolute MaximumExternalLabelWidthPt -- and it stayed
        // green straight through the live defect for the same reason the band-order
        // test did: it pinned the produced literal rather than the relationship. The
        // product owner ruled that available space is the only limit on label length.
        //
        // The fixture is chosen so the two rules give DIFFERENT answers. The bar sits
        // at the very start of the month, so the gap to its right is roughly 270pt --
        // far more than the retired 144pt default cap. At 4pt per character a
        // 60-character description is 240pt: wider than 144, so the old cap would
        // have ellipsised it, and comfortably inside the free space, so the new rule
        // must emit it whole. A fixture with a narrow gap would have passed under
        // BOTH rules and proved nothing.
        const int DescriptionLength = 60;
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(
                Event(1, start: new DateOnly(2024, 1, 1), finish: new DateOnly(2024, 1, 3))
                with
                {
                    LabelPosition = GanttLabelPosition.Right,
                    Description = new string('W', DescriptionLength),
                })
            with
            {
                LabelStyle = new SceneStyle("DefaultText"),
                Metrics = new FakeTextMetrics(static _ => 4.0, 10.0),
            });

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);

        // Filtered to row-owned text: the frame's own period/year labels also end in
        // ":label", so an unfiltered match would pick a chart label instead.
        SceneText description = Assert.Single(
            outcome.Result!.Scene.Primitives
                .OfType<SceneText>(),
            text => text.OwnerId.Kind == SceneOwnerKind.Row
                && text.PrimitiveId.EndsWith(":label", StringComparison.Ordinal));

        // The whole string survives: no ellipsis, and byte-for-byte the row's text.
        Assert.Equal(new string('W', DescriptionLength), description.Text);
        Assert.DoesNotContain('…', description.Text);

        // The measurable form of the same claim. The box is strictly wider than the
        // 144pt the retired default cap imposed, so reinstating ANY absolute cap --
        // at 36, 144, or 360 -- fails here rather than passing quietly.
        Assert.True(
            description.TextBounds.Width > 144,
            $"A label must fill the free gap, not stop at the retired 144pt cap (was {description.TextBounds.Width}pt).");
    }

    /// <summary>
    /// A label in the FINAL lane is still contained: an 18pt box on a two-slot lane
    /// hangs below the plot, so the box is clamped up into the chart bounds.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The label box is one worksheet row tall (ADR-0033 D3) and centred on its
    /// shape's own band. On the LAST lane the second slot's bar sits within half a
    /// row of the plot's bottom edge, so a centred 18pt box extends past it — and past
    /// <c>ChartBounds</c>, which is the plot plus only <c>ChartOuterPaddingPt</c>. The
    /// containment check then refused the candidate, the widest-gap fallback measured
    /// the same out-of-bounds box, and the label was suppressed or ellipsised for
    /// being too tall rather than too narrow. Nothing said so: the symptom was a
    /// missing description on the bottom row.
    /// </para>
    /// <para>
    /// <b>Why the box moves rather than shrinks.</b> The height is one row by owner
    /// ruling, so it is a contract rather than something to fit; only the top moves,
    /// and only as far as containment requires. Every other row in the fixture fits
    /// untouched, so the change is confined to the box that would otherwise escape.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_label_in_a_two_slot_final_lane_is_contained_by_the_chart_bounds()
    {
        // Seven single-event lanes, then a FINAL lane holding two stacked events. The
        // lane heights are seeded explicitly because ADR-0034 anchors each lane to its
        // measured worksheet row, and the last row is what puts the second slot against
        // the plot's bottom edge.
        GanttRowId finalLane = GanttRowId.New();
        GanttEvent[] events =
        [
            .. Enumerable.Range(1, 7).Select(row => Event(row) with { LaneId = GanttRowId.New() }),
            Event(8) with { LaneId = finalLane },
            Event(9) with { LaneId = finalLane },
        ];

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(events) with
            {
                // The production token value: one worksheet row (ADR-0033 D3).
                RowHeightPt = GanttCatalogues.MetricDefault("GanttRowHeightPt"),
                PlotBounds = new RectD(200, 60, 300, 132),
                LabelStyle = new SceneStyle("DefaultText"),
                Grid = PanelCellGrid.TryCreate(
                    [new PanelColumn("Id", 40), new PanelColumn("Description", 160)],
                    [.. Enumerable.Repeat(18.0, 6), 24.0],
                    18,
                    ["Id", "Description"]).Grid!,
            });

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        GanttScene scene = outcome.Result!.Scene;

        // The FINAL lane's second slot is the lowest bar in the chart. Located through
        // its bar rather than by row number, so the assertion follows the geometry
        // instead of the fixture.
        SceneRect finalBar = scene.Primitives
            .OfType<SceneRect>()
            .Where(rect => rect.PrimitiveId.EndsWith(":bar", StringComparison.Ordinal))
            .MaxBy(rect => rect.Bounds.Bottom)!;

        SceneText finalLabel = Assert.Single(
            scene.Primitives.OfType<SceneText>(),
            text => text.OwnerId == finalBar.OwnerId
                && text.PrimitiveId.EndsWith(":label", StringComparison.Ordinal));

        double rowHeightPt = GanttCatalogues.MetricDefault("GanttRowHeightPt");
        RectD bounds = finalLabel.TextBounds;

        // The height is preserved exactly; only the position moves.
        Assert.Equal(rowHeightPt, bounds.Height);

        // Non-vacuity, as arithmetic rather than a hard-coded number: the UNCLAMPED
        // position is the box centred on its own bar, and on this final lane that box
        // hangs below the chart. So the placed box must sit strictly higher than the
        // centred one — Y grows downwards, so containing an overhanging box moves its
        // top up — and the fixture must be one where the centred box really would have
        // been refused.
        double centredTop = finalBar.Bounds.Top + ((finalBar.Bounds.Height - rowHeightPt) / 2);
        Assert.True(
            centredTop + rowHeightPt > scene.ChartBounds.Bottom,
            $"fixture does not exercise the clamp: the centred box already fits "
            + $"({centredTop}..{centredTop + rowHeightPt} within {scene.ChartBounds.Bottom}).");
        Assert.True(
            bounds.Top < centredTop,
            $"clamping must move the box UP into the chart, not down: {bounds.Top} vs centred {centredTop}.");

        // Contained on every edge — the containment rule is what previously refused
        // this candidate outright, and the reason the label vanished.
        Assert.True(bounds.Top >= scene.ChartBounds.Top, $"label top {bounds.Top} above {scene.ChartBounds.Top}.");
        Assert.True(
            bounds.Bottom <= scene.ChartBounds.Bottom,
            $"label bottom {bounds.Bottom} below {scene.ChartBounds.Bottom}.");

        // The horizontal anchor is untouched by a vertical clamp: the box still hugs
        // the bar it belongs to, so this is not a repositioning of the label itself.
        Assert.True(
            bounds.Left >= finalBar.Bounds.Right,
            $"a vertical clamp must not move the label horizontally: {bounds.Left} vs bar right {finalBar.Bounds.Right}.");
    }

    /// <summary>
    /// A long label in one row does NOT govern the width of a label in another row.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This is the reported symptom -- "one row's text length may be governing them
    /// all" -- as a property, exercised through <see cref="SceneBuilder"/> because
    /// that is the only path the product uses. Testing <see cref="LabelPlanner"/>
    /// directly with a hand-built occupant list proved nothing: the planner trusts
    /// the list it is handed and <c>SceneBuilder.VerticalBand</c> is what filters
    /// it, so a direct-planner test "reproduces" a symptom the product cannot
    /// produce.
    /// </para>
    /// <para>
    /// Row 1's description is 60 characters; row 2's is 6. Each label's box is
    /// asserted against its OWN text, so a shared width cannot pass.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_long_label_in_one_row_does_not_govern_another_rows_label()
    {
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(
                Event(1, start: new DateOnly(2024, 1, 8), finish: new DateOnly(2024, 1, 12))
                    with { LabelPosition = GanttLabelPosition.Right, Description = new string('a', 60) },
                Event(2, start: new DateOnly(2024, 1, 8), finish: new DateOnly(2024, 1, 12))
                    with { LabelPosition = GanttLabelPosition.Right, Description = "bbbbbb" })
            with { LabelStyle = new SceneStyle("DefaultText") });

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);

        List<SceneText> labels = outcome.Result!.Scene.Primitives
            .OfType<SceneText>()
            .Where(text => text.OwnerId.Kind == SceneOwnerKind.Row
                && text.PrimitiveId.EndsWith(":label", StringComparison.Ordinal))
            .ToList();

        Assert.Equal(2, labels.Count);

        SceneText longest = Assert.Single(labels, text => text.Text.StartsWith("aaaaa", StringComparison.Ordinal));
        SceneText shortest = Assert.Single(labels, text => text.Text == "bbbbbb");

        // Each label's box tracks ITS OWN text. A shared width would make these
        // equal, which is the whole claim.
        Assert.True(
            longest.TextBounds.Width > shortest.TextBounds.Width * 3,
            $"Each label must be sized from its own text ({longest.TextBounds.Width:0.###} vs {shortest.TextBounds.Width:0.###}).");

        // The short label is never squeezed by the long one: it keeps its full
        // natural width and stays inside its own row's band.
        Assert.Equal(6 * 8, shortest.TextBounds.Width, 1);
        Assert.True(
            shortest.TextBounds.Top > longest.TextBounds.Bottom,
            "A label must not borrow vertical space from another row.");
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

        // §24: the line spans PlotBounds.Top lifted into the header row (ADR-0037 D1) down
        // to the plot's closing boundary (ADR-0038 D2). The TOP is asserted too: the
        // delineator was originally left unlifted, and a live probe
        // (scripts/probe-delineator-top.ps1) showed that geometry SLID on a top insert
        // (TopDelta +15.75, HeightDelta 0) while the bands beside it stretched. It is
        // the one plot-spanning shape that had been left on the header/body boundary.
        //
        // Both ends are asserted, so a builder that lifted one and not the other cannot
        // pass -- and the lift is asserted as the computed overlap, not a literal.
        SceneLine line = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneLine>(),
            candidate => candidate.PrimitiveId.EndsWith(":delineator", StringComparison.Ordinal));
        Assert.Equal(_plotBounds.Y - TopOverlapPt, line.From.Y, precision: 9);
        Assert.Equal(
            _plotBounds.Bottom + AnchorRowPt + (MajorBoundaryPt / 2),
            line.To.Y,
            precision: 9);

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
                [10.0],
                10,
                ["Id", "Start", "Finish"]).Grid,
            Panel = new PanelTheme(
                new SceneStyle("BodyFill"),
                new SceneStyle("BodyText"),
                new SceneStyle("HeaderFill"),
                new SceneStyle("HeaderText"),
                new SceneStyle("Border")),
            // A drawn panel belongs to a profile whose destination has no cells of
            // its own; the live profile is refused one (R4.8A D4).
            Profile = SceneCompositionProfile.Raster,
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
                [10.0],
                10,
                ["Id", "Type"]).Grid,
            Panel = new PanelTheme(
                new SceneStyle("BodyFill"),
                new SceneStyle("BodyText"),
                new SceneStyle("HeaderFill"),
                new SceneStyle("HeaderText"),
                new SceneStyle("Border")),
            Profile = SceneCompositionProfile.Raster,
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

    /// <summary>
    /// Only the LIVE profile omits the drawn title band (ADR-0030 D6); every export
    /// profile still draws it.
    /// </summary>
    /// <remarks>
    /// <para>
    /// D6 removed the drawn title from the live sheet only, because the title is the
    /// table's own title cell there — and leaving the band in place drew it straight
    /// over the year band, which is what the live screenshot showed.
    /// </para>
    /// <para>
    /// <b>This is the counterweight that makes D6 safe.</b> Without the export half,
    /// a later "the title band is never drawn" rewrite would pass the live assertion
    /// alone and silently delete an entity the export renderers and the entity
    /// guide's field-contract table both depend on.
    /// </para>
    /// </remarks>
    [Theory]
    [InlineData(SceneCompositionProfile.LiveExcel, false)]
    [InlineData(SceneCompositionProfile.Raster, true)]
    [InlineData(SceneCompositionProfile.PowerPoint, true)]
    [InlineData(SceneCompositionProfile.EditableExport, true)]
    public void Only_the_live_profile_omits_the_drawn_title_band(
        SceneCompositionProfile profile,
        bool expectTitleBand)
    {
        SceneBuildOutcome build = SceneBuilder.TryBuild(
            PanelSceneRequest() with
            {
                Profile = profile,

                // A live profile must NOT carry a data panel (R4.8A D4): the live
                // sheet's own cells ARE the panel. The panel is supplied only for the
                // export profiles, so this test can vary the profile without the live
                // case being refused for an unrelated reason.
                Panel = profile == SceneCompositionProfile.LiveExcel ? null : PanelThemeValue,
            });

        Assert.True(build.Succeeded, "Scene build refused: " + build.Refusal);

        bool hasTitle = build.Result!.Scene.Primitives
            .Any(primitive => primitive.PrimitiveId.Contains("title", StringComparison.Ordinal));

        Assert.Equal(expectTitleBand, hasTitle);
    }

    /// <summary>The panel theme the export-profile cases supply.</summary>
    private static PanelTheme PanelThemeValue { get; } =
        new(
            new SceneStyle("DataPanelFill", fillColour: ColourHex.Parse("#F2F2F2")),
            new SceneStyle("DefaultText", fillColour: ColourHex.Parse("#000000")),
            new SceneStyle("HeaderFill", fillColour: ColourHex.Parse("#D9D9D9")),
            new SceneStyle("HeaderFontSizePt", fillColour: ColourHex.Parse("#000000"), bold: true),
            new SceneStyle("Border", strokeColour: ColourHex.Parse("#7F7F7F"), outlineWidthPt: 0.5));

    /// <summary>
    /// The panel-bearing reference request, before it is built, so a test can vary
    /// one field (the composition profile) without duplicating the whole request.
    /// </summary>
    /// <returns>The request <see cref="BuildSceneWithPanel"/> builds.</returns>
    private static SceneBuildRequest PanelSceneRequest()
    {
        GanttValidationOutcome outcome = ReferenceSceneFixture.LoadValidated();
        return new SceneBuildRequest
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
                [.. Enumerable.Repeat(10.0, outcome.Events.Count)],
                10,
                ["Id", "Type", "Description", "Start", "Finish"]).Grid,
            Panel = PanelThemeValue,
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
            TitleBandHeightPt = 14,
            YearBandHeightPt = 16,
            PeriodBandHeightPt = 20,
            DelineatorLinePt = 1,
            DelineatorStackGapPt = 10,
            LabelGapPt = 2,
            RowHeightPt = 8,
            LabelStyle = new SceneStyle("DefaultText", fillColour: ColourHex.Parse("#000000")),
        };
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
                [.. Enumerable.Repeat(10.0, outcome.Events.Count)], 10,                ["Id", "Type", "Description", "Start", "Finish"]).Grid,
            // The data panel and the plot are unioned by FrameBandsBuilder to find
            // the content origin, and the title band is placed above that origin.
            // The panel therefore cannot start at y=0 or the title would sit above
            // the chart; it starts below the band stack instead.
            PlotBounds = new RectD(520, 110, 600, 290),
            Metrics = new FakeTextMetrics(static _ => 4.0, 10.0),
            LaneMetrics = new LaneLayoutMetrics(18, 3, 3, 2, 18, 9),
            FrameTheme = ReferenceSceneBuilder.FrameTheme,

            // The canonical reference scene is an EXPORT composition, not a live
            // sheet. ADR-0030 D6 removed the drawn title band from the LIVE profile
            // only - the title entity still exists for export - and this scene is
            // what the entity guide's field-contract table is asserted against, so
            // leaving it on the LiveExcel default silently dropped the title entity
            // from every equivalence test. Stated rather than relied upon, so the
            // choice is visible where the expectations live.
            Profile = SceneCompositionProfile.Raster,
            PlotStart = ReferenceSceneFixture.PlotStart,
            PlotFinish = ReferenceSceneFixture.PlotFinish,
            Scale = GanttTimeScale.Month,
            PeriodLabelFormat = GanttPeriodLabelFormat.MMM,
            DateFormat = GanttDateDisplayFormat.DdMMyyyy,
            Title = ReferenceSceneFixture.Title,
            GridLinePt = 0.5,
            MajorBoundaryPt = 1,
            MilestoneSizePt = 8,
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
            RowHeightPt = 8,
            LabelStyle = new SceneStyle("DefaultText", fillColour: ColourHex.Parse("#000000")),
        };

        SceneBuildOutcome build = SceneBuilder.TryBuild(request);
        Assert.True(build.Succeeded, "Scene build refused: " + build.Refusal);
        return build.Result!.Scene;
    }

    /// <summary>
    /// Gets the reference fixture's first As-Planned Activity row, whose <c>:bar</c>
    /// and <c>:label</c> primitives the R3.16 field catalogue asserts against.
    /// </summary>
    internal static GanttRowId BarProbeRow { get; } = GanttRowId.Parse("G-000000000000000000000000000000b1");

    /// <summary>Gets the fixture's first milestone row, whose <c>:marker</c> is asserted.</summary>
    internal static GanttRowId MilestoneProbeRow { get; } = GanttRowId.Parse("G-000000000000000000000000000000f1");

    /// <summary>Gets the fixture's As-Planned Procurement row, whose <c>:bar</c> is asserted.</summary>
    internal static GanttRowId ProcurementProbeRow { get; } = GanttRowId.Parse("G-000000000000000000000000000000e1");

    /// <summary>Gets the fixture's first Critical Interval row, whose <c>:critical</c> is asserted.</summary>
    internal static GanttRowId CriticalProbeRow { get; } = GanttRowId.Parse("G-000000000000000000000000000000c1");

    /// <summary>
    /// Gets the fixture's first delineator row, which contributes to the shared
    /// same-date line and keeps its own label.
    /// </summary>
    internal static GanttRowId DelineatorProbeRow { get; } = GanttRowId.Parse("G-000000000000000000000000000000a2");

    /// <summary>
    /// Gets the identifier of the shared same-date delineator line. ADR-0017 gives a
    /// deduplicated entity a membership-sensitive owner, so the identifier carries
    /// both contributing rows joined by the pipe that no row id can contain.
    /// </summary>
    internal static string SharedDelineatorPrimitiveId { get; } = $"{GanttRowId.Parse("G-000000000000000000000000000000a2")}|{GanttRowId.Parse("G-000000000000000000000000000000a3")}:delineator";

    /// <summary>
    /// Gets the row whose panel cells the R3.16 data-panel row asserts against.
    /// </summary>
    /// <remarks>
    /// The canonical <see cref="BuildScene"/> supplies no <c>PanelTheme</c>, so the
    /// committed golden contains no panel primitive at all. The panel row is therefore
    /// exercised against <see cref="BuildSceneWithPanel"/>, a purpose-built scene that
    /// adds only the optional panel - it never touches the golden fixture.
    /// </remarks>
    internal static GanttRowId PanelProbeRow { get; } = GanttRowId.Parse("G-000000000000000000000000000000b1");

    /// <summary>
    /// Builds a scene that additionally carries the §3 data panel.
    /// </summary>
    /// <returns>A built scene containing the panel's cell, text, and border primitives.</returns>
    /// <remarks>
    /// <c>SceneBuilder</c> emits panel primitives only when a <see cref="PanelTheme"/>
    /// is supplied, and the canonical reference scene deliberately does not supply
    /// one, so the committed golden stays free of them. This builds the same scene with
    /// the panel added, which is what lets the R3.16 data-panel row be exercised at all.
    /// </remarks>
    internal static GanttScene BuildSceneWithPanel()
    {
        GanttValidationOutcome outcome = ReferenceSceneFixture.LoadValidated();
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
                [.. Enumerable.Repeat(10.0, outcome.Events.Count)], 10,                ["Id", "Type", "Description", "Start", "Finish"]).Grid,
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
            TitleBandHeightPt = 14,
            YearBandHeightPt = 16,
            PeriodBandHeightPt = 20,
            DelineatorLinePt = 1,
            DelineatorStackGapPt = 10,
            LabelGapPt = 2,
            RowHeightPt = 8,
            LabelStyle = new SceneStyle("DefaultText", fillColour: ColourHex.Parse("#000000")),
            Panel = new PanelTheme(
                new SceneStyle("DataPanelFill", fillColour: ColourHex.Parse("#F2F2F2")),
                new SceneStyle("DefaultText", fillColour: ColourHex.Parse("#000000")),
                new SceneStyle("HeaderFill", fillColour: ColourHex.Parse("#D9D9D9")),
                new SceneStyle("HeaderFontSizePt", fillColour: ColourHex.Parse("#000000"), bold: true),
                new SceneStyle("Border", strokeColour: ColourHex.Parse("#7F7F7F"), outlineWidthPt: 0.5)),
            // A drawn panel belongs to a profile whose destination has no cells of
            // its own; the live profile is refused one (R4.8A D4).
            Profile = SceneCompositionProfile.Raster,
        };

        SceneBuildOutcome build = SceneBuilder.TryBuild(request);
        Assert.True(build.Succeeded, "The panel-bearing scene build refused: " + build.Refusal);
        return build.Result!.Scene;
    }
}
