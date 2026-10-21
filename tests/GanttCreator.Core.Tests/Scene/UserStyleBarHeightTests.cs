using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// Tests for the R5.6a bar height of a USER-AUTHORED style (WI-A): the style's
/// own height is honoured when it names one, and a style that names none
/// inherits the height its row's own kind carries rather than rendering a
/// degenerate zero-height bar.
/// </summary>
public sealed class UserStyleBarHeightTests
{
    private static readonly LaneLayoutMetrics _laneMetrics = new(18, 3, 3, 2, 18, 9);

    // A deliberately tight lane: 8pt tall with 3pt of padding above and below,
    // so the 6pt of padding alone fits it and any resolved height above 2pt
    // does not. It is what makes a milestone's resolved height observable.
    private static readonly LaneLayoutMetrics TightLaneMetrics = new(8, 3, 3, 2, 18, 9);

    private static readonly ITextMetrics _textMetrics = new FakeTextMetrics(_ => 8.0, 10.0);

    // A panel at 0..200 and a plot at 200..500 inside a 0..520 chart, so the plot
    // is inside the chart by construction. The plot top is 60 because the three
    // header bands total 14 + 16 + 20 = 50pt and are placed above the plot.
    private static readonly RectD _plotBounds = new(200, 60, 300, 140);

    /// <summary>
    /// A user-authored style. A user style is one the workbook's own rows
    /// carry, so it is not a code-owned preset: its own height and its own
    /// milestone size are overrides of the code-owned defaults.
    /// </summary>
    private static GanttStyleDefinition UserStyle(string key, double activityHeightPt, double milestoneSizePt = 0) =>
        new(
            key,
            new HashSet<GanttLabelPosition> { GanttLabelPosition.Auto, GanttLabelPosition.Inside },
            EntityColourCapability.Fill | EntityColourCapability.Stroke,
            GanttLabelPosition.Inside,
            "#92D050",
            "#404040",
            "#000000",
            GanttHatchPattern.None,
            0,
            0,
            0.5,
            activityHeightPt,
            milestoneSizePt);

    private static GanttEvent Event(
        int row,
        GanttEntityType type,
        string? styleKey,
        DateOnly? start = null,
        DateOnly? finish = null) =>
        new(
            row,
            GanttRowId.New(),
            GanttRowId.New(),
            0,
            type,
            $"Row {row}",
            start ?? new DateOnly(2024, 1, 5),
            finish ?? (type is GanttEntityType.AsPlannedMilestone or GanttEntityType.CustomMilestone
                ? null
                : new DateOnly(2024, 1, 10)),
            null,
            styleKey,
            null,
            null,
            null,
            true,
            null);

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

    private static SceneBuildRequest Request(
        GanttStyleRegistry registry,
        params GanttEvent[] events) =>
        new()
        {
            Events = events,
            Registry = registry,
            Grid = Grid(events.Length),
            PlotBounds = _plotBounds,
            Metrics = _textMetrics,
            LaneMetrics = _laneMetrics,
            FrameTheme = FrameTheme(),
            LabelStyle = new SceneStyle("DefaultText"),
            PlotStart = new DateOnly(2024, 1, 1),
            PlotFinish = new DateOnly(2024, 12, 31),
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
            ChartPadding = ChartPaddingPt.Uniform(0),
            MinimumHeaderLabelWidthPt = 0,
        };

    /// <summary>
    /// The height of the row's bar, read from the scene's own primitive.
    /// </summary>
    private static double BarHeightPt(SceneBuildOutcome outcome)
    {
        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        SceneRect bar = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneRect>(),
            rect => rect.PrimitiveId.EndsWith(":bar", StringComparison.Ordinal));
        return bar.Bounds.Height;
    }

    /// <summary>
    /// A user-authored style that names its own bar height renders at THAT
    /// height. This is the capability WI-A exists for: the height the user
    /// typed in the custom style form is the height the bar is drawn at.
    /// </summary>
    [Fact]
    public void A_user_style_with_its_own_bar_height_renders_at_that_height()
    {
        GanttStyleRegistry registry = new([UserStyle("MyStyle", 14)]);

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(registry, Event(1, GanttEntityType.CustomActivity, "MyStyle")));

        Assert.Equal(14, BarHeightPt(outcome), 3);
    }

    /// <summary>
    /// A user-authored style that names NO bar height inherits the height its
    /// row's own kind carries -- what the row would render at with no custom
    /// style at all. Before WI-A the zero flowed straight through and the bar
    /// was drawn with NO height, which is the degenerate bar the style
    /// contract's "not applicable to this entity kind" value must not become.
    /// </summary>
    [Fact]
    public void A_user_style_with_no_bar_height_inherits_its_kind_s_height()
    {
        GanttStyleRegistry registry = new([UserStyle("MyStyle", 0)]);

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(registry, Event(1, GanttEntityType.CustomActivity, "MyStyle")));

        // The Custom Activity kind's own default style carries the code-owned
        // bar height.
        Assert.Equal(
            GanttCatalogues.GetPreset("CustomActivity").ActivityHeightPt,
            BarHeightPt(outcome),
            3);
    }

    /// <summary>
    /// A user-authored style used by a MILESTONE inherits the milestone
    /// kind's own height, which is zero: a milestone is a diamond, not a bar.
    /// The fallback reads the kind's default style rather than a constant, so
    /// a constant here would give every milestone a bar it does not have.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A milestone renders as a polygon, not a rect, so its resolved height is
    /// observed through the LANE it contributes to: the lane's content height
    /// is the sum of its slots' resolved heights, and a height that does not
    /// fit the lane is reported as
    /// <c>LaneContentExceedsRowHeight</c>.
    /// </para>
    /// <para>
    /// The lane here is deliberately tight (8pt, with 3pt of padding above and
    /// below), so a milestone resolving to the code-owned 8pt bar height would
    /// need 14pt and would be REPORTED, while a milestone resolving to zero
    /// needs only the 6pt of padding and is not. The second assertion is the
    /// control that proves the first is discriminating: the same tight lane
    /// with a style that DOES name a 14pt height is reported.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_user_style_on_a_milestone_inherits_the_milestone_kind_s_zero_height()
    {
        GanttStyleRegistry inherited = new([UserStyle("MyStyle", 0, 12)]);
        GanttStyleRegistry overridden = new([UserStyle("MyStyle", 14, 12)]);

        SceneBuildOutcome zeroOutcome = SceneBuilder.TryBuild(
            Request(inherited, Event(1, GanttEntityType.CustomMilestone, "MyStyle"))
            with { LaneMetrics = TightLaneMetrics });

        SceneBuildOutcome overrideOutcome = SceneBuilder.TryBuild(
            Request(overridden, Event(1, GanttEntityType.CustomMilestone, "MyStyle"))
            with { LaneMetrics = TightLaneMetrics });

        Assert.True(zeroOutcome.Succeeded, "Scene build refused: " + zeroOutcome.Refusal);
        Assert.True(overrideOutcome.Succeeded, "Scene build refused: " + overrideOutcome.Refusal);

        // The inherited height is zero, so the milestone fits the tight lane.
        Assert.DoesNotContain(
            zeroOutcome.Result!.Scene.Warnings,
            warning => warning.Code == LaneLayoutBuilder.LaneContentExceedsLaneHeightCode);

        // The control: a style that DOES name a height overflows the same lane,
        // so the absence above is a real measurement and not an inert assertion.
        Assert.Contains(
            overrideOutcome.Result!.Scene.Warnings,
            warning => warning.Code == LaneLayoutBuilder.LaneContentExceedsLaneHeightCode);

        // The milestone's own size is the style's own 12pt, so the style's
        // milestone size is honoured too.
        Assert.Equal(12, DiamondSpan(zeroOutcome), 3);
    }

    /// <summary>
    /// The tip-to-tip span of the milestone diamond the scene planned.
    /// </summary>
    private static double DiamondSpan(SceneBuildOutcome outcome)
    {
        ScenePolygon diamond = Assert.Single(outcome.Result!.Scene.Primitives.OfType<ScenePolygon>());
        return diamond.Points!.Max(point => point.X) - diamond.Points!.Min(point => point.X);
    }

    /// <summary>
    /// A BUILT-IN style is unaffected: the registry entry's own height is
    /// what it renders at, whatever the catalogue's default for its kind
    /// says. This is the built-in protection R5.6a requires, and the
    /// regression guard for the WI-A change.
    /// </summary>
    [Fact]
    public void A_built_in_style_keeps_its_own_height()
    {
        GanttStyleRegistry registry = new(
        [
            new GanttStyleDefinition(
                "AsPlannedActivity",
                new HashSet<GanttLabelPosition> { GanttLabelPosition.Auto, GanttLabelPosition.Inside },
                EntityColourCapability.Fill | EntityColourCapability.Stroke,
                GanttLabelPosition.Inside,
                "#92D050",
                "#404040",
                "#000000",
                GanttHatchPattern.None,
                0,
                0,
                0.5,
                11,
                0),
        ]);

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(
            Request(registry, Event(1, GanttEntityType.AsPlannedActivity, "AsPlannedActivity")));

        Assert.Equal(11, BarHeightPt(outcome), 3);
    }
}
