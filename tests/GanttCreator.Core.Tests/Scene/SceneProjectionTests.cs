using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// Scene-level tests for R4.7B's projection wiring: a projected child shares its
/// parent's lane, creates no lane of its own, and consumes no lane height.
/// </summary>
/// <remarks>
/// These sit apart from <c>SceneBuilderTests</c> because that file owns the general
/// build contract, while these assert one specific relationship — the render lane a
/// child draws on. The fixture is built here rather than reused so a change to the
/// larger fixture's helpers cannot silently alter what these prove.
/// </remarks>
public sealed class SceneProjectionTests
{
    private static GanttRowId NewId() => GanttRowId.New();

    private static GanttEvent Event(
        int row,
        GanttEntityType type = GanttEntityType.AsPlannedActivity,
        GanttRowId? parentId = null) =>
        new(
            row,
            NewId(),
            NewId(),
            0,
            type,
            $"Row {row}",
            new DateOnly(2024, 1, 5),
            type is GanttEntityType.AsPlannedMilestone ? null : new DateOnly(2024, 1, 10),
            parentId,
            "AsPlannedActivity",
            null,
            null,
            null,
            true,
            null);

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

    private static readonly GanttStyleRegistry Registry = new(SharedStyles);

    /// <summary>
    /// The injected text-metrics seam. Core cannot measure text, so the request must
    /// supply it; omitting it is a <c>NullMetrics</c> refusal, which is what these
    /// tests were first hitting.
    /// </summary>
    private static readonly ITextMetrics Metrics = new FakeTextMetrics(_ => 8.0, 10.0);

    private static readonly LaneLayoutMetrics LaneMetrics = new(18, 3, 3, 2, 18, 9);

    /// <summary>
    /// A uniform measured grid carrying one height per projected panel row, built
    /// through the same seam <c>SceneBuilderTests</c> uses so the panel does not
    /// refuse the request.
    /// </summary>
    private static PanelCellGrid Grid(int rowCount) =>
        PanelCellGrid.TryCreate(
            [new PanelColumn("Id", 40), new PanelColumn("Description", 160)],
            [.. Enumerable.Repeat(10.0, rowCount)],
            10,
            ["Id", "Description"]).Grid!;

    private static readonly DateOnly PlotStart = new(2024, 1, 1);
    private static readonly DateOnly PlotFinish = new(2024, 1, 31);
    private static readonly RectD PlotBounds = new(200, 60, 300, 140);

    private static FrameBandsTheme FrameTheme() =>
        new(
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
            Registry = Registry,
            Grid = Grid(events.Length),
            PlotBounds = PlotBounds,
            Metrics = Metrics,
            LaneMetrics = LaneMetrics,
            FrameTheme = FrameTheme(),
            PlotStart = PlotStart,
            PlotFinish = PlotFinish,
            GridLinePt = 0.5,
            MajorBoundaryPt = 1,
            MilestoneSizePt = 8,
            CriticalLinePt = 1,
            TitleBandHeightPt = 14,
            YearBandHeightPt = 16,
            PeriodBandHeightPt = 20,
            DelineatorLinePt = 1,
            DelineatorStackGapPt = 10,
            LabelGapPt = 2,
            LabelHeightPt = 8,
            ChartOuterPaddingPt = 0,
            MinimumHeaderLabelWidthPt = 0,
        };

    /// <summary>
    /// A child activity shares its parent's lane: both draw, and they occupy ONE
    /// lane rather than two. Two bars and one lane is the whole point of projection
    /// — the child has its own source row but renders inside the parent's lane.
    /// </summary>
    [Fact]
    public void A_child_activity_shares_its_parents_lane()
    {
        GanttRowId parentId = NewId();
        GanttEvent parent = Event(1) with { Id = parentId };
        GanttEvent child = Event(2, parentId: parentId);

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(Request(parent, child));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);

        // Both entities are drawn.
        Assert.Equal(2, LaneBars(outcome).Count);

        // ...on one lane, proven by a single vertical band.
        Assert.Equal(1, DistinctLaneBands(outcome));
    }

    /// <summary>
    /// The lane is merged in geometry, not merely in name: the child occupies the
    /// parent's vertical band rather than a lane of its own below it.
    /// </summary>
    [Fact]
    public void A_child_activity_renders_within_the_parents_lane_geometry()
    {
        GanttRowId parentId = NewId();
        GanttEvent parent = Event(1) with { Id = parentId };
        GanttEvent child = Event(2, parentId: parentId);

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(Request(parent, child));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);

        IReadOnlyList<SceneRect> bars = LaneBars(outcome);
        Assert.Equal(2, bars.Count);

        double top = bars.Min(b => b.Bounds.Top);
        double bottom = bars.Max(b => b.Bounds.Bottom);
        foreach (SceneRect bar in bars)
        {
            Assert.Equal(top, bar.Bounds.Top);
            Assert.Equal(bottom, bar.Bounds.Bottom);
        }
    }

    /// <summary>
    /// Two children on one parent produce ONE lane, so a projected child never grows
    /// the lane the projection exists to keep stable.
    /// </summary>
    [Fact]
    public void Two_children_on_one_parent_produce_a_single_lane()
    {
        GanttRowId parentId = NewId();
        GanttEvent parent = Event(1) with { Id = parentId };
        GanttEvent first = Event(2, parentId: parentId);
        GanttEvent second = Event(3, parentId: parentId);

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(Request(parent, first, second));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        Assert.Equal(1, DistinctLaneBands(outcome));
    }

    /// <summary>
    /// A projected child adds no lane height, which is why an expanded child consumes
    /// worksheet space without consuming plot space.
    /// </summary>
    [Fact]
    public void A_projected_child_does_not_add_lane_height()
    {
        GanttRowId parentId = NewId();
        GanttEvent parent = Event(1) with { Id = parentId };

        SceneBuildOutcome alone = SceneBuilder.TryBuild(Request(parent));
        SceneBuildOutcome withChild = SceneBuilder.TryBuild(Request(parent, Event(2, parentId: parentId)));

        Assert.True(alone.Succeeded, "Scene build refused: " + alone.Refusal);
        Assert.True(withChild.Succeeded, "Scene build refused: " + withChild.Refusal);

        Assert.Equal(VerticalExtent(alone), VerticalExtent(withChild));
    }

    /// <summary>
    /// Top-level rows still own one lane each, so a table with no children behaves
    /// exactly as it did before projection existed.
    /// </summary>
    [Fact]
    public void Top_level_rows_keep_one_lane_each()
    {
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(Request(Event(1), Event(2), Event(3)));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        Assert.Equal(3, DistinctLaneBands(outcome));
    }

    /// <summary>
    /// A child naming a parent outside the batch refuses the scene as a whole rather
    /// than rendering half the hierarchy — a partially projected scene is
    /// indistinguishable from a correct one once drawn.
    /// </summary>
    [Fact]
    public void A_child_of_an_absent_parent_refuses_the_scene()
    {
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(Request(Event(1, parentId: NewId())));

        Assert.False(outcome.Succeeded);
        Assert.Equal(SceneBuilderRefusal.UnresolvableProjection, outcome.Refusal);
    }

    /// <summary>
    /// A grandchild has no unambiguous render lane and is refused, matching
    /// <see cref="ProjectionResolver"/> and the validator's depth rule.
    /// </summary>
    [Fact]
    public void A_grandchild_refuses_the_scene()
    {
        GanttRowId parentId = NewId();
        GanttRowId childId = NewId();
        GanttEvent parent = Event(1) with { Id = parentId };
        GanttEvent child = Event(2, parentId: parentId) with { Id = childId };

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(Request(parent, child, Event(3, parentId: childId)));

        Assert.False(outcome.Succeeded);
        Assert.Equal(SceneBuilderRefusal.UnresolvableProjection, outcome.Refusal);
    }

    /// <summary>
    /// A child whose parent is present but NOT rendered keeps its own row-scoped lane
    /// and is not absorbed into a lane that does not exist. The overlay pass then
    /// declines to draw it, which is the pre-R4.7B behaviour for this case.
    /// </summary>
    [Fact]
    public void A_child_of_an_unrendered_parent_is_not_absorbed()
    {
        GanttRowId parentId = NewId();
        GanttEvent parent = Event(1) with { Id = parentId, Visible = false };
        GanttEvent child = Event(2, GanttEntityType.CriticalInterval, parentId);

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(Request(parent, child));

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);
        Assert.DoesNotContain(outcome.Result!.Scene.Primitives, p => p.ZLayer == ZLayer.CriticalOverlay);
    }

    /// <summary>
    /// The row-owned rectangles in the scene, which is what a lane looks like drawn.
    /// Selected by owner kind rather than by identifier text: the identifier is
    /// built from the owner plus a role, so a prefix check would depend on a naming
    /// convention that is not this test's subject.
    /// </summary>
    private static IReadOnlyList<SceneRect> LaneBars(SceneBuildOutcome outcome) =>
    [
        .. outcome.Result!.Scene.Primitives
            .OfType<SceneRect>()
            .Where(r => r.OwnerId.Kind == SceneOwnerKind.Row)

    ];

    /// <summary>
    /// The number of distinct vertical bands. Two events sharing a band share a lane;
    /// each additional band is a lane that grew.
    /// </summary>
    private static int DistinctLaneBands(SceneBuildOutcome outcome) =>
        LaneBars(outcome)
            .Select(r => (Top: r.Bounds.Top, Bottom: r.Bounds.Bottom))
            .Distinct()
            .Count();

    private static double VerticalExtent(SceneBuildOutcome outcome)
    {
        IReadOnlyList<SceneRect> bars = LaneBars(outcome);
        return bars.Count == 0 ? 0 : bars.Max(b => b.Bounds.Bottom) - bars.Min(b => b.Bounds.Top);
    }
}
