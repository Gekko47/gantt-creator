using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

public sealed class CriticalOverlayBuilderTests
{
    // A 31-day January 2024 plot, 0..310pt, so one day is exactly 10pt.
    private static readonly TimeScale _scale =
        TimeScale.TryCreate(new DateOnly(2024, 1, 1), new DateOnly(2024, 1, 31), 0, 310).Scale!;

    private static readonly SceneStyle _style = new("Critical", strokeColour: ColourHex.Parse("#FF0000"));
    private static readonly GanttRowId _parentId = GanttRowId.New();
    private static readonly GanttRowId _childId = GanttRowId.New();

    // The parent's visible bar: 5 Jan to 15 Jan, so 40..150pt.
    private static readonly RectD _parentBounds = new(40, 56, 110, 8);

    private static Dictionary<GanttRowId, RectD> Parents() => new() { [_parentId] = _parentBounds };

    [Fact]
    public void Clips_an_interval_fully_inside_the_parent_without_warning()
    {
        CriticalOverlayCreationOutcome outcome = Build(5, 9);

        Assert.True(outcome.Succeeded);
        SceneRect overlay = Assert.IsType<SceneRect>(outcome.Result!.Primitive);
        // 5 Jan is 40 and 9 Jan is 80, so the overlay spans 40..90.
        Assert.Equal(40, overlay.Bounds.Left);
        Assert.Equal(90, overlay.Bounds.Right, 10);
        Assert.False(outcome.Result.WasClipped);
        Assert.Empty(outcome.Result.Warnings);
    }

    [Fact]
    public void Sits_on_the_parent_top_edge_with_the_documented_thickness()
    {
        CriticalOverlayCreationOutcome outcome = Build(5, 9, thickness: 2.25);

        SceneRect overlay = Assert.IsType<SceneRect>(outcome.Result!.Primitive);
        // §16: the stroke stays inside the body, centred on the parent top edge.
        Assert.Equal(56, overlay.Bounds.Top);
        Assert.Equal(2.25, overlay.Bounds.Height, 10);
    }

    [Fact]
    public void Emits_at_the_critical_overlay_layer_with_a_role_derived_id()
    {
        CriticalOverlayCreationOutcome outcome = Build(5, 9);

        SceneRect overlay = Assert.IsType<SceneRect>(outcome.Result!.Primitive);
        Assert.Equal(ZLayer.CriticalOverlay, overlay.ZLayer);
        Assert.EndsWith(":critical", overlay.PrimitiveId, StringComparison.Ordinal);
    }

    [Fact]
    public void Clips_a_child_wider_than_its_parent_and_warns()
    {
        CriticalOverlayCreationOutcome outcome = Build(1, 20);

        Assert.True(outcome.Succeeded);
        SceneRect? overlay = outcome.Result!.Primitive;
        Assert.NotNull(overlay);
        // The overlay never extends past the parent's visible span.
        Assert.Equal(40, overlay!.Bounds.Left);
        Assert.Equal(150, overlay.Bounds.Right, 10);
        Assert.True(outcome.Result.WasClipped);
        Assert.Equal(CriticalOverlayBuilder.ClippedToParentCode, Assert.Single(outcome.Result.Warnings).Code);
    }

    [Fact]
    public void Emits_no_overlay_and_one_warning_when_the_child_is_entirely_before_the_parent()
    {
        CriticalOverlayCreationOutcome outcome = Build(1, 3);

        Assert.True(outcome.Succeeded);
        Assert.Null(outcome.Result!.Primitive);
        Assert.Equal(CriticalOverlayBuilder.OutsideParentCode, Assert.Single(outcome.Result.Warnings).Code);
    }

    [Fact]
    public void Emits_no_overlay_and_one_warning_when_the_child_is_entirely_after_the_parent()
    {
        CriticalOverlayCreationOutcome outcome = Build(20, 25);

        Assert.True(outcome.Succeeded);
        Assert.Null(outcome.Result!.Primitive);
        Assert.Equal(CriticalOverlayBuilder.OutsideParentCode, Assert.Single(outcome.Result.Warnings).Code);
    }

    [Fact]
    public void Produces_one_overlay_per_disjoint_child_with_exact_edges()
    {
        CriticalOverlayCreationOutcome first = Build(5, 7);
        CriticalOverlayCreationOutcome second = Build(10, 12);

        // 5-7 Jan is 40..70 and 10-12 Jan is 90..120, so the two never touch.
        Assert.Equal(70, first.Result!.VisibleBounds!.Value.Right, 10);
        Assert.Equal(90, second.Result!.VisibleBounds!.Value.Left, 10);
        Assert.Empty(first.Result.Warnings);
        Assert.Empty(second.Result.Warnings);
    }

    [Fact]
    public void Shares_the_exact_edge_between_adjacent_intervals()
    {
        // §16/checklist: adjacent intervals share an edge with no gap and no
        // double line. The inclusive span rule means a child finishing on the 7th
        // ends at 70pt, and a child starting on the 8th begins at exactly 70pt.
        CriticalOverlayCreationOutcome first = Build(5, 7);
        CriticalOverlayCreationOutcome second = Build(8, 10);

        Assert.Equal(70, first.Result!.VisibleBounds!.Value.Right, 10);
        Assert.Equal(70, second.Result!.VisibleBounds!.Value.Left, 10);
        Assert.False(first.Result.WasClipped);
        Assert.False(second.Result.WasClipped);
    }

    [Fact]
    public void Keeps_overlapping_intervals_at_their_own_spans_without_merging()
    {
        CriticalOverlayCreationOutcome first = Build(5, 10);
        CriticalOverlayCreationOutcome second = Build(8, 12);

        // Each child keeps its own overlay; neither is merged or extended.
        Assert.Equal(40, first.Result!.VisibleBounds!.Value.Left);
        Assert.Equal(100, first.Result.VisibleBounds!.Value.Right, 10);
        Assert.Equal(70, second.Result!.VisibleBounds!.Value.Left, 10);
        Assert.Equal(120, second.Result.VisibleBounds!.Value.Right, 10);
    }

    [Fact]
    public void Clips_to_a_parent_that_was_itself_plot_clipped()
    {
        // The parent bar starts before the plot, so R3.6 clips it to 0..90. The
        // overlay must clip to that *visible* span, not the parent's raw dates.
        var clippedParent = new RectD(0, 56, 90, 8);
        SpanBarResult bar = SpanBarBuilder
            .TryBuild(
                new SpanBarRequest(ParentEvent(new DateOnly(2023, 12, 28), new DateOnly(2024, 1, 8)), _style, 60, 8),
                _scale
            )
            .Result!;
        RectD visible = bar.VisibleBounds!.Value;

        Assert.Equal(0, visible.Left);
        CriticalOverlayCreationOutcome overlay = Build(
            1,
            4,
            parentBounds: visible,
            parents: new Dictionary<GanttRowId, RectD> { [_parentId] = clippedParent }
        );

        // The child runs 0..40 but the parent's visible span ends at 90, so the
        // overlay is confined to the visible parent and not to the raw parent.
        Assert.Equal(0, overlay.Result!.VisibleBounds!.Value.Left);
        Assert.Equal(40, overlay.Result.VisibleBounds!.Value.Right, 10);
    }

    [Fact]
    public void Never_increments_the_finish_date_for_a_one_day_interval()
    {
        CriticalOverlayCreationOutcome outcome = Build(5, 5);

        // A single-day interval keeps exactly one day of width.
        Assert.Equal(10, outcome.Result!.VisibleBounds!.Value.Width, 10);
        Assert.Empty(outcome.Result.Warnings);
    }

    [Fact]
    public void Refuses_an_orphan_with_a_typed_outcome_rather_than_throwing()
    {
        CriticalOverlayCreationOutcome outcome = CriticalOverlayBuilder.TryBuild(
            new CriticalOverlayRequest(
                new GanttEvent(
                    2,
                    _childId,
                    null,
                    null,
                    GanttEntityType.CriticalInterval,
                    "Critical",
                    new DateOnly(2024, 1, 5),
                    new DateOnly(2024, 1, 9),
                    null,
                    null,
                    null,
                    null,
                    null,
                    true,
                    null
                ),
                _style,
                _parentBounds,
                2.25
            ),
            _scale,
            Parents()
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(CriticalOverlayRefusal.UnresolvedParent, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_parent_that_resolves_to_no_bar()
    {
        CriticalOverlayCreationOutcome outcome = CriticalOverlayBuilder.TryBuild(
            new CriticalOverlayRequest(Child(5, 9), _style, _parentBounds, 2.25),
            _scale,
            new Dictionary<GanttRowId, RectD>()
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(CriticalOverlayRefusal.UnresolvedParent, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_reversed_interval()
    {
        CriticalOverlayCreationOutcome outcome = Build(9, 5);

        Assert.False(outcome.Succeeded);
        Assert.Equal(CriticalOverlayRefusal.InvalidInterval, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_non_critical_interval_type()
    {
        CriticalOverlayCreationOutcome outcome = CriticalOverlayBuilder.TryBuild(
            new CriticalOverlayRequest(
                new GanttEvent(1, _childId, null, null, GanttEntityType.AsPlannedActivity, "x", new DateOnly(2024, 1, 5), new DateOnly(2024, 1, 9), _parentId, null, null, null, null, true, null),
                _style,
                _parentBounds,
                2.25
            ),
            _scale,
            Parents()
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(CriticalOverlayRefusal.NotACriticalInterval, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_zero_thickness_overlay()
    {
        CriticalOverlayCreationOutcome outcome = Build(5, 9, thickness: 0);

        Assert.False(outcome.Succeeded);
        Assert.Equal(CriticalOverlayRefusal.InvalidGeometry, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_null_request()
    {
        CriticalOverlayCreationOutcome outcome = CriticalOverlayBuilder.TryBuild(null, _scale, Parents());

        Assert.False(outcome.Succeeded);
        Assert.Equal(CriticalOverlayRefusal.NullRequest, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_null_time_scale()
    {
        CriticalOverlayCreationOutcome outcome = CriticalOverlayBuilder.TryBuild(
            new CriticalOverlayRequest(Child(5, 9), _style, _parentBounds, 2.25),
            null,
            Parents()
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(CriticalOverlayRefusal.InvalidTimeScale, outcome.Refusal);
    }

    [Fact]
    public void Produces_identical_geometry_across_repeated_runs()
    {
        SceneRect first = Assert.IsType<SceneRect>(Build(5, 9).Result!.Primitive);
        SceneRect second = Assert.IsType<SceneRect>(Build(5, 9).Result!.Primitive);
        SceneRect third = Assert.IsType<SceneRect>(Build(5, 9).Result!.Primitive);

        Assert.Equal(first.Bounds, second.Bounds);
        Assert.Equal(first.Bounds, third.Bounds);
    }

    private static CriticalOverlayCreationOutcome Build(
        int startDay,
        int finishDay,
        double thickness = 2.25,
        RectD? parentBounds = null,
        Dictionary<GanttRowId, RectD>? parents = null
    ) =>
        CriticalOverlayBuilder.TryBuild(
            new CriticalOverlayRequest(
                Child(startDay, finishDay),
                _style,
                parentBounds ?? _parentBounds,
                thickness
            ),
            _scale,
            parents ?? Parents()
        );

    private static GanttEvent Child(int startDay, int finishDay, GanttRowId? parentId = null) =>
        new(
            2,
            _childId,
            null,
            null,
            GanttEntityType.CriticalInterval,
            "Critical",
            new DateOnly(2024, 1, startDay),
            new DateOnly(2024, 1, finishDay),
            parentId ?? _parentId,
            null,
            null,
            null,
            null,
            true,
            null
        );

    private static GanttEvent ParentEvent(DateOnly start, DateOnly finish) =>
        new(
            1,
            _parentId,
            null,
            null,
            GanttEntityType.AsPlannedActivity,
            "Parent",
            start,
            finish,
            null,
            null,
            null,
            null,
            null,
            true,
            null
        );
}
