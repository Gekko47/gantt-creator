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

    // The visual slot this interval occupies: centre Y 60, so a bar of the default
    // 8pt predetermined height draws 56..64. There is no parent bar here at all
    // (owner ruling 2026-09-30) — the entity's geometry is its dates and its slot.
    private const double SlotCentreY = 60;

    [Fact]
    public void An_interval_spans_its_own_dates_without_warning()
    {
        CriticalOverlayCreationOutcome outcome = Build(5, 9);

        Assert.True(outcome.Succeeded);
        SceneRect bar = Assert.IsType<SceneRect>(outcome.Result!.Primitive);
        // 5 Jan is 40 and 9 Jan ends at 90, purely from its own dates.
        Assert.Equal(40, bar.Bounds.Left);
        Assert.Equal(90, bar.Bounds.Right, 10);
        Assert.False(outcome.Result.WasClipped);
        Assert.Empty(outcome.Result.Warnings);
    }

    [Fact]
    public void The_bar_height_is_half_the_predetermined_height()
    {
        // Owner ruling 2026-09-30 / ADR-0027 D3: the drawn height is HALF the
        // predetermined ActivityHeightPt. The retired CriticalLinePt thickness is
        // gone, so a caller cannot express an arbitrary overlay thickness at all.
        CriticalOverlayCreationOutcome outcome = Build(5, 9, predeterminedHeightPt: 2.25);

        SceneRect bar = Assert.IsType<SceneRect>(outcome.Result!.Primitive);
        Assert.Equal(1.125, bar.Bounds.Height, 10);
    }

    [Fact]
    public void The_bar_is_centred_on_its_own_slot()
    {
        // Centred exactly as an ordinary span bar is, rather than top-aligned to a
        // parent's bar that no longer participates.
        CriticalOverlayCreationOutcome outcome = Build(5, 9, predeterminedHeightPt: 8);

        SceneRect bar = Assert.IsType<SceneRect>(outcome.Result!.Primitive);
        Assert.Equal(4, bar.Bounds.Height, 10);
        Assert.Equal(58, bar.Bounds.Top, 10);
        Assert.Equal(60, bar.Bounds.Top + (bar.Bounds.Height / 2), 10);
    }

    [Fact]
    public void A_non_positive_predetermined_height_is_refused()
    {
        // The validator ships with this positive test: 0 was the silent value a
        // caller that forgot the retired CriticalLinePt used to pass, and it
        // refused the overlay, dropping the entity from the chart.
        CriticalOverlayCreationOutcome outcome =
            Build(5, 9, predeterminedHeightPt: 0);

        Assert.False(outcome.Succeeded);
        Assert.Equal(CriticalOverlayRefusal.InvalidGeometry, outcome.Refusal);
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
    public void An_interval_wider_than_its_parent_is_NOT_clipped()
    {
        // The decisive rule (owner ruling 2026-09-30): the critical interval may
        // NOT be for the full duration of its parent, and it is not limited to it
        // either. Its horizontal bounds come from its own dates like any other
        // activity, so 1-20 Jan draws 0..200pt even though the parent's bar was
        // 40..150pt. This is the exact case that used to emit ClippedToParent.
        CriticalOverlayCreationOutcome outcome = Build(1, 20);

        Assert.True(outcome.Succeeded);
        SceneRect bar = Assert.IsType<SceneRect>(outcome.Result!.Primitive);
        Assert.Equal(0, bar.Bounds.Left, 10);
        Assert.Equal(200, bar.Bounds.Right, 10);
        Assert.False(outcome.Result.WasClipped);
        Assert.Empty(outcome.Result.Warnings);
    }

    [Fact]
    public void An_interval_shorter_than_its_parent_is_drawn_at_its_own_width()
    {
        // The converse: 5-7 Jan inside a 5-15 Jan parent draws its own 30pt, not
        // the parent's full width.
        CriticalOverlayCreationOutcome outcome = Build(5, 7);

        Assert.True(outcome.Succeeded);
        SceneRect bar = Assert.IsType<SceneRect>(outcome.Result!.Primitive);
        Assert.Equal(40, bar.Bounds.Left, 10);
        Assert.Equal(70, bar.Bounds.Right, 10);
    }

    [Fact]
    public void An_interval_entirely_before_its_parent_is_still_drawn()
    {
        // Under the old rule this emitted no overlay at all. Now the parent is not
        // consulted, so 1-3 Jan is simply a bar in its own right.
        CriticalOverlayCreationOutcome outcome = Build(1, 3);

        Assert.True(outcome.Succeeded);
        SceneRect bar = Assert.IsType<SceneRect>(outcome.Result!.Primitive);
        Assert.Equal(0, bar.Bounds.Left, 10);
        Assert.Equal(30, bar.Bounds.Right, 10);
        Assert.Empty(outcome.Result.Warnings);
    }

    [Fact]
    public void An_interval_entirely_after_its_parent_is_still_drawn()
    {
        CriticalOverlayCreationOutcome outcome = Build(20, 25);

        Assert.True(outcome.Succeeded);
        SceneRect bar = Assert.IsType<SceneRect>(outcome.Result!.Primitive);
        Assert.Equal(190, bar.Bounds.Left, 10);
        Assert.Equal(250, bar.Bounds.Right, 10);
        Assert.Empty(outcome.Result.Warnings);
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
    public void An_interval_starting_before_the_plot_is_clipped_to_the_plot_only()
    {
        // The child starts 28 Dec, four days before the plot opens. Only the
        // visible portion may be emitted, and it must be emitted: dropping it
        // would lose real schedule information. There is no parent edge to clip to
        // as well, so the start collapses onto the plot's left edge and stays.
        var outcome = CriticalOverlayBuilder.TryBuild(
            new CriticalOverlayRequest(
                ChildOn(new DateOnly(2023, 12, 28), new DateOnly(2024, 1, 8)),
                _style,
                8,
                SlotCentreY
            ),
            _scale);

        Assert.True(outcome.Succeeded);
        SceneRect bar = Assert.IsType<SceneRect>(outcome.Result!.Primitive);
        // The start collapses onto 0 and 8 Jan is the inclusive right edge at 80.
        Assert.Equal(0, bar.Bounds.Left, 10);
        Assert.Equal(80, bar.Bounds.Right, 10);
        Assert.True(outcome.Result.WasClipped);
        Assert.Equal(CriticalOverlayBuilder.ClippedToPlotCode, Assert.Single(outcome.Result.Warnings).Code);
    }

    [Fact]
    public void An_interval_finishing_after_the_plot_is_clipped_to_the_plot_only()
    {
        // The mirror case: the child ends 5 Feb, well after the plot closes.
        var outcome = CriticalOverlayBuilder.TryBuild(
            new CriticalOverlayRequest(
                ChildOn(new DateOnly(2024, 1, 12), new DateOnly(2024, 2, 5)),
                _style,
                8,
                SlotCentreY
            ),
            _scale);

        Assert.True(outcome.Succeeded);
        SceneRect bar = Assert.IsType<SceneRect>(outcome.Result!.Primitive);
        // 12 Jan is 110 and the plot's right edge is 310.
        Assert.Equal(110, bar.Bounds.Left, 10);
        Assert.Equal(310, bar.Bounds.Right, 10);
        Assert.True(outcome.Result.WasClipped);
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
    public void An_orphan_is_drawn_rather_than_refused()
    {
        // The old builder refused an orphan with UnresolvedParent because it needed
        // the parent's bar to clip against. With no parent in the request there is
        // nothing left to refuse on, so the entity draws from its dates alone.
        // R2.5 still blocks a parentless critical interval at VALIDATION time; this
        // is defence in depth no longer being needed, not a relaxation of it.
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
                8,
                SlotCentreY
            ),
            _scale);

        Assert.True(outcome.Succeeded);
        SceneRect bar = Assert.IsType<SceneRect>(outcome.Result!.Primitive);
        Assert.Equal(40, bar.Bounds.Left, 10);
        Assert.Equal(90, bar.Bounds.Right, 10);
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
                8,
                SlotCentreY
            ),
            _scale);

        Assert.False(outcome.Succeeded);
        Assert.Equal(CriticalOverlayRefusal.NotACriticalInterval, outcome.Refusal);
    }

    [Fact]
    public void There_is_no_second_source_of_truth_for_the_span()
    {
        // The old builder took the parent's visible span TWICE — once in the request
        // and once in a dictionary — and a test proved the request won. With no
        // parent in the request there is no second source to disagree with, so the
        // only inputs are the interval's own dates and its slot.
        CriticalOverlayCreationOutcome outcome = Build(5, 9, slotCentreY: 100);

        Assert.True(outcome.Succeeded);
        SceneRect bar = Assert.IsType<SceneRect>(outcome.Result!.Primitive);
        // Same dates as the default build; only the slot moved, so only Y moved.
        Assert.Equal(40, bar.Bounds.Left, 10);
        Assert.Equal(90, bar.Bounds.Right, 10);
        Assert.Equal(98, bar.Bounds.Top, 10);
    }

    [Fact]
    public void Refuses_a_non_positive_predetermined_height()
    {
        CriticalOverlayCreationOutcome outcome = Build(5, 9, predeterminedHeightPt: 0);

        Assert.False(outcome.Succeeded);
        Assert.Equal(CriticalOverlayRefusal.InvalidGeometry, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_null_request()
    {
        CriticalOverlayCreationOutcome outcome = CriticalOverlayBuilder.TryBuild(null, _scale);

        Assert.False(outcome.Succeeded);
        Assert.Equal(CriticalOverlayRefusal.NullRequest, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_null_time_scale()
    {
        CriticalOverlayCreationOutcome outcome = CriticalOverlayBuilder.TryBuild(
            new CriticalOverlayRequest(Child(5, 9), _style, 8, SlotCentreY),
            null
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
        double predeterminedHeightPt = 8,
        double slotCentreY = SlotCentreY
    ) =>
        CriticalOverlayBuilder.TryBuild(
            new CriticalOverlayRequest(
                Child(startDay, finishDay),
                _style,
                predeterminedHeightPt,
                slotCentreY
            ),
            _scale);

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

    private static GanttEvent ChildOn(DateOnly start, DateOnly finish) =>
        new(
            2,
            _childId,
            null,
            null,
            GanttEntityType.CriticalInterval,
            "Critical",
            start,
            finish,
            _parentId,
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
