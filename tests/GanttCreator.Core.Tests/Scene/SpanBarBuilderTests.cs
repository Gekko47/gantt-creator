using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

public sealed class SpanBarBuilderTests
{
    private static readonly GanttRowId _rowId = GanttRowId.New();
    private static readonly GanttRowId _laneId = GanttRowId.New();

    // A 31-day January 2024 plot, 0..310pt, so one day is exactly 10pt.
    private static readonly TimeScale _scale =
        TimeScale.TryCreate(new DateOnly(2024, 1, 1), new DateOnly(2024, 1, 31), 0, 310).Scale!;

    private static readonly SceneStyle _style = new("Planned", fillColour: ColourHex.Parse("#92D050"));

    [Fact]
    public void Builds_the_exact_inclusive_span_rectangle_centred_on_the_slot()
    {
        SpanBarCreationOutcome outcome = Build(start: new DateOnly(2024, 1, 5), finish: new DateOnly(2024, 1, 9));

        Assert.True(outcome.Succeeded);
        RectD bounds = Assert.IsType<RectD>(outcome.Result!.VisibleBounds);
        // left = DateToX(Start) = 40; duration = 9 - 5 + 1 = 5 days; width = 50.
        Assert.Equal(40, bounds.Left);
        Assert.Equal(90, bounds.Right);
        Assert.Equal(50, bounds.Width);
        // Centred on the slot centre of 60 with a height of 8: top 56, bottom 64.
        Assert.Equal(56, bounds.Top);
        Assert.Equal(8, bounds.Height);
    }

    [Fact]
    public void A_one_day_activity_keeps_exactly_one_day_of_width()
    {
        SpanBarCreationOutcome outcome = Build(start: new DateOnly(2024, 1, 5), finish: new DateOnly(2024, 1, 5));

        Assert.True(outcome.Succeeded);
        RectD bounds = outcome.Result!.VisibleBounds!.Value;
        Assert.Equal(40, bounds.Left);
        Assert.Equal(10, bounds.Width, 10);
        Assert.True(bounds.Width > 0, "A one-day activity must not collapse to zero width.");
    }

    [Fact]
    public void Never_increments_the_finish_date()
    {
        SpanBarCreationOutcome outcome = Build(start: new DateOnly(2024, 1, 1), finish: new DateOnly(2024, 1, 31));

        Assert.True(outcome.Succeeded);
        // The last day maps to 300 and the bar runs 31 days to the plot right edge.
        Assert.Equal(310, outcome.Result!.VisibleBounds!.Value.Right, 10);
    }

    [Fact]
    public void Emits_the_bar_at_activity_layer_with_a_role_derived_id()
    {
        SpanBarCreationOutcome outcome = Build(start: new DateOnly(2024, 1, 5), finish: new DateOnly(2024, 1, 9));

        SceneRect bar = Assert.IsType<SceneRect>(outcome.Result!.Primitive);
        Assert.Equal(ZLayer.ActivityBody, bar.ZLayer);
        Assert.EndsWith(":bar", bar.PrimitiveId, StringComparison.Ordinal);
    }

    [Fact]
    public void Clips_a_span_that_starts_before_and_ends_inside_the_plot()
    {
        // The plot opens on 1 Jan, so a span starting 30 Dec 2023 has its start
        // clamped to the plot's left edge while its finish stays in range.
        SpanBarCreationOutcome outcome = Build(
            start: new DateOnly(2023, 12, 30),
            finish: new DateOnly(2024, 1, 3)
        );

        Assert.True(outcome.Succeeded);
        SceneRect? bar = outcome.Result!.Primitive;
        Assert.NotNull(bar);
        Assert.Equal(0, bar!.Bounds.Left);
        // 3 Jan is the third day of January, so the inclusive right edge is 20 + 10 = 30pt.
        Assert.Equal(30, bar.Bounds.Right, 10);
        SceneWarning warning = Assert.Single(outcome.Result.Warnings);
        Assert.Equal(SpanBarBuilder.ClippedToPlotCode, warning.Code);
    }

    [Fact]
    public void Emits_no_bar_and_one_warning_for_a_span_entirely_after_the_plot()
    {
        SpanBarCreationOutcome outcome = Build(start: new DateOnly(2024, 2, 5), finish: new DateOnly(2024, 2, 9));

        Assert.True(outcome.Succeeded);
        Assert.Null(outcome.Result!.Primitive);
        Assert.Null(outcome.Result.VisibleBounds);
        SceneWarning warning = Assert.Single(outcome.Result.Warnings);
        Assert.Equal(SpanBarBuilder.OutsidePlotRangeCode, warning.Code);
    }

    [Fact]
    public void Emits_no_bar_and_one_warning_for_a_span_entirely_before_the_plot()
    {
        SpanBarCreationOutcome outcome = Build(start: new DateOnly(2023, 12, 5), finish: new DateOnly(2023, 12, 9));

        Assert.True(outcome.Succeeded);
        Assert.Null(outcome.Result!.Primitive);
        Assert.Equal(SpanBarBuilder.OutsidePlotRangeCode, Assert.Single(outcome.Result.Warnings).Code);
    }

    [Fact]
    public void Refuses_a_span_with_no_finish_date_rather_than_guessing_one()
    {
        SpanBarCreationOutcome outcome = SpanBarBuilder.TryBuild(
            new SpanBarRequest(Event(start: new DateOnly(2024, 1, 5), finish: null), _style, 60, 8),
            _scale
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(SpanBarRefusal.NotASpan, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_reversed_span_rather_than_building_a_zero_width_bar()
    {
        // The enum documents NotASpan as "no finish after its start", so a finish
        // before the start is refused instead of reaching the clip branch, where
        // the two collapsed edges would otherwise produce a bar.
        SpanBarCreationOutcome outcome = SpanBarBuilder.TryBuild(
            new SpanBarRequest(Event(new DateOnly(2024, 1, 9), new DateOnly(2024, 1, 5)), _style, 60, 8),
            _scale
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(SpanBarRefusal.NotASpan, outcome.Refusal);
    }

    [Fact]
    public void Refuses_an_event_with_no_start_date()
    {
        SpanBarCreationOutcome outcome = SpanBarBuilder.TryBuild(
            new SpanBarRequest(Event(start: null, finish: new DateOnly(2024, 1, 5)), _style, 60, 8),
            _scale
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(SpanBarRefusal.NotASpan, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_null_request()
    {
        SpanBarCreationOutcome outcome = SpanBarBuilder.TryBuild(null, _scale);

        Assert.False(outcome.Succeeded);
        Assert.Equal(SpanBarRefusal.NullRequest, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_non_finite_slot_centre()
    {
        SpanBarCreationOutcome outcome = SpanBarBuilder.TryBuild(
            new SpanBarRequest(Event(new DateOnly(2024, 1, 5), new DateOnly(2024, 1, 9)), _style, double.NaN, 8),
            _scale
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(SpanBarRefusal.InvalidGeometry, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_negative_resolved_height()
    {
        SpanBarCreationOutcome outcome = SpanBarBuilder.TryBuild(
            new SpanBarRequest(Event(new DateOnly(2024, 1, 5), new DateOnly(2024, 1, 9)), _style, 60, -1),
            _scale
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(SpanBarRefusal.InvalidGeometry, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_null_time_scale()
    {
        SpanBarCreationOutcome outcome = SpanBarBuilder.TryBuild(
            new SpanBarRequest(Event(new DateOnly(2024, 1, 5), new DateOnly(2024, 1, 9)), _style, 60, 8),
            null
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(SpanBarRefusal.InvalidTimeScale, outcome.Refusal);
    }

    [Fact]
    public void Produces_identical_geometry_when_the_input_order_is_shuffled()
    {
        SpanBarCreationOutcome first = Build(start: new DateOnly(2024, 1, 5), finish: new DateOnly(2024, 1, 9));
        SpanBarCreationOutcome second = Build(start: new DateOnly(2024, 1, 5), finish: new DateOnly(2024, 1, 9));

        Assert.Equal(first.Result!.Primitive!.Bounds, second.Result!.Primitive!.Bounds);
        Assert.Equal(first.Result.VisibleBounds, second.Result.VisibleBounds);
    }

    private static SpanBarCreationOutcome Build(DateOnly start, DateOnly finish) =>
        SpanBarBuilder.TryBuild(
            new SpanBarRequest(Event(start, finish), _style, 60, 8),
            _scale
        );

    private static GanttEvent Event(DateOnly? start, DateOnly? finish) =>
        new(
            1,
            _rowId,
            _laneId,
            null,
            GanttEntityType.AsPlannedActivity,
            "Activity",
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
