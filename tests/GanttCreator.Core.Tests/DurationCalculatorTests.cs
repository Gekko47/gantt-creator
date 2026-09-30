namespace GanttCreator.Core.Tests;

/// <summary>
/// R4.7F: the <c>Duration</c> column is calculated in Core from date-only
/// semantics, inclusive for spans, a marker for point entities, blank for
/// structural rows, and never a value for invalid dates.
/// </summary>
public sealed class DurationCalculatorTests
{
    private static DateOnly Day(int day) => new(2026, 9, day);

    private static GanttEvent Event(
        int row,
        GanttEntityType type,
        DateOnly? start,
        DateOnly? finish = null,
        GanttRowId? parentId = null) =>
        new(
            row,
            GanttRowId.New(),
            GanttRowId.New(),
            0,
            type,
            $"Row {row}",
            start,
            finish,
            parentId,
            null,
            null,
            null,
            null,
            true,
            null);

    private static Dictionary<int, string?> Current(params (int Row, string? Value)[] cells)
    {
        Dictionary<int, string?> map = [];
        foreach ((int row, string? value) in cells)
        {
            map[row] = value;
        }

        return map;
    }

    [Theory]
    [InlineData(1, 1)]
    [InlineData(1, 5)]
    [InlineData(5, 9)]
    public void An_inclusive_span_yields_the_correct_day_count(int startDay, int finishDay)
    {
        Assert.True(DurationCalculator.TryCalculate(
            2,
            GanttEntityType.AsPlannedActivity,
            Day(startDay),
            Day(finishDay),
            out string? text,
            out _));

        Assert.Equal((finishDay - startDay + 1).ToString(System.Globalization.CultureInfo.InvariantCulture), text);
    }

    [Fact]
    public void A_one_day_span_is_one_and_not_zero()
    {
        // The off-by-one the row exists to prevent: an exclusive count would report 0
        // for a real single-day activity.
        Assert.True(DurationCalculator.TryCalculate(
            2,
            GanttEntityType.AsPlannedActivity,
            Day(3),
            Day(3),
            out string? text,
            out _));

        Assert.Equal("1", text);
    }

    [Fact]
    public void A_span_crossing_a_month_boundary_counts_across_it()
    {
        Assert.True(DurationCalculator.TryCalculate(
            2,
            GanttEntityType.AsPlannedActivity,
            new DateOnly(2026, 9, 28),
            new DateOnly(2026, 10, 3),
            out string? text,
            out _));

        // Sep 28, 29, 30, Oct 1, 2, 3.
        Assert.Equal("6", text);
    }

    [Fact]
    public void A_span_crossing_a_year_boundary_counts_across_it()
    {
        Assert.True(DurationCalculator.TryCalculate(
            2,
            GanttEntityType.AsPlannedActivity,
            new DateOnly(2026, 12, 30),
            new DateOnly(2027, 1, 2),
            out string? text,
            out _));

        Assert.Equal("4", text);
    }

    [Fact]
    public void A_leap_day_span_counts_the_leap_day()
    {
        Assert.True(DurationCalculator.TryInclusiveDays(
            new DateOnly(2024, 2, 28),
            new DateOnly(2024, 3, 1),
            out var days));

        Assert.Equal(3, days);
    }

    [Theory]
    [InlineData(GanttEntityType.AsPlannedMilestone)]
    [InlineData(GanttEntityType.AsBuiltMilestone)]
    [InlineData(GanttEntityType.BaselineMilestone)]
    [InlineData(GanttEntityType.CriticalMilestone)]
    [InlineData(GanttEntityType.Delineator)]
    public void A_point_entity_shows_the_marker_not_a_count(GanttEntityType type)
    {
        Assert.True(DurationCalculator.TryCalculate(2, type, Day(5), finish: null, out string? text, out _));

        Assert.Equal(DurationCalculator.NonDurationMarker, text);
    }

    [Fact]
    public void A_point_entity_with_a_stray_finish_ignores_it()
    {
        // The milestone/delineator contract is that Finish is never read for
        // geometry. A user who typed one anyway must not turn a milestone into a span.
        Assert.True(DurationCalculator.TryCalculate(
            2,
            GanttEntityType.AsPlannedMilestone,
            Day(5),
            Day(12),
            out string? text,
            out _));

        Assert.Equal(DurationCalculator.NonDurationMarker, text);
    }

    [Theory]
    [InlineData(GanttEntityType.Splitter)]
    [InlineData(GanttEntityType.Spacer)]
    public void A_structural_row_is_blank_not_a_marker(GanttEntityType type)
    {
        // A Splitter or Spacer states nothing about time, so it carries no marker
        // either -- a "-" would imply a date it does not have.
        Assert.True(DurationCalculator.TryCalculate(2, type, start: null, finish: null, out string? text, out _));

        Assert.Equal(string.Empty, text);
    }

    [Fact]
    public void A_critical_interval_computes_from_its_own_dates_not_its_parents()
    {
        // The parent's span is 10 days; the interval's own Sep 6-9 is 4. Reading the
        // parent's dates here is the defect ADR-0027's parent-independence removed.
        Assert.True(DurationCalculator.TryCalculate(
            5,
            GanttEntityType.CriticalInterval,
            Day(6),
            Day(9),
            out string? text,
            out _));

        Assert.Equal("4", text);
    }

    [Fact]
    public void A_date_time_input_normalises_to_its_date_component()
    {
        // D1. A workbook carrying date-times must not produce an off-by-one, because
        // the calculator is DateOnly-based and the time component never reaches it.
        DateOnly fromDateTime = DateOnly.FromDateTime(new DateTime(2026, 9, 5, 17, 45, 0));
        DateOnly fromDate = new(2026, 9, 5);

        Assert.True(DurationCalculator.TryInclusiveDays(fromDateTime, Day(9), out var withTime));
        Assert.True(DurationCalculator.TryInclusiveDays(fromDate, Day(9), out var withoutTime));

        Assert.Equal(withoutTime, withTime);
        Assert.Equal(5, withTime);
    }

    [Fact]
    public void A_missing_start_produces_no_value_and_a_typed_refusal()
    {
        Assert.False(DurationCalculator.TryCalculate(
            2,
            GanttEntityType.AsPlannedActivity,
            start: null,
            finish: Day(5),
            out string? text,
            out DurationRefusalEntry? refusal));

        Assert.Null(text);
        Assert.NotNull(refusal);
        Assert.Equal(DurationRefusal.StartMissing, refusal.Reason);
        Assert.Equal(2, refusal.RowNumber);
    }

    [Fact]
    public void A_missing_finish_produces_no_value_and_a_typed_refusal()
    {
        Assert.False(DurationCalculator.TryCalculate(
            2,
            GanttEntityType.AsPlannedActivity,
            Day(5),
            finish: null,
            out string? text,
            out DurationRefusalEntry? refusal));

        Assert.Null(text);
        Assert.NotNull(refusal);
        Assert.Equal(DurationRefusal.FinishMissing, refusal.Reason);
    }

    [Fact]
    public void A_finish_before_start_produces_no_value_not_a_negative_count()
    {
        Assert.False(DurationCalculator.TryCalculate(
            2,
            GanttEntityType.AsPlannedActivity,
            Day(9),
            Day(5),
            out string? text,
            out DurationRefusalEntry? refusal));

        // Never "0", never "-4": a wrong number reads as a real duration.
        Assert.Null(text);
        Assert.NotNull(refusal);
        Assert.Equal(DurationRefusal.FinishBeforeStart, refusal.Reason);
    }

    [Fact]
    public void A_milestone_with_no_start_refuses_rather_than_showing_the_marker()
    {
        // The marker asserts "this row has a date but no span". Without a date there
        // is nothing to assert, so it is a refusal.
        Assert.False(DurationCalculator.TryCalculate(
            2,
            GanttEntityType.AsPlannedMilestone,
            start: null,
            finish: null,
            out string? text,
            out DurationRefusalEntry? refusal));

        Assert.Null(text);
        Assert.Equal(DurationRefusal.StartMissing, refusal!.Reason);
    }

    [Fact]
    public void An_uncatalogued_type_is_refused_rather_than_defaulted()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => DurationCalculator.TryCalculate(
            2,
            (GanttEntityType)999,
            Day(5),
            Day(9),
            out _,
            out _));
    }

    [Fact]
    public void The_plan_writes_only_the_rows_whose_value_changes()
    {
        List<GanttEvent> events =
        [
            Event(2, GanttEntityType.AsPlannedActivity, Day(1), Day(5)),
            Event(3, GanttEntityType.AsPlannedActivity, Day(1), Day(10)),
        ];

        DurationWritePlan plan = DurationCalculator.Plan(events, Current((2, "5"), (3, "99")));

        DurationWrite only = Assert.Single(plan.Writes);
        Assert.Equal(3, only.RowNumber);
        Assert.Equal("10", only.Text);
    }

    [Fact]
    public void An_already_correct_sheet_produces_an_empty_write_set()
    {
        List<GanttEvent> events = [Event(2, GanttEntityType.AsPlannedActivity, Day(1), Day(5))];

        DurationWritePlan plan = DurationCalculator.Plan(events, Current((2, "5")));

        Assert.Empty(plan.Writes);
        Assert.Equal(0, plan.WriteCount);
    }

    [Fact]
    public void A_refused_row_clears_a_stale_value_from_a_previous_refresh()
    {
        // D4: leaving yesterday's number on a row whose dates are now invalid is
        // exactly the misleading value the row forbids.
        List<GanttEvent> events = [Event(2, GanttEntityType.AsPlannedActivity, Day(9), Day(5))];

        DurationWritePlan plan = DurationCalculator.Plan(events, Current((2, "12")));

        DurationWrite cleared = Assert.Single(plan.Writes);
        Assert.Equal(2, cleared.RowNumber);
        Assert.Equal(string.Empty, cleared.Text);
        Assert.Equal(DurationRefusal.FinishBeforeStart, Assert.Single(plan.Refusals).Reason);
    }

    [Fact]
    public void A_refused_row_that_is_already_blank_writes_nothing()
    {
        List<GanttEvent> events = [Event(2, GanttEntityType.AsPlannedActivity, Day(9), Day(5))];

        DurationWritePlan plan = DurationCalculator.Plan(events, Current((2, null)));

        Assert.Empty(plan.Writes);
        Assert.Single(plan.Refusals);
    }

    [Fact]
    public void The_plan_is_ordered_by_row_and_independent_of_input_order()
    {
        List<GanttEvent> events =
        [
            Event(5, GanttEntityType.AsPlannedActivity, Day(1), Day(5)),
            Event(2, GanttEntityType.AsPlannedActivity, Day(1), Day(5)),
            Event(3, GanttEntityType.AsPlannedActivity, Day(1), Day(5)),
        ];

        DurationWritePlan shuffled = DurationCalculator.Plan(events, Current());

        Assert.Equal([2, 3, 5], shuffled.Writes.Select(write => write.RowNumber));
    }

    [Fact]
    public void Plan_refuses_null_inputs_rather_than_writing_every_row()
    {
        List<GanttEvent> events = [Event(2, GanttEntityType.AsPlannedActivity, Day(1), Day(5))];

        Assert.Throws<ArgumentNullException>(() => DurationCalculator.Plan(events, null!));
        Assert.Throws<ArgumentNullException>(() => DurationCalculator.Plan(null!, Current()));
    }
}