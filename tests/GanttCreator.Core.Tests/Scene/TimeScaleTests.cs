using System.Globalization;
using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

public sealed class TimeScaleTests
{
    private static readonly DateOnly _plotStart = new(2024, 1, 1);
    private static readonly DateOnly _plotFinish = new(2024, 1, 5);

    [Fact]
    public void TryCreate_builds_a_valid_scale()
    {
        TimeScaleCreationOutcome outcome = Create();

        Assert.True(outcome.Succeeded);
        Assert.Equal(_plotStart, outcome.Scale!.PlotStart);
        Assert.Equal(_plotFinish, outcome.Scale.PlotFinish);
        Assert.Equal(5, outcome.Scale.TotalPlotDays);
        Assert.Equal(20, outcome.Scale.DayWidth);
    }

    [Fact]
    public void TryCreate_accepts_a_one_day_plot()
    {
        TimeScaleCreationOutcome outcome = TimeScale.TryCreate(new DateOnly(2024, 1, 1), new DateOnly(2024, 1, 1), 10, 30);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.Scale!.TotalPlotDays);
        Assert.Equal(20, outcome.Scale.DayWidth);
    }

    [Fact]
    public void TryCreate_rejects_start_after_finish()
    {
        TimeScaleCreationOutcome outcome = TimeScale.TryCreate(new DateOnly(2024, 1, 2), new DateOnly(2024, 1, 1), 0, 100);

        Assert.False(outcome.Succeeded);
        Assert.Equal(TimeScaleCreationRefusal.StartAfterFinish, outcome.Refusal);
    }

    [Fact]
    public void TryCreate_rejects_default_dates()
    {
        TimeScaleCreationOutcome defaultStart = TimeScale.TryCreate(default, new DateOnly(2024, 1, 1), 0, 100);
        TimeScaleCreationOutcome defaultFinish = TimeScale.TryCreate(new DateOnly(2024, 1, 1), default, 0, 100);

        Assert.Equal(TimeScaleCreationRefusal.DefaultDates, defaultStart.Refusal);
        Assert.Equal(TimeScaleCreationRefusal.DefaultDates, defaultFinish.Refusal);
    }

    [Theory]
    [InlineData(double.NaN, 100)]
    [InlineData(double.PositiveInfinity, 100)]
    [InlineData(double.NegativeInfinity, 100)]
    [InlineData(0, double.NaN)]
    [InlineData(0, double.PositiveInfinity)]
    [InlineData(100, 100)]
    [InlineData(100, 99)]
    public void TryCreate_rejects_invalid_plot_geometry(double left, double right)
    {
        TimeScaleCreationOutcome outcome = TimeScale.TryCreate(_plotStart, _plotFinish, left, right);

        Assert.False(outcome.Succeeded);
        Assert.Equal(TimeScaleCreationRefusal.NonFiniteOrDegenerateWidth, outcome.Refusal);
    }

    [Fact]
    public void DateToX_maps_start_of_day_and_boundaries()
    {
        TimeScale scale = Create().Scale!;

        Assert.Equal(0, scale.DateToX(_plotStart));
        Assert.Equal(20, scale.DateToX(new DateOnly(2024, 1, 2)));
        Assert.Equal(80, scale.DateToX(_plotFinish));
    }

    [Fact]
    public void TryDateToX_rejects_dates_outside_the_inclusive_range()
    {
        TimeScale scale = Create().Scale!;

        Assert.False(scale.TryDateToX(new DateOnly(2023, 12, 31), out _));
        Assert.False(scale.TryDateToX(new DateOnly(2024, 1, 6), out _));
        _ = Assert.Throws<ArgumentOutOfRangeException>(() => scale.DateToX(new DateOnly(2023, 12, 31)));
    }

    [Fact]
    public void DateToX_is_culture_invariant()
    {
        TimeScale scale = Create().Scale!;
        CultureInfo original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("tr-TR");
            Assert.Equal(20, scale.DateToX(new DateOnly(2024, 1, 2)));
            Assert.Equal(2, scale.DurationDays(new DateOnly(2024, 1, 2), new DateOnly(2024, 1, 3)));
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Fact]
    public void DurationDays_is_inclusive_and_does_not_mutate_finish()
    {
        TimeScale scale = Create().Scale!;
        DateOnly start = new(2024, 1, 2);
        DateOnly finish = new(2024, 1, 4);

        Assert.Equal(1, scale.DurationDays(start, start));
        Assert.Equal(3, scale.DurationDays(start, finish));
        Assert.Equal(new DateOnly(2024, 1, 4), finish);
    }

    [Fact]
    public void ActivityRightX_uses_fixed_start_and_duration_width()
    {
        TimeScale scale = Create().Scale!;
        DateOnly start = new(2024, 1, 2);
        DateOnly finish = new(2024, 1, 4);

        Assert.Equal(20, scale.DateToX(start));
        Assert.Equal(80, scale.ActivityRightX(start, finish));
        Assert.Equal(20, scale.ActivityRightX(start, start) - scale.DateToX(start));
        Assert.Equal(20, scale.DateToX(start));
    }

    [Fact]
    public void TryDurationDays_rejects_invalid_activity_dates()
    {
        TimeScale scale = Create().Scale!;

        Assert.False(scale.TryDurationDays(new DateOnly(2023, 12, 31), _plotStart, out _));
        Assert.False(scale.TryDurationDays(_plotStart, new DateOnly(2024, 1, 6), out _));
        Assert.False(scale.TryDurationDays(new DateOnly(2024, 1, 3), new DateOnly(2024, 1, 2), out _));
        _ = Assert.Throws<ArgumentOutOfRangeException>(() => scale.DurationDays(new DateOnly(2024, 1, 3), new DateOnly(2024, 1, 2)));
    }

    [Fact]
    public void PointEventX_uses_only_the_date_x()
    {
        TimeScale scale = Create().Scale!;
        Assert.Equal(20, scale.PointEventX(new DateOnly(2024, 1, 2)));
    }

    [Fact]
    public void DateToX_is_monotonic_for_shuffled_reversed_and_duplicate_dates()
    {
        TimeScale scale = Create().Scale!;
        DateOnly[] dates =
        [
            new(2024, 1, 4),
            new(2024, 1, 2),
            new(2024, 1, 4),
            new(2024, 1, 1),
            new(2024, 1, 3),
            new(2024, 1, 5),
            new(2024, 1, 2),
        ];

        double[] xs = [.. dates.Order().Select(scale.DateToX)];
        for (var index = 1; index < xs.Length; index++)
        {
            Assert.True(xs[index] >= xs[index - 1]);
        }
    }

    [Fact]
    public void DateToX_handles_leap_day_and_year_boundary()
    {
        TimeScale leap = TimeScale.TryCreate(new DateOnly(2024, 2, 28), new DateOnly(2024, 3, 1), 0, 40).Scale!;
        TimeScale year = TimeScale.TryCreate(new DateOnly(2023, 12, 31), new DateOnly(2024, 1, 1), 0, 20).Scale!;

        Assert.Equal(40.0 / 3, leap.DateToX(new DateOnly(2024, 2, 29)));
        Assert.Equal(1, leap.DurationDays(new DateOnly(2024, 2, 29), new DateOnly(2024, 2, 29)));
        Assert.Equal(10, year.DateToX(new DateOnly(2024, 1, 1)));
        Assert.Equal(2, year.DurationDays(new DateOnly(2023, 12, 31), new DateOnly(2024, 1, 1)));
    }

    private static TimeScaleCreationOutcome Create() =>
        TimeScale.TryCreate(_plotStart, _plotFinish, 0, 100);
}
