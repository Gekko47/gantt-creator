namespace GanttCreator.Core.Tests;

/// <summary>
/// R5.1 D1/D2: the plot-range extent resolves per end from its own mode, and
/// every bad input is a typed refusal rather than a silent fall back.
/// </summary>
public sealed class PlotRangeResolverTests
{
    private static GanttEvent Event(
        int row,
        GanttEntityType type = GanttEntityType.AsPlannedActivity,
        DateOnly? start = null,
        DateOnly? finish = null) =>
        new(
            row,
            GanttRowId.New(),
            GanttRowId.New(),
            0,
            type,
            $"Row {row}",
            start ?? new DateOnly(2024, 1, 10),
            finish ?? (type is GanttEntityType.AsPlannedMilestone ? null : new DateOnly(2024, 1, 20)),
            null,
            "AsPlannedActivity",
            null,
            null,
            null,
            true,
            null);

    [Fact]
    public void Automatic_mode_derives_the_padded_extent_from_visible_events()
    {
        PlotRangeOutcome outcome = PlotRangeResolver.TryResolve(
            [Event(1), Event(2, start: new DateOnly(2024, 2, 5), finish: new DateOnly(2024, 2, 15))],
            "DataRange",
            "DataRange",
            null,
            null,
            paddingDays: 3);

        Assert.True(outcome.Succeeded);
        Assert.Equal(new DateOnly(2024, 1, 7), outcome.Start);
        Assert.Equal(new DateOnly(2024, 2, 18), outcome.Finish);
    }

    [Fact]
    public void Automatic_mode_counts_a_milestone_start_but_ignores_its_finish()
    {
        var milestone = Event(1, GanttEntityType.AsPlannedMilestone, new DateOnly(2024, 3, 10), new DateOnly(2024, 9, 1));

        PlotRangeOutcome outcome = PlotRangeResolver.TryResolve(
            [milestone],
            "DataRange",
            "DataRange",
            null,
            null,
            paddingDays: 0);

        // The 1 Sept finish must not widen the range: milestones read Start
        // as their single date, so the extent is the point date itself.
        Assert.True(outcome.Succeeded);
        Assert.Equal(new DateOnly(2024, 3, 10), outcome.Start);
        Assert.Equal(new DateOnly(2024, 3, 10), outcome.Finish);
    }

    [Fact]
    public void Automatic_mode_with_no_start_date_refuses_rather_than_inventing_a_range()
    {
        var dateless = Event(1, GanttEntityType.Splitter, start: null, finish: null) with
        {
            Start = null,
            Finish = null,
        };

        PlotRangeOutcome outcome = PlotRangeResolver.TryResolve(
            [dateless],
            "DataRange",
            "DataRange",
            null,
            null,
            paddingDays: 3);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PlotRangeRefusal.NoPlotRange, outcome.Refusal);
    }

    [Fact]
    public void Null_events_are_refused()
    {
        PlotRangeOutcome outcome = PlotRangeResolver.TryResolve(
            null, "DataRange", "DataRange", null, null, paddingDays: 3);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PlotRangeRefusal.NullEvents, outcome.Refusal);
    }

    [Theory]
    [InlineData("Automatic", "DataRange")]
    [InlineData("DataRange", "automatic")]
    [InlineData("", "DataRange")]
    [InlineData("DataRange", null)]
    public void Unknown_mode_text_is_refused(string startMode, string? finishMode)
    {
        PlotRangeOutcome outcome = PlotRangeResolver.TryResolve(
            [Event(1)], startMode, finishMode, null, null, paddingDays: 3);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PlotRangeRefusal.UnknownMode, outcome.Refusal);
    }

    [Fact]
    public void Explicit_mode_uses_the_supplied_dates_verbatim()
    {
        PlotRangeOutcome outcome = PlotRangeResolver.TryResolve(
            [Event(1)],
            "Explicit",
            "Explicit",
            "2024-02-01",
            "2024-04-30",
            paddingDays: 99);

        // Padding must not touch an explicit end: the 99-day pad applies to
        // automatic derivation only, so the dates come back exactly.
        Assert.True(outcome.Succeeded);
        Assert.Equal(new DateOnly(2024, 2, 1), outcome.Start);
        Assert.Equal(new DateOnly(2024, 4, 30), outcome.Finish);
    }

    [Fact]
    public void Explicit_mode_accepts_a_single_day_range()
    {
        PlotRangeOutcome outcome = PlotRangeResolver.TryResolve(
            [Event(1)],
            "Explicit",
            "Explicit",
            "2024-05-06",
            "2024-05-06",
            paddingDays: 3);

        Assert.True(outcome.Succeeded);
        Assert.Equal(new DateOnly(2024, 5, 6), outcome.Start);
        Assert.Equal(new DateOnly(2024, 5, 6), outcome.Finish);
    }

    [Fact]
    public void Mixed_modes_resolve_each_end_from_its_own_source()
    {
        PlotRangeOutcome outcome = PlotRangeResolver.TryResolve(
            [Event(1, start: new DateOnly(2024, 1, 10), finish: new DateOnly(2024, 1, 20))],
            "Explicit",
            "DataRange",
            "2024-01-01",
            null,
            paddingDays: 0);

        Assert.True(outcome.Succeeded);
        Assert.Equal(new DateOnly(2024, 1, 1), outcome.Start);
        Assert.Equal(new DateOnly(2024, 1, 20), outcome.Finish);
    }

    [Fact]
    public void Explicit_start_after_explicit_finish_is_refused()
    {
        PlotRangeOutcome outcome = PlotRangeResolver.TryResolve(
            [Event(1)],
            "Explicit",
            "Explicit",
            "2024-03-02",
            "2024-03-01",
            paddingDays: 3);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PlotRangeRefusal.StartAfterFinish, outcome.Refusal);
    }

    [Fact]
    public void Explicit_mode_with_a_missing_date_is_refused()
    {
        PlotRangeOutcome outcome = PlotRangeResolver.TryResolve(
            [Event(1)], "Explicit", "DataRange", "   ", null, paddingDays: 3);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PlotRangeRefusal.MissingExplicitDate, outcome.Refusal);
    }

    [Fact]
    public void Explicit_mode_with_an_unparsable_date_is_refused()
    {
        PlotRangeOutcome outcome = PlotRangeResolver.TryResolve(
            [Event(1)], "DataRange", "Explicit", null, "not-a-date", paddingDays: 3);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PlotRangeRefusal.UnparsableExplicitDate, outcome.Refusal);
    }

    [Theory]
    [InlineData("DataRange", "datarange")]
    [InlineData("Explicit", "explicit")]
    public void Mode_parsing_matches_ordinal_with_no_case_folding(string text, string lowered)
    {
        Assert.True(PlotRangeModes.TryParse(text, out _));
        Assert.False(PlotRangeModes.TryParse(lowered, out _));
    }

    [Theory]
    [InlineData("Automatic")]
    [InlineData("")]
    [InlineData(null)]
    public void Mode_parsing_rejects_unknown_text(string? text)
    {
        Assert.False(PlotRangeModes.TryParse(text, out _));
    }
}
