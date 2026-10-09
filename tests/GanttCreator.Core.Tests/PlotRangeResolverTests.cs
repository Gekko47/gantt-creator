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
            "01/02/2024",
            "30/04/2024",
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
            "06/05/2024",
            "06/05/2024",
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
            "01/01/2024",
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
            "02/03/2024",
            "01/03/2024",
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

    // --- TryParsePlotDate (the ribbon edit-box parser, R5.1) ---
    // Exact dates only: the default display format dd/MM/yyyy, plus the bare
    // Excel serial the edit box echoes back. The activity multi-format
    // equivalents are refused here (see the parser doc block).

    [Theory]

    // The chart's own dd/MM/yyyy contract, exact.
    [InlineData("15/03/2026", 2026, 3, 15)]
    [InlineData("  15/03/2026  ", 2026, 3, 15)]
    public void TryParsePlotDate_accepts_the_default_display_format(string text, int year, int month, int day)
    {
        Assert.True(PlotRangeResolver.TryParsePlotDate(text, out DateOnly parsed));
        Assert.Equal(new DateOnly(year, month, day), parsed);
    }

    [Theory]
    // The activity auto-parser equivalents are NOT plot input: an ambiguous
    // entry must fail here rather than silently mean a different date than
    // the chart's dd/MM/yyyy contract. The serial path stays (it is the
    // echo, not a format), as do the refusal positives below.
    [InlineData("2026-03-15")]
    [InlineData("15-Mar-2026")]
    [InlineData("15/Mar/2026")]
    [InlineData("15 Mar 2026")]
    [InlineData("3/15/2026")]
    [InlineData("03/15/2026")]
    public void TryParsePlotDate_refuses_activity_auto_parser_equivalents(string text)
    {
        Assert.False(PlotRangeResolver.TryParsePlotDate(text, out DateOnly parsed));
        Assert.Equal(default, parsed);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("not-a-date")]
    [InlineData("15/13/2026")]
    [InlineData("32/01/2026")]
    [InlineData("15/02/2026 extra")]
    public void TryParsePlotDate_refuses_non_dates_rather_than_defaulting(string? text)
    {
        // Positive test for the failure path: every non-date returns false
        // with a default DateOnly, so the caller reverts instead of
        // persisting a wrong value.
        Assert.False(PlotRangeResolver.TryParsePlotDate(text, out DateOnly parsed));
        Assert.Equal(default, parsed);
    }

    [Fact]
    public void TryParsePlotDate_refuses_the_default_date_value()
    {
        // 01/001 parses as DateOnly.MinValue, which is the DefaultDates
        // refusal everywhere else; the parser must not bless it.
        Assert.False(PlotRangeResolver.TryParsePlotDate("01/01/0001", out _));
    }

    [Theory]
    // The serial round-trip: the ribbon edit box echoes a typed date back as
    // its Excel number (e.g. 46118 for 06/04/2026), and the previous
    // text-only parser rejected it, so a typed date could never round-trip.
    // 06/04/2026 is 46118; the .0 double render must parse identically.
    [InlineData("46118", 2026, 4, 6)]
    [InlineData("46118.0", 2026, 4, 6)]
    [InlineData("  46118  ", 2026, 4, 6)]
    public void TryParsePlotDate_accepts_a_bare_Excel_serial(string text, int year, int month, int day)
    {
        Assert.True(PlotRangeResolver.TryParsePlotDate(text, out DateOnly parsed));
        Assert.Equal(new DateOnly(year, month, day), parsed);
    }

    [Theory]
    // Serial bounds follow the shared 1900-epoch converter: 0 and negatives
    // are below MinSerial, 2958466 is above MaxSerial, and a fractional serial
    // truncates to its calendar date (FromOADate semantics, shared with the
    // table reader so the two cannot disagree).
    [InlineData("0")]
    [InlineData("-5")]
    [InlineData("2958466")]
    public void TryParsePlotDate_refuses_an_out_of_range_serial(string text)
    {
        Assert.False(PlotRangeResolver.TryParsePlotDate(text, out DateOnly parsed));
        Assert.Equal(default, parsed);
    }

    [Fact]
    public void TryParsePlotDate_truncates_a_fractional_serial_to_its_calendar_date()
    {
        // 46118.5 is 06/04/2026 midday: the date half survives, the time half
        // is truncated by FromOADate.
        Assert.True(PlotRangeResolver.TryParsePlotDate("46118.5", out DateOnly parsed));
        Assert.Equal(new DateOnly(2026, 4, 6), parsed);
    }

    [Fact]
    public void ValidateSettings_accepts_the_ddMMyyyy_stored_form()
    {
        // The ribbon persists normalised dd/MM/yyyy text; the validator must
        // accept exactly what the setters write.
        PlotRangeResolver.ValidateSettings(
            new PlotRangeSettings
            {
                Mode = nameof(PlotRangeMode.Explicit),
                StartDate = "15/03/2026",
                FinishDate = "30/06/2026",
            });
    }

    [Fact]
    public void ValidateSettings_rejects_a_non_date_with_an_actionable_exception()
    {
        // Positive test for the validator's bad-input path.
        var exception = Assert.Throws<ArgumentException>(
            () => PlotRangeResolver.ValidateSettings(
                new PlotRangeSettings
                {
                    Mode = nameof(PlotRangeMode.Explicit),
                    StartDate = "not-a-date",
                    FinishDate = "30/06/2026",
                }));

        Assert.Contains("Unparsable start date", exception.Message, StringComparison.Ordinal);
    }
}
