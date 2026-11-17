using System.Globalization;
using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// R5.1a: the plot-date calendar pickers, the input-time inversion guard, the
/// blank-entry revert and the split settings payload.
/// </summary>
/// <remarks>
/// The picker seam is substituted, so no test builds a WinForms control off the
/// STA thread and no test can block on a modal dialog.
/// </remarks>
public sealed class PlotDatePickerTests
{
    private static int s_nextRowId;

    [Fact]
    public void PickPlotStartDate_seeds_from_the_effective_date_and_bounds_by_the_finish()
    {
        PickService service = new();
        var writer = new Mock<IConfigCatalogueWriter>();
        _ = writer.Setup(w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()))
            .Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);
        service.SetPlotStartDate("01-Mar-26");
        service.SetPlotFinishDate("30-Apr-26");
        service.Answer = new DateOnly(2026, 3, 15);

        service.PickPlotStartDate();

        PlotDatePickerRequest request = Assert.Single(service.Requests);
        Assert.Equal(new DateOnly(2026, 3, 1), request.Initial);
        Assert.Null(request.Minimum);
        Assert.Equal(new DateOnly(2026, 4, 30), request.Maximum);
        Assert.Equal("15-Mar-26", service.GetPlotStartDate());
        Assert.False(service.IsPlotStartAuto());
    }

    [Fact]
    public void PickPlotFinishDate_bounds_from_the_start()
    {
        PickService service = new();
        var writer = new Mock<IConfigCatalogueWriter>();
        _ = writer.Setup(w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()))
            .Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);
        service.SetPlotStartDate("01-Mar-26");
        service.SetPlotFinishDate("30-Apr-26");
        service.Answer = new DateOnly(2026, 4, 20);

        service.PickPlotFinishDate();

        PlotDatePickerRequest request = Assert.Single(service.Requests);
        Assert.Equal(new DateOnly(2026, 3, 1), request.Minimum);
        Assert.Null(request.Maximum);
        Assert.Equal("20-Apr-26", service.GetPlotFinishDate());
    }

    [Fact]
    public void PickPlotStartDate_writes_nothing_when_the_dialog_is_cancelled()
    {
        PickService service = new();
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        service.SetCatalogueWriter(null);
        service.SetPlotStartDate("01-Mar-26");
        service.SetCatalogueWriter(writer.Object);
        var ribbon = new Mock<IRibbonUI>();
        service.SetRibbon(ribbon.Object);
        service.Answer = null;

        service.PickPlotStartDate();

        _ = Assert.Single(service.Requests);
        Assert.Equal("01-Mar-26", service.GetPlotStartDate());
        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Never);
        Assert.Empty(InvalidatedIds(ribbon));
    }

    [Fact]
    public void PickPlotStartDate_is_a_no_op_for_the_date_already_in_force()
    {
        PickService service = new();
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        service.SetCatalogueWriter(null);
        service.SetPlotStartDate("15-Mar-26");
        service.SetCatalogueWriter(writer.Object);
        service.Answer = new DateOnly(2026, 3, 15);

        service.PickPlotStartDate();

        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Never);
    }

    [Fact]
    public void PickPlotStartDate_never_opens_the_picker_in_auto_mode()
    {
        PickService service = new();
        service.SetCatalogueWriter(null);

        Assert.True(service.IsPlotStartAuto());
        service.PickPlotStartDate();

        Assert.Empty(service.Requests);
    }

    [Fact]
    public void PickPlotStartDate_falls_back_to_the_clock_when_no_effective_date_parses()
    {
        PickService service = new() { TodayAnswer = new DateOnly(2026, 7, 4) };
        service.SetCatalogueWriter(null);
        var reader = new Mock<IConfigCatalogueReader>();
        _ = reader.Setup(r => r.Read()).Returns(ConfigReadOutcome.Ok(
            "wb",
            new Dictionary<string, string>
            {
                ["PlotStartMode"] = "Explicit",
                ["PlotFinishMode"] = "DataRange",
                ["PlotStartDate"] = string.Empty,
                ["PlotFinishDate"] = string.Empty,
            },
            GanttStyleRegistry.Empty,
            new Dictionary<string, double>()));
        service.SetCatalogueReader(reader.Object);

        service.Refresh();
        service.Answer = new DateOnly(2026, 7, 4);
        service.PickPlotStartDate();

        PlotDatePickerRequest request = Assert.Single(service.Requests);
        Assert.Equal(new DateOnly(2026, 7, 4), request.Initial);
    }

    [Fact]
    public void SetPlotStartDate_refuses_a_start_after_the_auto_derived_finish()
    {
        PickService service = new();
        service.SetCatalogueWriter(null);
        ArmingDerivedBounds(service);
        Mock<IConfigCatalogueWriter> writer = UnusedWriter();
        service.SetCatalogueWriter(writer.Object);

        // The derived finish is 30-Apr-26; a later start must be refused at the
        // input rather than failing the whole refresh with StartAfterFinish.
        service.SetPlotStartDate("15-May-26");

        Assert.Contains("after the plot finish", Assert.Single(service.Warnings), StringComparison.Ordinal);
        Assert.Equal("01-Mar-26", service.GetPlotStartDate());
        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Never);
    }

    [Fact]
    public void SetPlotFinishDate_refuses_a_finish_before_the_auto_derived_start()
    {
        PickService service = new();
        service.SetCatalogueWriter(null);
        ArmingDerivedBounds(service);
        Mock<IConfigCatalogueWriter> writer = UnusedWriter();
        service.SetCatalogueWriter(writer.Object);

        service.SetPlotFinishDate("28-Feb-26");

        Assert.Contains("before the plot start", Assert.Single(service.Warnings), StringComparison.Ordinal);
        Assert.Equal("30-Apr-26", service.GetPlotFinishDate());
        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Never);
    }

    [Fact]
    public void SetPlotStartDate_accepts_a_single_day_range()
    {
        PickService service = new();
        var writer = new Mock<IConfigCatalogueWriter>();
        _ = writer.Setup(w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()))
            .Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);
        service.SetPlotFinishDate("30-Apr-26");

        service.SetPlotStartDate("30-Apr-26");

        Assert.Empty(service.Warnings);
        Assert.Equal("30-Apr-26", service.GetPlotStartDate());
        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.AtLeastOnce);
    }

    [Fact]
    public void SetPlotStartDate_commits_when_the_opposite_end_has_no_readable_date()
    {
        PickService service = new();
        var writer = new Mock<IConfigCatalogueWriter>();
        IReadOnlyDictionary<string, string>? written = null;
        _ = writer.Setup(w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()))
            .Callback<IReadOnlyDictionary<string, string>>(values => written = values)
            .Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);

        // Nothing has been loaded, so the automatic finish has no derived date to
        // invert against: the only pass-through of the guard.
        service.SetPlotStartDate("15-Mar-26");

        Assert.Empty(service.Warnings);
        Assert.NotNull(written);
    }

    [Fact]
    public void SetPlotStartDate_blank_reverts_silently_without_writing()
    {
        PickService service = new();
        service.SetCatalogueWriter(null);
        service.SetPlotStartDate("15-Mar-26");
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        service.SetCatalogueWriter(writer.Object);
        var ribbon = new Mock<IRibbonUI>();
        service.SetRibbon(ribbon.Object);

        service.SetPlotStartDate("   ");

        Assert.Empty(service.Warnings);
        Assert.Equal("15-Mar-26", service.GetPlotStartDate());
        Assert.False(service.IsPlotStartAuto());
        Assert.Equal(
            [
                RibbonControlIds.PlotStartDate,
                RibbonControlIds.PlotStartCalendar,
                RibbonControlIds.PlotStartAuto,
            ],
            InvalidatedIds(ribbon));
    }

    [Fact]
    public void A_plot_commit_writes_only_the_four_plot_range_keys()
    {
        PickService service = new();
        var writer = new Mock<IConfigCatalogueWriter>();
        IReadOnlyDictionary<string, string>? written = null;
        _ = writer.Setup(w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()))
            .Callback<IReadOnlyDictionary<string, string>>(values => written = values)
            .Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);

        service.SetPlotStartDate("15-Mar-26");

        Assert.NotNull(written);
        Assert.Equal(
            ["PlotFinishDate", "PlotFinishMode", "PlotStartDate", "PlotStartMode"],
            written!.Keys.OrderBy(key => key, StringComparer.Ordinal));
    }

    [Fact]
    public void A_layout_commit_writes_only_the_five_layout_keys()
    {
        PickService service = new();
        var writer = new Mock<IConfigCatalogueWriter>();
        IReadOnlyDictionary<string, string>? written = null;
        _ = writer.Setup(w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()))
            .Callback<IReadOnlyDictionary<string, string>>(values => written = values)
            .Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);

        service.SetTimeScale(GanttTimeScale.Quarter);

        Assert.NotNull(written);
        Assert.Equal(
            ["Margin", "MarginCm", "PeriodLabelFormat", "SizePreset", "TimeScale"],
            written!.Keys.OrderBy(key => key, StringComparer.Ordinal));
    }

    [Fact]
    public void SetPlotStartDate_reverts_the_row_when_the_write_is_refused()
    {
        PickService service = new();
        service.SetCatalogueWriter(null);
        service.SetPlotStartDate("01-Mar-26");
        var writer = new Mock<IConfigCatalogueWriter>();
        _ = writer.Setup(w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()))
            .Returns(ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.HostRejected));
        service.SetCatalogueWriter(writer.Object);
        var ribbon = new Mock<IRibbonUI>();
        service.SetRibbon(ribbon.Object);

        service.SetPlotStartDate("15-Mar-26");

        Assert.Equal("01-Mar-26", service.GetPlotStartDate());
        Assert.Contains("could not be saved", Assert.Single(service.Warnings), StringComparison.Ordinal);
        Assert.Contains(RibbonControlIds.PlotStartDate, InvalidatedIds(ribbon));
    }

    [Fact]
    public void UpdateEffectivePlotDates_refreshes_only_an_auto_end_whose_text_changed()
    {
        PickService service = new();
        service.SetCatalogueWriter(null);
        ArmingDerivedBounds(service);
        var ribbon = new Mock<IRibbonUI>();
        service.SetRibbon(ribbon.Object);

        // Pin the arrange: the pass derived the month-snapped bounds.
        Assert.Equal("01-Mar-26", service.GetPlotStartDate());
        Assert.Equal("30-Apr-26", service.GetPlotFinishDate());

        // Same range the pass rendered: nothing to repaint.
        service.UpdateEffectivePlotDates(new DateOnly(2026, 3, 1), new DateOnly(2026, 4, 30));
        Assert.Empty(InvalidatedIds(ribbon));

        // Only the finish moved.
        service.UpdateEffectivePlotDates(new DateOnly(2026, 3, 1), new DateOnly(2026, 5, 31));
        Assert.Equal("31-May-26", service.GetPlotFinishDate());
        Assert.Equal([RibbonControlIds.PlotFinishDate], InvalidatedIds(ribbon));
    }

    [Fact]
    public void UpdateEffectivePlotDates_never_rewrites_an_explicit_end()
    {
        PickService service = new();
        var writer = new Mock<IConfigCatalogueWriter>();
        _ = writer.Setup(w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()))
            .Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);
        service.SetPlotStartDate("15-Mar-26");

        service.UpdateEffectivePlotDates(new DateOnly(2026, 2, 1), new DateOnly(2026, 2, 28));

        Assert.Equal("15-Mar-26", service.GetPlotStartDate());
    }

    [Fact]
    public void OnPlotStartCalendarClick_routes_to_the_picker()
    {
        PickService service = new();
        var writer = new Mock<IConfigCatalogueWriter>();
        _ = writer.Setup(w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()))
            .Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);
        service.SetPlotStartDate("01-Mar-26");
        service.SetPlotFinishDate("30-Apr-26");
        service.Answer = new DateOnly(2026, 3, 15);

        GanttRibbon.OnPlotStartCalendarClick(null, service);

        _ = Assert.Single(service.Requests);
        Assert.Equal("15-Mar-26", service.GetPlotStartDate());
    }

    [Fact]
    public void OnPlotFinishCalendarClick_throws_for_a_null_state_service()
    {
        Assert.Throws<ArgumentNullException>(() => GanttRibbon.OnPlotFinishCalendarClick(null, null!));
    }

    /// <summary>
    /// Loads a table whose spans derive the month-snapped bounds 01-Mar-26 to
    /// 30-Apr-26, with both plot ends automatic — the state the inversion guard
    /// has to read through.
    /// </summary>
    /// <param name="service">The service to arm.</param>
    private static void ArmingDerivedBounds(PickService service)
    {
        var reader = new Mock<IConfigCatalogueReader>();
        _ = reader.Setup(r => r.Read()).Returns(ConfigReadOutcome.Ok(
            "wb",
            new Dictionary<string, string>
            {
                ["PlotStartMode"] = "DataRange",
                ["PlotFinishMode"] = "DataRange",
                ["PlotStartDate"] = string.Empty,
                ["PlotFinishDate"] = string.Empty,
            },
            GanttStyleRegistry.Empty,
            new Dictionary<string, double>()));
        service.SetCatalogueReader(reader.Object);

        var table = new Mock<IGanttTableReader>();
        _ = table.Setup(r => r.Read()).Returns(GanttTableReadOutcome.Ok(
        [
            PlotBoundsRow(1, "10/03/2026", "31/03/2026"),
        ]));
        service.SetTableReader(table.Object);

        service.Refresh();
    }

    private static string RowId() =>
        "G-" + (++s_nextRowId).ToString("x32", CultureInfo.InvariantCulture);

    private static GanttRowDto PlotBoundsRow(int row, string start, string finish) =>
        new(
            row,
            RowId(),
            RowId(),
            null,
            "As-Planned Activity",
            $"Row {row}",
            DateOnly.ParseExact(start, "dd/MM/yyyy", CultureInfo.InvariantCulture),
            DateOnly.ParseExact(finish, "dd/MM/yyyy", CultureInfo.InvariantCulture),
            null,
            null,
            null,
            null,
            null,
            true,
            null);

    /// <summary>A writer that must never be called, kept loose so a stray call is reported by Verify.</summary>
    /// <returns>The mock to pass to the service.</returns>
    private static Mock<IConfigCatalogueWriter> UnusedWriter()
    {
        var writer = new Mock<IConfigCatalogueWriter>();
        _ = writer.Setup(w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()))
            .Returns(ConfigWriteOutcome.Ok());
        return writer;
    }

    private static List<string> InvalidatedIds(Mock<IRibbonUI> ribbon) =>
    [
        .. ribbon.Invocations
            .Where(invocation => invocation.Method.Name == nameof(IRibbonUI.InvalidateControl))
            .Select(invocation => (string)invocation.Arguments[0]!),
    ];

    /// <summary>
    /// The service under test, with the picker and clock seams captured instead of
    /// shown.
    /// </summary>
    private sealed class PickService : RibbonStateService
    {
        internal List<PlotDatePickerRequest> Requests { get; } = [];

        internal List<string> Warnings { get; } = [];

        internal DateOnly? Answer { get; set; }

        internal DateOnly TodayAnswer { get; set; } = new(2026, 1, 5);

        internal PickService() => SetWarningPresenter(message => Warnings.Add(message));

        protected override DateOnly? ShowPlotDatePicker(PlotDatePickerRequest request)
        {
            Requests.Add(request);
            return Answer;
        }

        internal override DateOnly Today() => TodayAnswer;
    }
}
