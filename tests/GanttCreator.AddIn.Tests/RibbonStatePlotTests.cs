
using GanttCreator.Core;
using GanttCreator.Office;
using Moq;
using static GanttCreator.AddIn.Tests.RibbonStateTestSupport;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The plot-range and scale preferences of <see cref="RibbonStateService"/>: the AUTO/EXPLICIT
/// toggles at each end, the date boxes that normalise and revert, the stored month label form, and
/// the rule that a refused write restores the value the workbook actually holds.
/// </summary>
/// <remarks>
/// Shares the <c>diagnostics-service</c> collection: these facts arm the same session singleton the
/// host-lifecycle facts arm and reset, so xUnit runs the collection's classes sequentially. The
/// constructor and <see cref="Dispose"/> reset it around every fact.
/// </remarks>
[Collection("diagnostics-service")]
public sealed class RibbonStatePlotTests : IDisposable
{
    public RibbonStatePlotTests() => RibbonStateService.Reset();

    public void Dispose()
    {
        RibbonStateService.Reset();
        GC.SuppressFinalize(this);
    }
    [Fact]
    public void TogglePlotStartAuto_from_explicit_to_auto_persists_and_preserves_the_start_date()
    {
        // Fix plan ruling 1: entering AUTO no longer clears the stored date —
        // the disabled edit box keeps showing it until the effective-date
        // derivation replaces it, and returning to explicit restores it. The
        // R5.1 selection-change invariant holds: no render; the persisted
        // payload is the assertion, not a ribbon refresh, which Set never
        // triggers. The finish end — still explicit — is preserved untouched.
        var service = RibbonStateService.Instance;
        service.SetPlotStartDate("15-Mar-26");
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        _ = writer
            .Setup(w => w.WriteSettings(It.Is<IReadOnlyDictionary<string, string>>(
                settings =>
                    settings["PlotStartMode"] == "DataRange"
                    && settings["PlotFinishMode"] == "DataRange"
                    && settings["PlotStartDate"] == "15-Mar-26"
                    && settings["PlotFinishDate"] == string.Empty)))
            .Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);

        service.TogglePlotStartAuto();

        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Once);
        Assert.True(service.IsPlotStartAuto());
        Assert.Equal("15-Mar-26", service.GetPlotStartDate());
    }

    /// <summary>
    /// The month label form is the user's stored selection, not a derivation of
    /// the scale: choosing the MM row persists MM even though the month scale's
    /// canonical form is MMM. Without this pin the persist path would derive
    /// the format from the scale on every commit and silently clobber the
    /// selection back to MMM the next time any ribbon setting changed (R5.12).
    /// </summary>
    [Fact]
    public void SetTimeScale_persists_the_selected_month_form()
    {
        var service = RibbonStateService.Instance;
        service.SetCatalogueWriter(null);
        service.SetTimeScale(GanttTimeScale.Month);

        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        _ = writer
            .Setup(w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()))
            .Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);

        service.SetTimeScale(GanttTimeScale.Month, GanttPeriodLabelFormat.MM);

        Assert.Equal(GanttPeriodLabelFormat.MM, service.GetPeriodLabelFormat());

        // The predicate IS the assertion: the persisted payload must carry the
        // selection. A payload that derived the form from the scale would write
        // MMM, and this Verify finds no matching invocation.
        writer.Verify(
            w => w.WriteSettings(It.Is<IReadOnlyDictionary<string, string>>(
                settings => settings["PeriodLabelFormat"] == nameof(GanttPeriodLabelFormat.MM))),
            Times.Once);
    }

    /// <summary>
    /// A named form the scale cannot carry is refused, not coerced, so the
    /// service cannot be driven into storing a pair the catalogue reader
    /// rejects. The stored pair is left exactly as it was.
    /// </summary>
    [Fact]
    public void SetTimeScale_refuses_a_named_form_the_scale_cannot_carry()
    {
        var service = RibbonStateService.Instance;
        service.SetCatalogueWriter(null);
        service.SetTimeScale(GanttTimeScale.Month, GanttPeriodLabelFormat.MM);

        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        service.SetCatalogueWriter(writer.Object);

        service.SetTimeScale(GanttTimeScale.Week, GanttPeriodLabelFormat.MM);

        Assert.Equal(GanttTimeScale.Month, service.GetTimeScale());
        Assert.Equal(GanttPeriodLabelFormat.MM, service.GetPeriodLabelFormat());
        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Never);
    }

    /// <summary>
    /// A scale switch normalises an incompatible stored form to the new scale's
    /// canonical format, so the stored pair the catalogue reader validates is
    /// never MM-against-Week. MM is the notable case: it does not survive the
    /// round trip, because the pair — not the form — is the stored unit.
    /// </summary>
    [Fact]
    public void SetTimeScale_normalises_an_incompatible_stored_format()
    {
        var service = RibbonStateService.Instance;
        service.SetCatalogueWriter(null);
        service.SetTimeScale(GanttTimeScale.Month, GanttPeriodLabelFormat.MM);

        service.SetTimeScale(GanttTimeScale.Quarter);

        Assert.Equal(GanttPeriodLabelFormat.Quarter, service.GetPeriodLabelFormat());

        service.SetTimeScale(GanttTimeScale.Month);

        Assert.Equal(GanttPeriodLabelFormat.MMM, service.GetPeriodLabelFormat());
    }

    /// <summary>
    /// The stored month form is adopted on capture, so reopening a workbook
    /// saved with MM restores the numeric band instead of silently reverting to
    /// the canonical MMM.
    /// </summary>
    [Fact]
    public void Refresh_adopts_the_stored_month_label_form()
    {
        var service = RibbonStateService.Instance;
        service.SetCatalogueWriter(null);
        service.SetTimeScale(GanttTimeScale.Month);
        service.SetTimeScale(GanttTimeScale.Month, GanttPeriodLabelFormat.MMM);

        var catalogueReader = new Mock<IConfigCatalogueReader>(MockBehavior.Strict);
        _ = catalogueReader
            .Setup(r => r.Read())
            .Returns(ConfigReadOutcome.Ok(
                "workbook",
                new Dictionary<string, string>
                {
                    ["PlotStartMode"] = nameof(PlotRangeMode.DataRange),
                    ["PlotFinishMode"] = nameof(PlotRangeMode.DataRange),
                    ["TimeScale"] = nameof(GanttTimeScale.Month),
                    ["PeriodLabelFormat"] = nameof(GanttPeriodLabelFormat.MM),
                },
                GanttStyleRegistry.Empty, new Dictionary<string, double>()));
        service.SetCatalogueReader(catalogueReader.Object);

        service.Refresh();

        Assert.Equal(GanttPeriodLabelFormat.MM, service.GetPeriodLabelFormat());
    }

    [Fact]
    public void TogglePlotStartAuto_from_auto_to_explicit_seeds_from_the_current_plot_bound()
    {
        // Owner ruling 2026-10-08: unchecking AUTO seeds the edit box from
        // the CURRENT plot bound (the month-snapped chart bound the AUTO box
        // was displaying) — never from the stale stored date — and persists
        // Explicit at once, so the next Refresh renders exactly what the box
        // shows. The workbook holds a stale stored 15-Mar-26 while the data
        // derives the 01-Jun-26 bound; the seeded date must be the bound.
        var service = RibbonStateService.Instance;
        var catalogueReader = new Mock<IConfigCatalogueReader>(MockBehavior.Strict);
        _ = catalogueReader
            .Setup(r => r.Read())
            .Returns(ConfigReadOutcome.Ok(
                "workbook",
                new Dictionary<string, string>
                {
                    ["PlotStartMode"] = "DataRange",
                    ["PlotFinishMode"] = "DataRange",
                    ["PlotStartDate"] = "15-Mar-26",
                    ["PlotFinishDate"] = string.Empty,
                },
                GanttStyleRegistry.Empty, new Dictionary<string, double>()));
        var tableReader = new Mock<IGanttTableReader>(MockBehavior.Strict);
        _ = tableReader
            .Setup(r => r.Read())
            .Returns(GanttTableReadOutcome.Ok(
            [
                new GanttRowDto(
                    1,
                    NewRowId(),
                    NewRowId('b'),
                    0,
                    "As-Planned Activity",
                    "Seeded row",
                    new DateOnly(2026, 6, 10),
                    new DateOnly(2026, 6, 20),
                    null,
                    null,
                    null,
                    null,
                    null,
                    true,
                    null),
            ]));
        service.SetCatalogueReader(catalogueReader.Object);
        service.SetTableReader(tableReader.Object);
        service.Refresh();
        Assert.Equal("01-Jun-26", service.GetPlotStartDate());
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        _ = writer
            .Setup(w => w.WriteSettings(It.Is<IReadOnlyDictionary<string, string>>(
                settings =>
                    settings["PlotStartMode"] == "Explicit"
                    && settings["PlotStartDate"] == "01-Jun-26")))
            .Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);

        service.SetPlotStartAuto(false);

        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Once);
        Assert.False(service.IsPlotStartAuto());
        Assert.Equal("01-Jun-26", service.GetPlotStartDate());
    }

    [Fact]
    public void TogglePlotStartAuto_from_auto_to_explicit_arms_the_edit_box_without_persisting()
    {
        // No current bound is known here (no effective date was ever
        // derived), so the uncheck arms the edit box without persisting:
        // an Explicit mode without a date is the MissingExplicitDate
        // refusal and would break the next Refresh. The date commit owns
        // that write.
        var service = RibbonStateService.Instance;
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        service.SetCatalogueWriter(writer.Object);

        service.TogglePlotStartAuto();

        Assert.False(service.IsPlotStartAuto());
        Assert.Equal(string.Empty, service.GetPlotStartDate());
        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Never);
    }

    [Fact]
    public void TogglePlotStartAuto_degrades_gracefully_without_a_writer()
    {
        // A null writer leaves the persist step as a no-op: the in-memory
        // state commit must not throw and no refresh is required. Start from
        // explicit so the toggle actually reaches the persist step.
        var service = RibbonStateService.Instance;
        service.SetCatalogueWriter(null);
        service.SetPlotStartDate("15-Mar-26");

        var exception = Record.Exception(service.TogglePlotStartAuto);

        Assert.Null(exception);
        Assert.True(service.IsPlotStartAuto());
    }

    [Fact]
    public void TogglePlotFinishAuto_from_explicit_to_auto_persists_and_preserves_the_finish_date()
    {
        // Per-end analogue of the ruling-1 preserve: the start end (also
        // explicit) is preserved untouched in the payload.
        var service = RibbonStateService.Instance;
        service.SetPlotStartDate("01-Mar-26");
        service.SetPlotFinishDate("15-Jun-26");
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        _ = writer
            .Setup(w => w.WriteSettings(It.Is<IReadOnlyDictionary<string, string>>(
                settings =>
                    settings["PlotStartMode"] == "Explicit"
                    && settings["PlotFinishMode"] == "DataRange"
                    && settings["PlotStartDate"] == "01-Mar-26"
                    && settings["PlotFinishDate"] == "15-Jun-26")))
            .Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);

        service.TogglePlotFinishAuto();

        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Once);
        Assert.True(service.IsPlotFinishAuto());
        Assert.Equal("15-Jun-26", service.GetPlotFinishDate());
        Assert.Equal("01-Mar-26", service.GetPlotStartDate());
    }

    [Fact]
    public void TogglePlotFinishAuto_from_auto_to_explicit_seeds_from_the_current_plot_bound()
    {
        // Owner ruling 2026-10-08, finish-end analogue: unchecking AUTO
        // seeds from the current plot bound, never the stale stored date,
        // and persists Explicit at once.
        var service = RibbonStateService.Instance;
        var catalogueReader = new Mock<IConfigCatalogueReader>(MockBehavior.Strict);
        _ = catalogueReader
            .Setup(r => r.Read())
            .Returns(ConfigReadOutcome.Ok(
                "workbook",
                new Dictionary<string, string>
                {
                    ["PlotStartMode"] = "DataRange",
                    ["PlotFinishMode"] = "DataRange",
                    ["PlotStartDate"] = string.Empty,
                    ["PlotFinishDate"] = "15-Jun-26",
                },
                GanttStyleRegistry.Empty, new Dictionary<string, double>()));
        var tableReader = new Mock<IGanttTableReader>(MockBehavior.Strict);
        _ = tableReader
            .Setup(r => r.Read())
            .Returns(GanttTableReadOutcome.Ok(
            [
                new GanttRowDto(
                    1,
                    NewRowId(),
                    NewRowId('b'),
                    0,
                    "As-Planned Activity",
                    "Seeded row",
                    new DateOnly(2026, 6, 10),
                    new DateOnly(2026, 9, 20),
                    null,
                    null,
                    null,
                    null,
                    null,
                    true,
                    null),
            ]));
        service.SetCatalogueReader(catalogueReader.Object);
        service.SetTableReader(tableReader.Object);
        service.Refresh();
        Assert.Equal("30-Sep-26", service.GetPlotFinishDate());
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        _ = writer
            .Setup(w => w.WriteSettings(It.Is<IReadOnlyDictionary<string, string>>(
                settings =>
                    settings["PlotFinishMode"] == "Explicit"
                    && settings["PlotFinishDate"] == "30-Sep-26")))
            .Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);

        service.SetPlotFinishAuto(false);

        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Once);
        Assert.False(service.IsPlotFinishAuto());
        Assert.Equal("30-Sep-26", service.GetPlotFinishDate());
    }

    [Fact]
    public void TogglePlotFinishAuto_from_auto_to_explicit_arms_the_edit_box_without_persisting()
    {
        // Mirror of the start-side rule: no Explicit-without-date write.
        var service = RibbonStateService.Instance;
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        service.SetCatalogueWriter(writer.Object);

        service.TogglePlotFinishAuto();

        Assert.False(service.IsPlotFinishAuto());
        Assert.Equal(string.Empty, service.GetPlotFinishDate());
        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Never);
    }

    [Fact]
    public void TogglePlotFinishAuto_degrades_gracefully_without_a_writer()
    {
        var service = RibbonStateService.Instance;
        service.SetCatalogueWriter(null);
        service.SetPlotFinishDate("15-Jun-26");

        var exception = Record.Exception(service.TogglePlotFinishAuto);
        Assert.Null(exception);
        Assert.True(service.IsPlotFinishAuto());
    }

    [Fact]
    public void Refresh_normalises_a_stored_serial_to_ddMMyyyy_instead_of_echoing_the_number()
    {
        // Positive test for the numeric-echo load path: a workbook holding a
        // bare Excel serial (06-Apr-26 = 06-Apr-26, e.g. persisted pre-fix or
        // hand-edited) must display the standard formatted date, not the
        // number, after the workbook load. Both ends are explicit so no
        // table derivation is involved.
        var service = RibbonStateService.Instance;
        var catalogueReader = new Mock<IConfigCatalogueReader>(MockBehavior.Strict);
        _ = catalogueReader
            .Setup(r => r.Read())
            .Returns(ConfigReadOutcome.Ok(
                "workbook",
                new Dictionary<string, string>
                {
                    ["PlotStartMode"] = "Explicit",
                    ["PlotFinishMode"] = "Explicit",
                    ["PlotStartDate"] = "06-Apr-26",
                    ["PlotFinishDate"] = "13-Jul-26",
                },
                GanttStyleRegistry.Empty, new Dictionary<string, double>()));
        service.SetCatalogueReader(catalogueReader.Object);

        service.Refresh();

        Assert.False(service.IsPlotStartAuto());
        Assert.False(service.IsPlotFinishAuto());
        Assert.Equal("06-Apr-26", service.GetPlotStartDate());
        Assert.Equal("13-Jul-26", service.GetPlotFinishDate());
    }

    [Fact]
    public void Refresh_treats_a_miscased_mode_as_explicit_like_the_refresh_pipeline()
    {
        // Positive test for the single-authority mode parse: the catalogue
        // layer refuses miscased modes (PlotRangeModes.TryParse is ordinal),
        // so a display path that case-folded "datarange" to AUTO would show
        // AUTO while Refresh refuses UnknownMode. The ribbon must agree
        // with Refresh: unknown text means explicit.
        var service = RibbonStateService.Instance;
        var catalogueReader = new Mock<IConfigCatalogueReader>(MockBehavior.Strict);
        _ = catalogueReader
            .Setup(r => r.Read())
            .Returns(ConfigReadOutcome.Ok(
                "workbook",
                new Dictionary<string, string>
                {
                    ["PlotStartMode"] = "datarange",
                    ["PlotFinishMode"] = "datarange",
                    ["PlotStartDate"] = "01-Jun-26",
                    ["PlotFinishDate"] = "30-Sep-26",
                },
                GanttStyleRegistry.Empty, new Dictionary<string, double>()));
        service.SetCatalogueReader(catalogueReader.Object);

        service.Refresh();

        Assert.False(service.IsPlotStartAuto());
        Assert.False(service.IsPlotFinishAuto());
        Assert.Equal("01-Jun-26", service.GetPlotStartDate());
        Assert.Equal("30-Sep-26", service.GetPlotFinishDate());
    }

    [Fact]
    public void SetPlotStartDate_normalises_ddMMyyyy_and_pivots_to_explicit()
    {
        // A valid typed date pivots the start end to Explicit and normalises
        // to the repository dd/MM/yyyy storage form, which the getText getter
        // then returns.
        var service = RibbonStateService.Instance;
        service.SetCatalogueWriter(null);

        service.SetPlotStartDate("15-Mar-26");

        Assert.False(service.IsPlotStartAuto());
        Assert.Equal("15-Mar-26", service.GetPlotStartDate());
    }

    [Fact]
    public void SetPlotStartDate_refuses_activity_auto_parser_equivalents()
    {
        // Positive for the exact-date rule: the plot boxes take the chart's
        // dd/MM/yyyy dates only, never the activity multi-format equivalents.
        // ISO, US month-first, and dd-MMM-yyyy entries are no-ops that leave
        // the stored value, mode, and writer untouched.
        var service = RibbonStateService.Instance;
        service.SetCatalogueWriter(null);
        service.SetPlotStartDate("15-Mar-26");
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        service.SetCatalogueWriter(writer.Object);

        service.SetPlotStartDate("2026-03-15");
        service.SetPlotStartDate("03/15/2026");
        service.SetPlotStartDate("15-Mar-2026");

        Assert.False(service.IsPlotStartAuto());
        Assert.Equal("15-Mar-26", service.GetPlotStartDate());
        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Never);
    }

    [Fact]
    public void SetPlotStartDate_reverts_on_an_unparseable_entry()
    {
        // Positive for the revert rule: a non-date leaves the last valid
        // stored value and mode untouched, and never reaches the writer, so
        // the edit box keeps showing the previously stored date on the next
        // getText query.
        var service = RibbonStateService.Instance;
        service.SetPlotStartDate("15-Mar-26");
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        service.SetCatalogueWriter(writer.Object);

        service.SetPlotStartDate("not-a-date");

        Assert.False(service.IsPlotStartAuto());
        Assert.Equal("15-Mar-26", service.GetPlotStartDate());
        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Never);
    }

    [Fact]
    public void SetPlotStartDate_ignores_blank_input()
    {
        // Positive: clearing the box is a no-op rather than a blank commit.
        var service = RibbonStateService.Instance;
        service.SetPlotStartDate("15-Mar-26");
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        service.SetCatalogueWriter(writer.Object);

        service.SetPlotStartDate("   ");

        Assert.Equal("15-Mar-26", service.GetPlotStartDate());
        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Never);
    }

    [Fact]
    public void SetPlotFinishDate_normalises_ddMMyyyy_and_pivots_to_explicit()
    {
        var service = RibbonStateService.Instance;
        service.SetCatalogueWriter(null);

        service.SetPlotFinishDate("30-Jun-26");

        Assert.False(service.IsPlotFinishAuto());
        Assert.Equal("30-Jun-26", service.GetPlotFinishDate());
    }

    [Fact]
    public void SetPlotFinishDate_reverts_on_an_unparseable_entry()
    {
        // Positive for the revert rule on the finish end.
        var service = RibbonStateService.Instance;
        service.SetPlotFinishDate("30-Jun-26");
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        service.SetCatalogueWriter(writer.Object);

        service.SetPlotFinishDate("31/02/2026");

        Assert.False(service.IsPlotFinishAuto());
        Assert.Equal("30-Jun-26", service.GetPlotFinishDate());
        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Never);
    }
}
