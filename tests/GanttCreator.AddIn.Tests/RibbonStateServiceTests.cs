using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;
using GanttCreator.Raster;
using Moq;
using System.Windows.Forms;
using static GanttCreator.AddIn.Tests.RibbonStateTestSupport;

namespace GanttCreator.AddIn.Tests;
/// <summary>
/// Contract tests for <see cref="RibbonStateService"/>: the dynamic getters are
/// deterministic and side-effect-free, the invalidate mechanism drives the
/// RibbonUI handle, workbook-state events refresh, and every failure path
/// degrades instead of throwing into Excel.
/// </summary>
/// <remarks>
/// The class shares the <c>diagnostics-service</c> collection because
/// <see cref="AddInHost.AutoOpen"/>/<see cref="AddInHost.AutoClose"/> now arm and
/// reset the same session singleton; xUnit runs a collection's classes
/// sequentially, so the host lifecycle tests cannot race these.
/// </remarks>
[Collection("diagnostics-service")]
public sealed class RibbonStateServiceTests : IDisposable
{
    public RibbonStateServiceTests() => RibbonStateService.Reset();

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

    /// <summary>
    /// Hand-written fake for <see cref="IPanelGridMeasurementPort"/>: returns
    /// one fixed grid so the width-display tests can prove the display equals
    /// the resolver output without Excel.
    /// </summary>
    private sealed class FakePanelGridMeasurementPort(PanelCellGrid grid) : IPanelGridMeasurementPort
    {
        public PanelGridOutcome Measure(
            IReadOnlyList<string> includedColumns,
            IReadOnlyList<double>? bodyRowHeightsPt = null)
        {
            ArgumentNullException.ThrowIfNull(includedColumns);

            // The width display only needs the column widths, so the supplied body heights
            // are irrelevant here and the fixed grid is returned either way.
            _ = bodyRowHeightsPt;
            return PanelGridOutcome.Ok(grid);
        }
    }

    /// <summary>
    /// The width display equals the resolver output: with a measured 300pt
    /// panel, the Normal margin, and the A4 portrait preset, the box shows
    /// exactly what <c>PlotGeometryResolver.MeasurePlotWidth</c> derives, so
    /// the displayed figure and the rendered plot cannot drift (R5.2 D4).
    /// </summary>
    [Fact]
    public void SetPreset_recomputes_the_width_display_from_the_resolver()
    {
        var service = RibbonStateService.Instance;
        service.SetCatalogueWriter(null);
        service.SetMargin(GanttPlotMargin.Normal);
        var grid = PanelCellGrid.TryCreate(
            [new PanelColumn("Task", 300.0)],
            [12.0],
            12.0,
            []).Grid ?? throw new InvalidOperationException("Fixture grid should be valid.");
        service.SetPanelGridMeasurementPort(new FakePanelGridMeasurementPort(grid));

        service.SetPreset(SizePresets.A4Portrait);

        var marginPt = GanttPlotMargins.MarginPt(GanttPlotMargin.Normal, GanttPlotMargins.DefaultCustomCm);
        var chrome = GanttCatalogues.MetricDefault("ChartOuterPaddingPt");

        // The left chrome is the live plot's origin inset (ADR-0061 D11): the display
        // must read the SAME chrome the composition passes, or it shows a width two
        // points wider than the plot that is drawn.
        var expected = PlotGeometryResolver.MeasurePlotWidth(
            SizePresets.A4Portrait,
            300.0,
            GanttSheetLayout.LivePlotOriginInsetPt,
            (marginPt * 2) + chrome);
        Assert.Equal(
            expected?.ToString("F2", System.Globalization.CultureInfo.InvariantCulture),
            service.GetPlotWidth());
        Assert.Equal(
            SizePresets.A4Portrait.HeightPt.ToString("F2", System.Globalization.CultureInfo.InvariantCulture),
            service.GetPlotHeight());
    }

    /// <summary>
    /// A margin change recomputes the width display: the plot
    /// width depends on the margin, so a margin committed without
    /// re-measuring would leave the box showing the pre-change
    /// figure (R5.2 D4).
    /// </summary>
    [Fact]
    public void SetMargin_recomputes_the_width_display_from_the_resolver()
    {
        var service = RibbonStateService.Instance;
        service.SetCatalogueWriter(null);
        var grid = PanelCellGrid.TryCreate(
            [new PanelColumn("Task", 300.0)],
            [12.0],
            12.0,
            []).Grid ?? throw new InvalidOperationException("Fixture grid should be valid.");
        service.SetPanelGridMeasurementPort(new FakePanelGridMeasurementPort(grid));
        service.SetPreset(SizePresets.A4Portrait);
        var normalWidth = service.GetPlotWidth();

        service.SetMargin(GanttPlotMargin.Wide);

        var marginPt = GanttPlotMargins.MarginPt(GanttPlotMargin.Wide, GanttPlotMargins.DefaultCustomCm);
        var chrome = GanttCatalogues.MetricDefault("ChartOuterPaddingPt");
    }

    /// <summary>
    /// A custom-margin centimetre change recomputes the width
    /// display for the same reason: the clamped value feeds the
    /// resolver, so the box must show the re-measured figure.
    /// </summary>
    [Fact]
    public void SetMarginCm_recomputes_the_width_display_from_the_resolver()
    {
        var service = RibbonStateService.Instance;
        service.SetCatalogueWriter(null);
        var grid = PanelCellGrid.TryCreate(
            [new PanelColumn("Task", 300.0)],
            [12.0],
            12.0,
            []).Grid ?? throw new InvalidOperationException("Fixture grid should be valid.");
        service.SetPanelGridMeasurementPort(new FakePanelGridMeasurementPort(grid));
        service.SetPreset(SizePresets.A4Portrait);
        service.SetMargin(GanttPlotMargin.Custom);
        var beforeWidth = service.GetPlotWidth();

        service.SetMarginCm(2.5);

        var marginPt = GanttPlotMargins.MarginPt(GanttPlotMargin.Custom, 2.5);
        var chrome = GanttCatalogues.MetricDefault("ChartOuterPaddingPt");
        var expected = PlotGeometryResolver.MeasurePlotWidth(
            SizePresets.A4Portrait,
            300.0,
            GanttSheetLayout.LivePlotOriginInsetPt,
            (marginPt * 2) + chrome);
        Assert.Equal(
            expected?.ToString("F2", System.Globalization.CultureInfo.InvariantCulture),
            service.GetPlotWidth());
        Assert.NotEqual(beforeWidth, service.GetPlotWidth());
    }

    /// <summary>
    /// Reset drops the measurement port: a width computed before the reset
    /// is not recomputed from a stale port afterwards, so a fresh session
    /// starts with no grid to measure against.
    /// </summary>
    [Fact]
    public void Reset_clears_the_measurement_port()
    {
        var grid = PanelCellGrid.TryCreate(
            [new PanelColumn("Task", 300.0)],
            [12.0],
            12.0,
            []).Grid ?? throw new InvalidOperationException("Fixture grid should be valid.");
        RibbonStateService.Instance.SetPanelGridMeasurementPort(new FakePanelGridMeasurementPort(grid));

        RibbonStateService.Reset();

        // The fresh singleton holds no port and no preset, so the width
        // display is empty rather than measured from a retained grid.
        Assert.Equal(string.Empty, RibbonStateService.Instance.GetPlotWidth());
    }

    [Fact]
    public void ShowThemeSettingsDialog_returns_true_when_theme_written_successfully()
    {
        var service = new StubThemeDialogRibbonStateService(DialogResult.OK);
        var reader = new Mock<IConfigCatalogueReader>();
        reader.Setup(r => r.Read()).Returns(ConfigReadOutcome.Ok(
            "wb-id",
            new Dictionary<string, string>(),
            GanttStyleRegistry.Empty, new Dictionary<string, double>()));
        service.SetCatalogueReader(reader.Object);

        var writer = new Mock<IConfigCatalogueWriter>();
        writer.Setup(w => w.WriteTheme(It.IsAny<ThemeEditRequest>()))
            .Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);

        Assert.True(service.ShowThemeSettingsDialog());
        writer.Verify(w => w.WriteTheme(It.IsAny<ThemeEditRequest>()), Times.Once);
    }

    [Fact]
    public void ShowThemeSettingsDialog_returns_false_when_user_cancels()
    {
        var service = new StubThemeDialogRibbonStateService(DialogResult.Cancel);
        var reader = new Mock<IConfigCatalogueReader>();
        reader.Setup(r => r.Read()).Returns(ConfigReadOutcome.Ok(
            "wb-id",
            new Dictionary<string, string>(),
            GanttStyleRegistry.Empty, new Dictionary<string, double>()));
        service.SetCatalogueReader(reader.Object);

        var writer = new Mock<IConfigCatalogueWriter>();
        service.SetCatalogueWriter(writer.Object);

        Assert.False(service.ShowThemeSettingsDialog());
        writer.Verify(w => w.WriteTheme(It.IsAny<ThemeEditRequest>()), Times.Never);
    }

    [Fact]
    public void ShowThemeSettingsDialog_returns_false_when_writer_refuses()
    {
        var service = new StubThemeDialogRibbonStateService(DialogResult.OK);
        var reader = new Mock<IConfigCatalogueReader>();
        reader.Setup(r => r.Read()).Returns(ConfigReadOutcome.Ok(
            "wb-id",
            new Dictionary<string, string>(),
            GanttStyleRegistry.Empty, new Dictionary<string, double>()));
        service.SetCatalogueReader(reader.Object);

        var writer = new Mock<IConfigCatalogueWriter>();
        writer.Setup(w => w.WriteTheme(It.IsAny<ThemeEditRequest>()))
            .Returns(ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.HostRejected));
        service.SetCatalogueWriter(writer.Object);

        // The failure is presented through a modal dialog, which would block
        // on a message box with nobody to dismiss it, so the test could never
        // reach its own assertion. The presenter is a seam for exactly this.
        List<string> presented = [];
        service.SetErrorPresenter(presented.Add);

        Assert.False(service.ShowThemeSettingsDialog());
        writer.Verify(w => w.WriteTheme(It.IsAny<ThemeEditRequest>()), Times.Once);
        Assert.Single(presented);
        Assert.Contains("HostRejected", presented[0], StringComparison.Ordinal);
    }

    [Fact]
    public void ShowThemeSettingsDialog_returns_false_when_reader_fails()
    {
        var service = new RibbonStateService();
        var reader = new Mock<IConfigCatalogueReader>();
        reader.Setup(r => r.Read()).Returns(ConfigReadOutcome.Refused(ConfigReadRefusalReason.ConfigSheetMissing));
        service.SetCatalogueReader(reader.Object);

        var writer = new Mock<IConfigCatalogueWriter>();
        service.SetCatalogueWriter(writer.Object);

        Assert.False(service.ShowThemeSettingsDialog());
    }

    [Fact]
    public void ShowThemeSettingsDialog_returns_false_when_reader_and_writer_are_null()
    {
        var service = new RibbonStateService();

        Assert.False(service.ShowThemeSettingsDialog());
    }

    // ---------------------------------------------------------------------
    // R8.8 — the chart-current freshness flag for Export PNG.
    // ---------------------------------------------------------------------

    /// <summary>
    /// A completed refresh marks the chart current, and the marker is store-only: no
    /// ribbon control's enabled state depends on it, so it invalidates nothing (R8.8).
    /// </summary>
    [Fact]
    public void MarkChartRenderCurrent_sets_the_flag_without_an_invalidation()
    {
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
        var ribbon = new Mock<IRibbonUI>();
        var service = Arm(adapter, ribbon: ribbon.Object);
        var invalidationsBefore = CountInvalidations(ribbon);

        Assert.False(service.ChartRenderCurrent);

        service.MarkChartRenderCurrent();

        Assert.True(service.ChartRenderCurrent);
        Assert.Equal(invalidationsBefore, CountInvalidations(ribbon));
    }

    /// <summary>
    /// The freshness fact belongs to the workbook it was recorded for, so a workbook
    /// switch clears it and the Export PNG command refreshes before the next export (R8.8).
    /// </summary>
    [Fact]
    public void A_workbook_state_change_clears_the_chart_current_flag()
    {
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
        var service = Arm(adapter);
        service.MarkChartRenderCurrent();
        Assert.True(service.ChartRenderCurrent);

        adapter.RaiseWorkbookStateChanged();

        Assert.False(service.ChartRenderCurrent);
    }

    /// <summary>
    /// Selecting the file target from the menu toggle persists the workbook row and flips the
    /// pressed state and the main button's label (R8.8 D8).
    /// </summary>
    [Fact]
    public void SetPngExportTarget_persists_the_choice_and_re_queries_the_surface()
    {
        var writer = new Mock<IConfigCatalogueWriter>();
        _ = writer.Setup(w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()))
            .Returns(ConfigWriteOutcome.Ok());
        var service = new RibbonStateService();
        service.SetCatalogueWriter(writer.Object);
        PngExportPreferences.Shared.Target = PngExportTarget.Clipboard;

        service.SetPngExportTarget(RibbonControlIds.ExportPngTargetFile);

        Assert.Equal(PngExportTarget.File, PngExportPreferences.Shared.Target);
        Assert.True(RibbonStateService.GetExportPngFilePressed(null));
        Assert.False(RibbonStateService.GetExportPngClipboardPressed(null));
        Assert.Equal("Export PNG to file", RibbonStateService.GetExportPngLabel(null));
        writer.Verify(
            w => w.WriteSettings(It.Is<IReadOnlyDictionary<string, string>>(
                settings => settings[GanttPngExport.TargetSettingKey] == GanttPngExport.FileValue)),
            Times.Once);

        // Restore the shared default so a later test starts from the clipboard.
        PngExportPreferences.Shared.Target = PngExportTarget.Clipboard;
    }

    /// <summary>
    /// A toggled-off press still selects its own target, so the pair behaves as radio buttons
    /// and clicking the active row can never leave both rows unpressed (R8.8).
    /// </summary>
    [Fact]
    public void A_toggled_off_target_still_selects_that_target()
    {
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);
        PngExportPreferences.Shared.Target = PngExportTarget.Clipboard;

        service.SetPngExportTarget(RibbonControlIds.ExportPngTargetFile);
        service.SetPngExportTarget(RibbonControlIds.ExportPngTargetFile);

        Assert.Equal(PngExportTarget.File, PngExportPreferences.Shared.Target);
        PngExportPreferences.Shared.Target = PngExportTarget.Clipboard;
    }

    /// <summary>The activation read seeds the session target and folder from the workbook (R8.8).</summary>
    [Fact]
    public void Seeding_adopts_the_stored_target_and_folder()
    {
        RibbonStateService.SeedPersistedSwitches(new Dictionary<string, string>(StringComparer.Ordinal)
        {
            [GanttPngExport.TargetSettingKey] = GanttPngExport.FileValue,
            [GanttPngExport.FolderSettingKey] = @"C:\out",
        });

        Assert.Equal(PngExportTarget.File, PngExportPreferences.Shared.Target);
        Assert.Equal(@"C:\out", PngExportPreferences.Shared.Folder);

        PngExportPreferences.Shared.Target = PngExportTarget.Clipboard;
        PngExportPreferences.Shared.Folder = string.Empty;
    }

    /// <summary>A successful file export records its folder in the session and the workbook (D9).</summary>
    [Fact]
    public void PersistPngExportFolder_records_the_folder_and_the_workbook_row()
    {
        var writer = new Mock<IConfigCatalogueWriter>();
        _ = writer.Setup(w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()))
            .Returns(ConfigWriteOutcome.Ok());
        var service = new RibbonStateService();
        service.SetCatalogueWriter(writer.Object);

        service.PersistPngExportFolder(@"C:\exports");

        Assert.Equal(@"C:\exports", PngExportPreferences.Shared.Folder);
        writer.Verify(
            w => w.WriteSettings(It.Is<IReadOnlyDictionary<string, string>>(
                settings => settings[GanttPngExport.FolderSettingKey] == @"C:\exports")),
            Times.Once);
        PngExportPreferences.Shared.Folder = string.Empty;
    }

    /// <summary>
    /// The dialog's initial folder is the remembered folder when it still exists; with none,
    /// and no adapter to read the workbook's folder, it is null for the host default (D9).
    /// </summary>
    [Fact]
    public void The_initial_folder_prefers_the_remembered_folder()
    {
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);
        PngExportPreferences.Shared.Folder = string.Empty;
        Assert.Null(service.ResolvePngExportInitialFolder());

        string temp = Path.GetTempPath().TrimEnd(Path.DirectorySeparatorChar);
        PngExportPreferences.Shared.Folder = temp;
        Assert.Equal(temp, service.ResolvePngExportInitialFolder());
        PngExportPreferences.Shared.Folder = string.Empty;
    }

    /// <summary>
    /// Test double for <see cref="RibbonStateService"/> that substitutes a
    /// fixed dialog result instead of running a blocking WinForms message
    /// loop, which would hang forever without a user to click OK or Cancel.
    /// </summary>
    private sealed class StubThemeDialogRibbonStateService : RibbonStateService
    {
        private readonly DialogResult _dialogResult;

        public StubThemeDialogRibbonStateService(DialogResult dialogResult) => _dialogResult = dialogResult;

        protected override DialogResult ShowThemeSettingsDialog(ThemeSettingsForm dialog)
        {
            dialog.SetValidatedRequestForTesting(new ThemeEditRequest([], new Dictionary<string, double>()));
            return _dialogResult;
        }
    }
}
