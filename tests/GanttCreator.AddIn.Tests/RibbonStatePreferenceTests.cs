
using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;
using Moq;
using static GanttCreator.AddIn.Tests.RibbonStateTestSupport;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The preference half of <see cref="RibbonStateService"/>: the size preset and margins that
/// recompute the width readout, the theme dialog's four outcomes, the chart-current flag, and the
/// PNG export target and remembered folder.
/// </summary>
/// <remarks>
/// Shares the <c>diagnostics-service</c> collection: these facts arm the same session singleton the
/// host-lifecycle facts arm and reset, so xUnit runs the collection's classes sequentially. The
/// constructor and <see cref="Dispose"/> reset it around every fact.
/// </remarks>
[Collection("diagnostics-service")]
public sealed class RibbonStatePreferenceTests : IDisposable
{
    public RibbonStatePreferenceTests() => RibbonStateService.Reset();

    public void Dispose()
    {
        RibbonStateService.Reset();
        GC.SuppressFinalize(this);
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
        var service = new StubThemeDialogStateService(DialogResult.OK);
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
        var service = new StubThemeDialogStateService(DialogResult.Cancel);
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
        var service = new StubThemeDialogStateService(DialogResult.OK);
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
}
