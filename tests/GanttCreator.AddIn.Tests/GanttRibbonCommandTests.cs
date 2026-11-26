using System.Globalization;
using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
using GanttCreator.Core.Logging;
using GanttCreator.Office;
using Moq;
using static GanttCreator.AddIn.Tests.RibbonTestSupport;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The ribbon's command callbacks: what each click routes through the command boundary, what it
/// tells the ribbon state afterwards, and what it refuses before reaching a command.
/// </summary>
/// <remarks>
/// These facts assert the <em>routing</em> contract only. A command's own behaviour lives with the
/// command's tests, so a failure here points at the ribbon surface rather than at an export.
/// </remarks>
[Collection("diagnostics-service")]
public sealed class GanttRibbonCommandTests
{
    [Fact]
    public void OnOpenLogClick_routes_through_the_command_boundary()
    {
        var shown = new List<string>();
        var written = new List<string>();
        var log = new Mock<IRollingLog>();
        log
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>(
                (fmt, args) => written.Add(string.Format(CultureInfo.InvariantCulture, fmt, args)));
        var boundary = new CommandBoundary(presenter: shown.Add);
        boundary.SetLog(log.Object);
        var executed = false;

        GanttRibbon.OnOpenLogClick(null, boundary, () => executed = true);

        Assert.True(executed, "The command must run.");
        Assert.Empty(written);
        Assert.Empty(shown);
    }

    [Fact]
    public void OnOpenLogClick_failure_produces_one_record_one_dialog_and_the_fallback_name()
    {
        var shown = new List<string>();
        var written = new List<string>();
        var log = new Mock<IRollingLog>();
        log
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>(
                (fmt, args) => written.Add(string.Format(CultureInfo.InvariantCulture, fmt, args)));
        var boundary = new CommandBoundary(presenter: shown.Add);
        boundary.SetLog(log.Object);

        GanttRibbon.OnOpenLogClick(null, boundary, () => throw new InvalidOperationException("simulated"));

        var record = Assert.Single(written);
        Assert.Contains("CommandError", record, StringComparison.Ordinal);
        Assert.Contains("command=OnOpenLogClick", record, StringComparison.Ordinal);
        Assert.Single(shown);
    }

    [Fact]
    public void OnInitialiseSheetClick_routes_through_the_command_boundary()
    {
        var shown = new List<string>();
        var written = new List<string>();
        var log = new Mock<IRollingLog>();
        log
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>(
                (fmt, args) => written.Add(string.Format(CultureInfo.InvariantCulture, fmt, args)));
        var boundary = new CommandBoundary(presenter: shown.Add);
        boundary.SetLog(log.Object);
        var executed = false;

        GanttRibbon.OnInitialiseSheetClick(null, boundary, () => executed = true);

        Assert.True(executed, "The initialise command must run.");
        Assert.Empty(written);
        Assert.Empty(shown);
    }

    [Fact]
    public void OnInitialiseSheetClick_failure_produces_one_record_one_dialog_and_the_fallback_name()
    {
        var shown = new List<string>();
        var written = new List<string>();
        var log = new Mock<IRollingLog>();
        log
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>(
                (fmt, args) => written.Add(string.Format(CultureInfo.InvariantCulture, fmt, args)));
        var boundary = new CommandBoundary(presenter: shown.Add);
        boundary.SetLog(log.Object);

        GanttRibbon.OnInitialiseSheetClick(
            null, boundary, () => throw new InvalidOperationException("simulated"));

        var record = Assert.Single(written);
        Assert.Contains("CommandError", record, StringComparison.Ordinal);
        Assert.Contains("command=OnInitialiseSheetClick", record, StringComparison.Ordinal);
        Assert.Single(shown);
    }

    [Fact]
    public void OnInitialiseSheetClick_uses_the_control_id_as_the_command_name_when_available()
    {
        // The boundary records the stable command name; for Initialise sheet the
        // control ID is btnInitialiseSheet, not the fallback method name.
        var written = new List<string>();
        var log = new Mock<IRollingLog>();
        log
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>(
                (fmt, args) => written.Add(string.Format(CultureInfo.InvariantCulture, fmt, args)));
        var boundary = new CommandBoundary(presenter: _ => { });
        boundary.SetLog(log.Object);

        var control = new Mock<IRibbonControl>();
        control.SetupGet(c => c.Id).Returns(RibbonControlIds.InitialiseSheet);
        GanttRibbon.OnInitialiseSheetClick(
            control.Object, boundary, () => throw new InvalidOperationException("simulated"));

        var record = Assert.Single(written);
        Assert.Contains("command=btnInitialiseSheet", record, StringComparison.Ordinal);
    }

    [Fact]
    public void Every_ribbon_command_refreshes_the_ribbon_state_after_the_boundary_run()
    {
        // Work item R1.5 decision D3: every command run through the ribbon's
        // boundary ends with a ribbon-state refresh — once on success and once
        // when the command throws (the boundary absorbs the failure). R2.2's
        // Initialise sheet command is no exception: it refreshes on both paths.
        var diagnosticsRefreshes = 0;
        var openLogRefreshes = 0;
        var initialiseRefreshes = 0;
        var validateRefreshes = 0;
        var boundary = new CommandBoundary(presenter: _ => { });

        GanttRibbon.OnDiagnosticsClick(null, boundary, () => { }, () => diagnosticsRefreshes++);
        GanttRibbon.OnDiagnosticsClick(
            null, boundary, () => throw new InvalidOperationException("simulated"), () => diagnosticsRefreshes++);
        GanttRibbon.OnOpenLogClick(null, boundary, () => { }, () => openLogRefreshes++);
        GanttRibbon.OnOpenLogClick(
            null, boundary, () => throw new InvalidOperationException("simulated"), () => openLogRefreshes++);
        GanttRibbon.OnInitialiseSheetClick(null, boundary, () => { }, () => initialiseRefreshes++);
        GanttRibbon.OnInitialiseSheetClick(
            null, boundary, () => throw new InvalidOperationException("simulated"), () => initialiseRefreshes++);
        GanttRibbon.OnValidateSheetClick(null, boundary, () => { }, () => validateRefreshes++);
        GanttRibbon.OnValidateSheetClick(
            null, boundary, () => throw new InvalidOperationException("simulated"), () => validateRefreshes++);

        Assert.Equal(2, diagnosticsRefreshes);
        Assert.Equal(2, openLogRefreshes);
        Assert.Equal(2, initialiseRefreshes);
        Assert.Equal(2, validateRefreshes);
    }

    [Fact]
    public void OnValidateSheetClick_routes_through_the_command_boundary()
    {
        var shown = new List<string>();
        var written = new List<string>();
        var log = new Mock<IRollingLog>();
        log
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>(
                (fmt, args) => written.Add(string.Format(CultureInfo.InvariantCulture, fmt, args)));
        var boundary = new CommandBoundary(presenter: shown.Add);
        boundary.SetLog(log.Object);
        var executed = false;

        GanttRibbon.OnValidateSheetClick(null, boundary, () => executed = true);

        Assert.True(executed, "The validate command must run.");
        Assert.Empty(written);
        Assert.Empty(shown);
    }

    [Fact]
    public void OnValidateSheetClick_failure_produces_one_record_one_dialog_and_the_fallback_name()
    {
        var shown = new List<string>();
        var written = new List<string>();
        var log = new Mock<IRollingLog>();
        log
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>(
                (fmt, args) => written.Add(string.Format(CultureInfo.InvariantCulture, fmt, args)));
        var boundary = new CommandBoundary(presenter: shown.Add);
        boundary.SetLog(log.Object);

        GanttRibbon.OnValidateSheetClick(
            null, boundary, () => throw new InvalidOperationException("simulated"));

        var record = Assert.Single(written);
        Assert.Contains("CommandError", record, StringComparison.Ordinal);
        Assert.Contains("command=OnValidateSheetClick", record, StringComparison.Ordinal);
        Assert.Single(shown);
    }

    [Fact]
    public void OnValidateSheetClick_uses_the_control_id_as_the_command_name_when_available()
    {
        var written = new List<string>();
        var log = new Mock<IRollingLog>();
        log
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>(
                (fmt, args) => written.Add(string.Format(CultureInfo.InvariantCulture, fmt, args)));
        var boundary = new CommandBoundary(presenter: _ => { });
        boundary.SetLog(log.Object);

        var control = new Mock<IRibbonControl>();
        control.SetupGet(c => c.Id).Returns(RibbonControlIds.ValidateSheet);
        GanttRibbon.OnValidateSheetClick(
            control.Object, boundary, () => throw new InvalidOperationException("simulated"));

        var record = Assert.Single(written);
        Assert.Contains("command=btnValidateSheet", record, StringComparison.Ordinal);
    }

    /// <summary>
    /// R4.9: the Refresh callback routes across the one boundary, names the command
    /// it ran, and refreshes the ribbon state afterwards.
    /// </summary>
    /// <remarks>
    /// The command is injected as a throwing delegate so this test exercises the
    /// <em>routing</em> — that the exception crossed <see cref="CommandBoundary"/>
    /// rather than escaping the callback — without a real workbook behind it. A
    /// separate test covers the command's own behaviour.
    /// </remarks>
    [Fact]
    public void OnRefreshSheetClick_routes_across_the_boundary_and_refreshes_state()
    {
        var written = new List<string>();
        var log = new Mock<IRollingLog>();
        log
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>(
                (fmt, args) => written.Add(string.Format(CultureInfo.InvariantCulture, fmt, args)));
        var boundary = new CommandBoundary(presenter: _ => { });
        boundary.SetLog(log.Object);

        var control = new Mock<IRibbonControl>();
        control.SetupGet(c => c.Id).Returns(RibbonControlIds.RefreshSheet);
        var stateRefreshed = 0;

        GanttRibbon.OnRefreshSheetClick(
            control.Object,
            boundary,
            () => throw new InvalidOperationException("simulated"),
            () => stateRefreshed++);

        var record = Assert.Single(written);
        Assert.Contains("command=btnRefreshSheet", record, StringComparison.Ordinal);

        // The state refresh is what re-evaluates every getEnabled, so a button that
        // changed availability during the command must be re-queried afterwards.
        Assert.Equal(1, stateRefreshed);
    }

    /// <summary>Null dependencies are refused at the boundary, not at first use.</summary>
    [Fact]
    public void OnRefreshSheetClick_throws_for_a_null_boundary_or_command()
    {
        var written = new List<string>();
        var log = new Mock<IRollingLog>();
        log
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>(
                (fmt, args) => written.Add(string.Format(CultureInfo.InvariantCulture, fmt, args)));
        var boundary = new CommandBoundary(presenter: _ => { });
        boundary.SetLog(log.Object);

        var control = new Mock<IRibbonControl>();

        Assert.Throws<ArgumentNullException>(
            () => GanttRibbon.OnRefreshSheetClick(control.Object, null!, () => { }));
        Assert.Throws<ArgumentNullException>(
            () => GanttRibbon.OnRefreshSheetClick(control.Object, boundary, null!));
    }

    [Fact]
    public void OnValidateSheetClick_throws_for_a_null_boundary_or_command()
    {
        Assert.IsType<ArgumentNullException>(
            Record.Exception(() => GanttRibbon.OnValidateSheetClick(null, null!, () => { })));
        Assert.IsType<ArgumentNullException>(
            Record.Exception(() => GanttRibbon.OnValidateSheetClick(
                null, new CommandBoundary(presenter: _ => { }), null!)));
    }
    [Fact]
    public void OnThemeSettingsClick_throws_for_null_state_service()
    {
        var control = new Mock<IRibbonControl>();
        Assert.Throws<ArgumentNullException>(
            () => GanttRibbon.OnThemeSettingsClick(control.Object, null!));
    }

    [Fact]
    public void OnThemeSettingsClick_does_not_throw_when_dialog_returns_false()
    {
        // When ShowThemeSettingsDialog returns false (write refused or cancelled),
        // RunRefreshSheet is not called, so no exception should occur.
        var service = new RibbonStateService();
        var reader = new Mock<IConfigCatalogueReader>();
        reader.Setup(r => r.Read()).Returns(ConfigReadOutcome.Refused(ConfigReadRefusalReason.ConfigSheetMissing));
        service.SetCatalogueReader(reader.Object);

        var writer = new Mock<IConfigCatalogueWriter>();
        service.SetCatalogueWriter(writer.Object);

        var control = new Mock<IRibbonControl>();
        Record.Exception(() => GanttRibbon.OnThemeSettingsClick(control.Object, service));
    }

    [Fact]
    public void OnThemeSettingsClick_notifies_the_ribbon_state_after_a_successful_write()
    {
        // R5.6a: the Theme settings dialog can create a custom style, and the
        // ribbon's custom-style dropdown projects its items from the workbook's
        // style registry. The command that wrote the style must therefore
        // re-capture and invalidate the ribbon state, or the dropdown keeps the
        // list it had before the dialog opened and the new style never appears.
        var service = new StubThemeDialogStateService(DialogResult.OK);
        var reader = new Mock<IConfigCatalogueReader>();
        reader.Setup(r => r.Read()).Returns(ConfigReadOutcome.Ok(
            "wb-id",
            new Dictionary<string, string>(),
            GanttStyleRegistry.Empty,
            new Dictionary<string, double>()));
        service.SetCatalogueReader(reader.Object);

        var writer = new Mock<IConfigCatalogueWriter>();
        writer.Setup(w => w.WriteTheme(It.IsAny<ThemeEditRequest>())).Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);

        var completed = 0;
        var refreshed = 0;
        var control = new Mock<IRibbonControl>();

        GanttRibbon.OnThemeSettingsClick(
            control.Object,
            service,
            onCompleted: () => completed++,
            refresh: () => refreshed++);

        Assert.Equal(1, refreshed);
        Assert.Equal(1, completed);
    }

    [Fact]
    public void OnThemeSettingsClick_does_not_notify_the_ribbon_state_when_the_dialog_is_cancelled()
    {
        var service = new StubThemeDialogStateService(DialogResult.Cancel);
        var reader = new Mock<IConfigCatalogueReader>();
        reader.Setup(r => r.Read()).Returns(ConfigReadOutcome.Ok(
            "wb-id",
            new Dictionary<string, string>(),
            GanttStyleRegistry.Empty,
            new Dictionary<string, double>()));
        service.SetCatalogueReader(reader.Object);

        var writer = new Mock<IConfigCatalogueWriter>();
        service.SetCatalogueWriter(writer.Object);

        var completed = 0;
        var refreshed = 0;
        var control = new Mock<IRibbonControl>();

        GanttRibbon.OnThemeSettingsClick(
            control.Object,
            service,
            onCompleted: () => completed++,
            refresh: () => refreshed++);

        Assert.Equal(0, refreshed);
        Assert.Equal(0, completed);
    }

    private sealed class ThrowingVisibleStateService : RibbonStateService
    {
        internal override bool GetVisible(string? groupId) =>
            throw new InvalidOperationException("probe failed");
    }

    [Fact]
    public void Get_visible_fails_closed_when_the_state_service_throws()
    {
        var control = new Mock<IRibbonControl>();
        _ = control.SetupGet(c => c.Id).Returns(RibbonControlIds.DataGroup);

        Assert.False(GanttRibbon.GetVisible(control.Object, new ThrowingVisibleStateService()));
    }

    [Fact]
    public void GetCustomUI_returns_null_for_any_other_ribbon_id()
    {
        Assert.Null(Ribbon.GetCustomUI("Microsoft.Excel.Dataset"));
        Assert.Null(Ribbon.GetCustomUI(string.Empty));
        Assert.Null(Ribbon.GetCustomUI("Other"));
    }

    [Fact]
    public void OnDiagnosticsClick_routes_through_the_command_boundary()
    {
        var shown = new List<string>();
        var written = new List<string>();
        var log = new Mock<IRollingLog>();
        log
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>(
                (fmt, args) => written.Add(string.Format(CultureInfo.InvariantCulture, fmt, args)));
        var boundary = new CommandBoundary(presenter: shown.Add);
        boundary.SetLog(log.Object);

        GanttRibbon.OnDiagnosticsClick(null, boundary, () => throw new InvalidOperationException("simulated"));

        var record = Assert.Single(written);
        Assert.Contains("CommandError", record, StringComparison.Ordinal);
        Assert.Single(shown);
    }

    [Fact]
    public void OnDiagnosticsClick_success_runs_command_without_dialog()
    {
        var shown = new List<string>();
        var written = new List<string>();
        var log = new Mock<IRollingLog>();
        log
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>(
                (fmt, args) => written.Add(string.Format(CultureInfo.InvariantCulture, fmt, args)));
        var boundary = new CommandBoundary(presenter: shown.Add);
        boundary.SetLog(log.Object);
        var executed = false;

        GanttRibbon.OnDiagnosticsClick(null, boundary, () => executed = true);

        Assert.True(executed, "The command must run.");
        Assert.Empty(written);
        Assert.Empty(shown);
    }

    [Fact]
    public void OnDiagnosticsClick_probe_failure_uses_fallback_command_name()
    {
        var shown = new List<string>();
        var written = new List<string>();
        var log = new Mock<IRollingLog>();
        log
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>(
                (fmt, args) => written.Add(string.Format(CultureInfo.InvariantCulture, fmt, args)));
        var boundary = new CommandBoundary(presenter: shown.Add);
        boundary.SetLog(log.Object);
        var control = new Mock<IRibbonControl>();
        control.SetupGet(c => c.Id).Throws(new InvalidOperationException("probe broken"));

        GanttRibbon.OnDiagnosticsClick(control.Object, boundary, () => throw new InvalidOperationException("simulated"));

        var record = Assert.Single(written);
        Assert.Contains("command=OnDiagnosticsClick", record, StringComparison.Ordinal);
        Assert.Single(shown);
    }

    [Fact]
    public void OnRepairConfigClick_routes_through_the_command_boundary()
    {
        var written = new List<string>();
        var log = new Mock<IRollingLog>();
        log
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>((format, args) =>
                written.Add(string.Format(CultureInfo.InvariantCulture, format, args)));
        var boundary = new CommandBoundary(presenter: _ => { });
        boundary.SetLog(log.Object);
        var executed = false;

        GanttRibbon.OnRepairConfigClick(null, boundary, () => executed = true);

        Assert.True(executed);
        Assert.Empty(written);
    }

    [Fact]
    public void OnRepairConfigClick_failure_routes_to_the_boundary()
    {
        var shown = new List<string>();
        var log = new Mock<IRollingLog>();
        log
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>((format, args) => { });
        var boundary = new CommandBoundary(presenter: shown.Add);
        boundary.SetLog(log.Object);

        GanttRibbon.OnRepairConfigClick(null, boundary, () => throw new InvalidOperationException("simulated"));

        Assert.Single(shown);
    }

    [Fact]
    public void ResolveCommandName_uses_control_id_and_falls_back_to_method_name()
    {
        var control = new Mock<IRibbonControl>();
        control.SetupGet(c => c.Id).Returns("btnDiagnostics");

        Assert.Equal("btnDiagnostics", GanttRibbon.ResolveCommandName(control.Object));
        Assert.Equal("OnDiagnosticsClick", GanttRibbon.ResolveCommandName(null));

        control.SetupGet(c => c.Id).Returns(string.Empty);
        Assert.Equal("OnDiagnosticsClick", GanttRibbon.ResolveCommandName(control.Object));

        control.SetupGet(c => c.Id).Returns("   ");
        Assert.Equal("OnOpenLogClick", GanttRibbon.ResolveCommandName(control.Object, "OnOpenLogClick"));
    }

    [Fact]
    public void OnLoad_publishes_the_ribbon_handle_and_invalidates_once()
    {
        var stateService = new RibbonStateService();
        var ribbon = new Mock<IRibbonUI>();

        GanttRibbon.OnLoad(ribbon.Object, stateService);

        Assert.Same(ribbon.Object, stateService.GetRibbon());
        Assert.Equal(
            1,
            ribbon.Invocations.Count(invocation => invocation.Method.Name == nameof(IRibbonUI.Invalidate)));
    }

    [Fact]
    public void OnLoad_with_a_null_handle_never_throws_and_publishes_null()
    {
        var stateService = new RibbonStateService();

        var exception = Record.Exception(() => GanttRibbon.OnLoad(null, stateService));

        Assert.Null(exception);
        Assert.Null(stateService.GetRibbon());
    }

    [Fact]
    public void GetEnabled_routes_through_the_state_service()
    {
        var adapter = new Mock<IExcelApplicationAdapter>();
        adapter.Setup(a => a.HasActiveWorkbook()).Returns(true);
        var stateService = new RibbonStateService();
        stateService.SetApplicationAdapter(adapter.Object);
        stateService.SetLogAvailabilitySource(() => true);
        stateService.Refresh();

        var diagnostics = new Mock<IRibbonControl>();
        diagnostics.SetupGet(c => c.Id).Returns(RibbonControlIds.Diagnostics);
        var openLog = new Mock<IRibbonControl>();
        openLog.SetupGet(c => c.Id).Returns(RibbonControlIds.OpenLog);

        Assert.True(GanttRibbon.GetEnabled(diagnostics.Object, stateService));
        Assert.True(GanttRibbon.GetEnabled(openLog.Object, stateService));

        // The getter must be a pure snapshot read: no second probe of the
        // state source beyond the refresh capture.
        adapter.Verify(a => a.HasActiveWorkbook(), Times.Once);
    }

    [Fact]
    public void GetEnabled_stays_enabled_when_the_control_id_probe_fails()
    {
        // Both facts false: only the fail-open probe path can return true.
        var stateService = new RibbonStateService();
        stateService.SetApplicationAdapter(new Mock<IExcelApplicationAdapter>().Object);
        stateService.SetLogAvailabilitySource(() => false);
        stateService.Refresh();

        var control = new Mock<IRibbonControl>();
        control.SetupGet(c => c.Id).Throws(new InvalidOperationException("probe failed"));

        Assert.True(GanttRibbon.GetEnabled(control.Object, stateService));
    }
}
