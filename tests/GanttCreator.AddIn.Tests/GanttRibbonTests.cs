using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using ExcelDna.Integration.CustomUI;
using GanttCreator.Core.Logging;
using GanttCreator.Office;
using Moq;

// CA2000: test fixtures are deliberately not disposed — the code under test
// (GanttRibbon) does not own disposable resources, and these are plain strings
// / reflection results. Scoped to this file.
#pragma warning disable CA2000

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contract tests for <see cref="GanttRibbon"/>: the Excel-DNA reflection
/// contract (public, ComVisible, ExcelRibbon-derived, parameterless
/// constructor), the <see cref="GetCustomUI"/> behaviour, and the RibbonX
/// namespace / IDs / callback contract. None of these require Excel.
/// </summary>
public class GanttRibbonTests
{
    private const string WorkbookRibbonId = "Microsoft.Excel.Workbook";

    private static readonly GanttRibbon Ribbon = new();

    // The set of RibbonX attribute names that dispatch to a callback method on
    // the ribbon class. Any of these appearing in the document must resolve to a
    // public method with a compatible signature (the callback contract).
    private static readonly HashSet<string> CallbackAttributeNames =
    [
        "onLoad",
        "onAction",
        "onChange",
        "getPressed",
        "getEnabled",
        "getVisible",
        "getImage",
        "getLabel",
        "getSize",
        "getScreentip",
        "getSupertip",
        "getDescription",
        "getKeyTip",
        "getText",
        "getContent",
        "getShowLabel",
        "getShowImage",
    ];

    private const string NamespaceCustomUI2010 =
        "http://schemas.microsoft.com/office/2009/07/customui";

    [Fact]
    public void GanttRibbon_is_public_ComVisible_ExcelRibbon_derived_and_has_parameterless_ctor()
    {
        Type type = typeof(GanttRibbon);

        Assert.True(type.IsPublic, "Excel-DNA requires a public ribbon type.");
        ComVisibleAttribute? ComVisibleAttribute =
            type.GetCustomAttribute<ComVisibleAttribute>(inherit: true);
        Assert.True(
            ComVisibleAttribute is not null && ComVisibleAttribute.Value == true,
            "The ribbon type must be [ComVisible(true)] for COM registration.");
        Assert.True(
            typeof(ExcelRibbon).IsAssignableFrom(type),
            "Excel-DNA only auto-registers ExcelRibbon descendants (IsRibbonType).");
        System.Reflection.ConstructorInfo? ctor = type.GetConstructor(Type.EmptyTypes);
        Assert.NotNull(ctor);
        Assert.True(
            ctor!.IsPublic,
            "Excel-DNA instantiates the ribbon via Activator.CreateInstance, which requires a public parameterless constructor.");
    }

    [Fact]
    public void GetCustomUI_returns_the_ribbon_xml_for_the_workbook_id()
    {
        string? xml = Ribbon.GetCustomUI(WorkbookRibbonId);

        Assert.NotNull(xml);
        Assert.NotEmpty(xml);
        Assert.Contains("customUI", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void GetCustomUI_returns_null_for_any_other_ribbon_id()
    {
        Assert.Null(Ribbon.GetCustomUI("Microsoft.Excel.Dataset"));
        Assert.Null(Ribbon.GetCustomUI(string.Empty));
        Assert.Null(Ribbon.GetCustomUI("Other"));
    }

    [Fact]
    public void Ribbon_xml_parses_and_uses_the_office_2009_namespace()
    {
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;

        XDocument doc = XDocument.Parse(xml);
        XElement? root = doc.Root;
        Assert.NotNull(root);
        Assert.Equal("customUI", root.Name.LocalName, ignoreCase: false);

        // The 2009/07 namespace targets Excel 2010+, the supported baseline.
        string ns = root.Name.NamespaceName;
        Assert.Equal(NamespaceCustomUI2010, ns);
    }

    [Fact]
    public void Ribbon_declares_exactly_one_gantt_creator_tab_with_expected_id_and_label()
    {
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);

        // Elements live in the CustomUI 2009/07 default namespace, so queries
        // must be namespace-qualified to match.
        XNamespace ns = NamespaceCustomUI2010;
        List<XElement> tabs = doc.Descendants(ns + "tab").ToList();
        Assert.Single(tabs);

        XElement tab = tabs[0];
        Assert.Equal(
            "tabGanttCreator",
            tab.Attribute("id")?.Value,
            ignoreCase: false);
        Assert.Equal(
            "Gantt Creator",
            tab.Attribute("label")?.Value,
            ignoreCase: false);
    }

    [Fact]
    public void Callback_contract_all_callbacks_in_the_real_xml_resolve_to_public_methods()
    {
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;

        List<string> unresolved = ValidateCallbacks(xml, typeof(GanttRibbon));

        Assert.Empty(unresolved);
    }

    [Fact]
    public void Callback_contract_reports_a_missing_callback_method_positive_test()
    {
        // A deliberately broken document whose onAction references a method that
        // does not exist on the ribbon type. Proves the validator actually
        // catches violations rather than vacuously passing.
        const string brokenXml = @"<?xml version='1.0' encoding='utf-8'?>
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='tabGanttCreator' label='Gantt Creator'>
        <group id='grpPlaceholder' label='Gantt Creator'>
          <button id='btnBroken' label='Broken' onAction='DoesNotExist'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        List<string> unresolved = ValidateCallbacks(brokenXml, typeof(GanttRibbon));

        Assert.Contains("DoesNotExist", unresolved);
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

    [Fact]
    public void Ribbon_declares_getEnabled_on_exactly_the_two_gated_controls()
    {
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;

        var gatedIds = doc.Descendants(ns + "button")
            .Where(button => button.Attribute("getEnabled") is not null)
            .Select(button => button.Attribute("id")?.Value)
            .ToList();

        Assert.Equal(
            [RibbonControlIds.Diagnostics, RibbonControlIds.OpenLog],
            gatedIds);
    }

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
    public void Every_ribbon_command_refreshes_the_ribbon_state_after_the_boundary_run()
    {
        // Work item R1.5 decision D3: every command run through the ribbon's
        // boundary ends with a ribbon-state refresh — once on success and once
        // when the command throws (the boundary absorbs the failure).
        var diagnosticsRefreshes = 0;
        var openLogRefreshes = 0;
        var boundary = new CommandBoundary(presenter: _ => { });

        GanttRibbon.OnDiagnosticsClick(null, boundary, () => { }, () => diagnosticsRefreshes++);
        GanttRibbon.OnDiagnosticsClick(
            null, boundary, () => throw new InvalidOperationException("simulated"), () => diagnosticsRefreshes++);
        GanttRibbon.OnOpenLogClick(null, boundary, () => { }, () => openLogRefreshes++);
        GanttRibbon.OnOpenLogClick(
            null, boundary, () => throw new InvalidOperationException("simulated"), () => openLogRefreshes++);

        Assert.Equal(2, diagnosticsRefreshes);
        Assert.Equal(2, openLogRefreshes);
    }

    /// <summary>
    /// Parses the RibbonX <paramref name="xml"/> and returns the callback method
    /// names (the values of known callback attributes) that do not resolve to a
    /// public instance method on <paramref name="ribbonType"/> with a compatible
    /// signature: zero parameters, or one parameter of type
    /// <see cref="IRibbonControl"/> or <see cref="IRibbonUI"/>.
    /// </summary>
    /// <param name="xml">The RibbonX document to validate.</param>
    /// <param name="ribbonType">The ribbon class type.</param>
    /// <returns>The set of unresolved callback method names, if any.</returns>
    private static List<string> ValidateCallbacks(
        string xml, Type ribbonType)
    {
        XDocument doc = XDocument.Parse(xml);
        List<string> unresolved = [];

        foreach (XElement element in doc.Descendants())
        {
            foreach (XAttribute attribute in element.Attributes())
            {
                if (!CallbackAttributeNames.Contains(attribute.Name.LocalName))
                {
                    continue;
                }

                string methodName = attribute.Value;
                if (!CallbackResolves(ribbonType, methodName))
                {
                    unresolved.Add(methodName);
                }
            }
        }

        return unresolved;
    }

    private static bool CallbackResolves(Type ribbonType, string methodName)
    {
        System.Reflection.MethodInfo? method = ribbonType.GetMethod(
            methodName,
            BindingFlags.Instance | BindingFlags.Public);

        if (method is null)
        {
            return false;
        }

        System.Reflection.ParameterInfo[] parameters = method.GetParameters();
        if (parameters.Length == 0)
        {
            return true;
        }

        if (parameters.Length == 1)
        {
            Type paramType = parameters[0].ParameterType;
            return paramType == typeof(IRibbonControl)
                || paramType == typeof(IRibbonUI);
        }

        return false;
    }
}
