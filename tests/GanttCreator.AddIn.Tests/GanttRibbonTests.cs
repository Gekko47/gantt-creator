using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
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

    [Fact]
    public void Ribbon_declares_getEnabled_on_exactly_the_four_gated_controls()
    {
        // R2.2 adds the Initialise sheet button to the Data group and R2.6 adds
        // Validate beside it; both are gated on the same workbook fact as
        // Diagnostics. The set is pinned here so adding or dropping a gated
        // control fails the contract instead of silently changing the surface.
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;

        var gatedIds = doc.Descendants(ns + "button")
            .Where(button => button.Attribute("getEnabled") is not null)
            .Select(button => button.Attribute("id")?.Value)
            .ToList();

        Assert.Equal(
            [
                RibbonControlIds.InitialiseSheet,
                RibbonControlIds.ValidateSheet,
                RibbonControlIds.RepairConfig,
                RibbonControlIds.AddActivity,
                RibbonControlIds.AddMilestone,
                RibbonControlIds.AddDelineator,
                RibbonControlIds.RefreshSheet,
                RibbonControlIds.Diagnostics,
                RibbonControlIds.OpenLog,
            ],
            gatedIds);
    }

    [Fact]
    public void Ribbon_declares_getEnabled_on_each_control_exactly_once()
    {
        // Each gated control declares getEnabled exactly once: a duplicated
        // attribute on one control would make the count above pass while the
        // ribbon rendered with a duplicate callback.
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;

        foreach (string id in new[]
        {
            RibbonControlIds.Diagnostics,
            RibbonControlIds.OpenLog,
            RibbonControlIds.InitialiseSheet,
            RibbonControlIds.ValidateSheet,
            RibbonControlIds.RepairConfig,
            RibbonControlIds.AddActivity,
            RibbonControlIds.AddMilestone,
            RibbonControlIds.AddDelineator,
            RibbonControlIds.RefreshSheet,
        })
        {
            var matches = doc.Descendants(ns + "button")
                .Where(b => b.Attribute("id")?.Value == id)
                .Select(b => b.Attribute("getEnabled"))
                .Count(attr => attr is not null);

            Assert.Equal(1, matches);
        }
    }

    [Fact]
    public void Ribbon_buttons_use_gallery_verified_imageMso_ids()
    {
        // imageMso fails silently on unknown IDs (no error, no log): the
        // Diagnostics button showed no icon with imageMso="Information",
        // which is absent from the Office 2010 Icons Gallery workbook
        // (customUI14.xml: 7344 unique IDs, no "Information"). Pin
        // evidence-backed IDs so a typo regresses to a failing test
        // instead of a silently missing icon.
        //
        // Two bases, both required in practice: the ID must be present in
        // the Office 2010 Icons Gallery (Office2010IconsGallery.docx, whose
        // embedded customUI14.xml lists 7344 unique imageMso IDs) and name a
        // command in the Excel idMso table of [MS-CUSTOMUI]-250218
        // (`.microsoft/` reference, git-ignored).
        // imageMso="OfficeDiagnostics" was a valid gallery ID but absent
        // from the Excel table and rendered nothing in F5. FileOpen and Help
        // satisfy both and render (Help glyph confirmed by a human operator
        // in F5, 2026-09-16).
        //
        // TableInsertExcel satisfies both: present in the gallery, and the
        // Excel idMso row idMso=TableInsertExcel, control=button,
        // label="Table" — the Excel Insert-Table command, the semantically
        // exact match for "Initialise sheet", which creates tblGanttData.
        // Note the id is TableInsertExcel, not its UI label "Table", which
        // is itself absent from the gallery. The gallery check is validated
        // against the four documented outcomes that file reproduces
        // (Information absent, OfficeDiagnostics present, Help present,
        // FileOpen present). Evidence and the remaining F5 glyph step:
        // docs/work-items/R2.2-initialise-sheet.md decision D8.
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;

        string? ImageMso(string id) =>
            doc.Descendants(ns + "button")
                .First(b => b.Attribute("id")?.Value == id)
                .Attribute("imageMso")?.Value;

        Assert.Equal("Help", ImageMso(RibbonControlIds.Diagnostics));
        Assert.Equal("FileOpen", ImageMso(RibbonControlIds.OpenLog));
        Assert.Equal("TableInsertExcel", ImageMso(RibbonControlIds.InitialiseSheet));
        Assert.Equal("Spelling", ImageMso(RibbonControlIds.ValidateSheet));
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

    /// <summary>
    /// The Plot Area group owns no Apply control: the Apply-button
    /// contract is replaced by per-end persistence, and the ribbon state
    /// does not consult an obsolete control ID.
    /// </summary>
    [Fact]
    public void PlotArea_group_has_no_apply_button()
    {
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;

        var buttonIds = doc.Descendants(ns + "button")
            .Select(button => button.Attribute("id")?.Value)
            .ToList();

        Assert.DoesNotContain("btnApplyPlotSettings", buttonIds, StringComparer.Ordinal);
        Assert.DoesNotContain("btnPlotSettingsAuto", buttonIds, StringComparer.Ordinal);
        Assert.DoesNotContain("btnPlotSettingsExplicit", buttonIds, StringComparer.Ordinal);
    }

    [Fact]
    public void PlotArea_group_declares_start_and_finish_rows_with_checkboxes_and_editboxes()
    {
        // Fix plan ruling 3: the flat label stack is restructured into proper
        // start/finish rows — one AUTO box row and one date box row per end,
        // each label sitting beside its control. The edit boxes carry getText
        // for the effective display value, getEnabled for the auto toggle, and
        // onChange for the commit path.
        string xml = Ribbon.GetCustomUI(WorkbookRibbonId)!;
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = NamespaceCustomUI2010;

        XElement group = doc.Descendants(ns + "group")
            .Single(element => element.Attribute("id")?.Value == "grpPlotArea");
        Assert.Equal("Plot Area", group.Attribute("label")?.Value);

        var labelIds = group.Descendants(ns + "labelControl")
            .Select(element => element.Attribute("id")?.Value)
            .ToList();
        Assert.Equal(
            ["lblPlotStartDate", "lblPlotStartAuto", "lblPlotFinishDate", "lblPlotFinishAuto"],
            labelIds);

        XElement startCheck = group.Descendants(ns + "checkBox")
            .Single(element => element.Attribute("id")?.Value == "chkPlotStartAuto");
        Assert.Equal("GetPlotStartAuto", startCheck.Attribute("getPressed")?.Value);
        Assert.Equal("OnPlotStartAutoClick", startCheck.Attribute("onAction")?.Value);

        XElement finishCheck = group.Descendants(ns + "checkBox")
            .Single(element => element.Attribute("id")?.Value == "chkPlotFinishAuto");
        Assert.Equal("GetPlotFinishAuto", finishCheck.Attribute("getPressed")?.Value);
        Assert.Equal("OnPlotFinishAutoClick", finishCheck.Attribute("onAction")?.Value);

        XElement startBox = group.Descendants(ns + "editBox")
            .Single(element => element.Attribute("id")?.Value == "edtPlotStartDate");
        Assert.Equal("GetPlotStartDate", startBox.Attribute("getText")?.Value);
        Assert.Equal("GetPlotStartDateEnabled", startBox.Attribute("getEnabled")?.Value);
        Assert.Equal("OnPlotStartDateChange", startBox.Attribute("onChange")?.Value);

        // The box is sized for the default display format dd-mmm-yy.
        Assert.Equal("dd-mmm-yy", startBox.Attribute("sizeString")?.Value);

        XElement finishBox = group.Descendants(ns + "editBox")
            .Single(element => element.Attribute("id")?.Value == "edtPlotFinishDate");
        Assert.Equal("GetPlotFinishDate", finishBox.Attribute("getText")?.Value);
        Assert.Equal("GetPlotFinishDateEnabled", finishBox.Attribute("getEnabled")?.Value);
        Assert.Equal("OnPlotFinishDateChange", finishBox.Attribute("onChange")?.Value);
        Assert.Equal("dd-mmm-yy", finishBox.Attribute("sizeString")?.Value);
    }

    /// <summary>
    /// The edit-box commit path forwards the edited text: an exact dd/MM/yyyy
    /// commit persists the normalised date; the routing itself needs no workbook.
    /// </summary>
    [Fact]
    public void OnPlotStartDateChange_forwards_the_edited_text_to_the_service()
    {
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);

        GanttRibbon.OnPlotStartDateChange(null, "15-Mar-26", service);

        Assert.Equal("15-Mar-26", service.GetPlotStartDate());
        Assert.False(service.IsPlotStartAuto());
    }

    /// <summary>
    /// The setter invalidates exactly once per commit, so the clicked control
    /// repaints from the new snapshot: without an invalidate the ribbon
    /// getters never re-query and the box visually sticks.
    /// </summary>
    [Fact]
    public void SetPlotStartAuto_invalidates_once_per_commit()
    {
        var ribbon = new Mock<IRibbonUI>();
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);
        service.SetRibbon(ribbon.Object);

        service.SetPlotStartAuto(true);

        Assert.Equal(
            1,
            ribbon.Invocations.Count(invocation => invocation.Method.Name == nameof(IRibbonUI.Invalidate)));
    }

    /// <summary>
    /// The finish-end analogue: one commit, one invalidate.
    /// </summary>
    [Fact]
    public void SetPlotFinishAuto_invalidates_once_per_commit()
    {
        var ribbon = new Mock<IRibbonUI>();
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);
        service.SetRibbon(ribbon.Object);

        service.SetPlotFinishAuto(true);

        Assert.Equal(
            1,
            ribbon.Invocations.Count(invocation => invocation.Method.Name == nameof(IRibbonUI.Invalidate)));
    }

    /// <summary>
    /// Positive for the true-bounds display: with both ends AUTO and event
    /// data present, the boxes show the month-snapped chart bounds Refresh
    /// renders (01-Mar-26–30-Apr-26 for 10 Mar–31 Mar spans with the
    /// default 3-day pad), not the raw padded data extent (07/03–03/04).
    /// </summary>
    [Fact]
    public void Plot_date_getters_show_the_true_chart_bounds_in_auto_mode()
    {
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);
        var catalogue = new Mock<IConfigCatalogueReader>();
        _ = catalogue.Setup(r => r.Read()).Returns(ConfigReadOutcome.Ok(
            "wb",
            new Dictionary<string, string>
            {
                ["PlotStartMode"] = "DataRange",
                ["PlotFinishMode"] = "DataRange",
                ["PlotStartDate"] = string.Empty,
                ["PlotFinishDate"] = string.Empty,
            },
            GanttStyleRegistry.Empty));
        service.SetCatalogueReader(catalogue.Object);
        var table = new Mock<IGanttTableReader>();
        _ = table.Setup(r => r.Read()).Returns(GanttTableReadOutcome.Ok(
        [
            PlotBoundsRow(1, "10/03/2026", "31/03/2026"),
        ]));
        service.SetTableReader(table.Object);

        service.Refresh();

        Assert.Equal("01-Mar-26", service.GetPlotStartDate());
        Assert.Equal("30-Apr-26", service.GetPlotFinishDate());
    }

    private static int s_nextPlotBoundsId;

    private static string PlotBoundsId() => "G-" + (++s_nextPlotBoundsId).ToString("x32", CultureInfo.InvariantCulture);

    private static GanttRowDto PlotBoundsRow(int row, string start, string finish) =>
        new(
            row,
            PlotBoundsId(),
            PlotBoundsId(),
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

    [Fact]
    public void OnPlotFinishDateChange_forwards_the_edited_text_to_the_service()
    {
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);

        GanttRibbon.OnPlotFinishDateChange(null, "30-Jun-26", service);

        Assert.Equal("30-Jun-26", service.GetPlotFinishDate());
        Assert.False(service.IsPlotFinishAuto());
    }

    [Fact]
    public void Plot_date_getters_answer_from_the_snapshot()
    {
        var adapter = new Mock<IExcelApplicationAdapter>();
        _ = adapter.Setup(a => a.HasActiveWorkbook()).Returns(true);
        var tableReader = new Mock<IGanttTableReader>();
        _ = tableReader.Setup(t => t.Read()).Returns(GanttTableReadOutcome.Ok([]));
        var service = new RibbonStateService();
        service.SetApplicationAdapter(adapter.Object);
        service.SetTableReader(tableReader.Object);
        service.SetCatalogueWriter(null);
        service.SetPlotStartDate("15-Mar-26");
        service.SetPlotFinishDate("30-Jun-26");
        service.Refresh();

        var control = new Mock<IRibbonControl>();

        Assert.Equal("15-Mar-26", GanttRibbon.GetPlotStartDate(control.Object, service));
        Assert.Equal("30-Jun-26", GanttRibbon.GetPlotFinishDate(control.Object, service));

        // Explicit on both ends: the edit boxes must be enabled, and the
        // AUTO checkboxes unchecked.
        Assert.True(GanttRibbon.GetPlotStartDateEnabled(control.Object, service));
        Assert.True(GanttRibbon.GetPlotFinishDateEnabled(control.Object, service));
        Assert.False(GanttRibbon.GetPlotStartAuto(control.Object, service));
        Assert.False(GanttRibbon.GetPlotFinishAuto(control.Object, service));
    }

    [Fact]
    public void Plot_date_getters_disable_the_edit_box_while_auto_is_on()
    {
        var service = new RibbonStateService();

        var control = new Mock<IRibbonControl>();

        Assert.False(GanttRibbon.GetPlotStartDateEnabled(control.Object, service));
        Assert.False(GanttRibbon.GetPlotFinishDateEnabled(control.Object, service));
        Assert.True(GanttRibbon.GetPlotStartAuto(control.Object, service));
        Assert.True(GanttRibbon.GetPlotFinishAuto(control.Object, service));
    }

    [Fact]
    public void Plot_date_getText_degrades_to_empty_when_the_snapshot_probe_fails()
    {
        // Fail-open is the contract: a getter that throws into Excel's
        // dispatch breaks the whole ribbon, so every plot getter degrades to
        // the safe direction — empty text, AUTO checked (the initial state),
        // and the edit box enabled rather than permanently grey.
        var broken = new ThrowingStateService();
        var control = new Mock<IRibbonControl>();

        Assert.Equal(string.Empty, GanttRibbon.GetPlotStartDate(control.Object, broken));
        Assert.Equal(string.Empty, GanttRibbon.GetPlotFinishDate(control.Object, broken));
        Assert.True(GanttRibbon.GetPlotStartAuto(control.Object, broken));
        Assert.True(GanttRibbon.GetPlotFinishAuto(control.Object, broken));
        Assert.True(GanttRibbon.GetPlotStartDateEnabled(control.Object, broken));
        Assert.True(GanttRibbon.GetPlotFinishDateEnabled(control.Object, broken));
    }

    private sealed class ThrowingStateService : RibbonStateService
    {
        internal override string GetPlotStartDate() => throw new InvalidOperationException("probe failed");

        internal override string GetPlotFinishDate() => throw new InvalidOperationException("probe failed");

        internal override bool IsPlotStartAuto() => throw new InvalidOperationException("probe failed");

        internal override bool IsPlotFinishAuto() => throw new InvalidOperationException("probe failed");

        internal override bool GetEnabled(string? controlId) => true;
    }

    /// <summary>Null dependencies are refused at the boundary, not at first use.</summary>
    [Fact]
    public void OnPlotStartAutoClick_throws_for_a_null_state_service()
    {
        var control = new Mock<IRibbonControl>();

        Assert.Throws<ArgumentNullException>(
            () => GanttRibbon.OnPlotStartAutoClick(control.Object, true, null!));
    }

    /// <summary>Null dependencies are refused at the boundary, not at first use.</summary>
    [Fact]
    public void OnPlotFinishAutoClick_throws_for_a_null_state_service()
    {
        var control = new Mock<IRibbonControl>();

        Assert.Throws<ArgumentNullException>(
            () => GanttRibbon.OnPlotFinishAutoClick(control.Object, true, null!));
    }

    /// <summary>
    /// The checkbox honours Excel's pressed state (fix plan ruling 2): checking
    /// AUTO enters automatic mode, unchecking arms explicit mode.
    /// </summary>
    [Fact]
    public void OnPlotStartAutoClick_honours_the_pressed_state_Excel_reports()
    {
        var control = new Mock<IRibbonControl>();

        var autoService = new RibbonStateService();
        autoService.SetCatalogueWriter(null);
        GanttRibbon.OnPlotStartAutoClick(control.Object, true, autoService);
        Assert.True(autoService.IsPlotStartAuto());

        var explicitService = new RibbonStateService();
        explicitService.SetCatalogueWriter(null);
        explicitService.SetPlotStartDate("15-Mar-26");
        GanttRibbon.OnPlotStartAutoClick(control.Object, false, explicitService);
        Assert.False(explicitService.IsPlotStartAuto());
    }

    /// <summary>
    /// The checkbox honours Excel's pressed state (fix plan ruling 2): checking
    /// AUTO enters automatic mode, unchecking arms explicit mode.
    /// </summary>
    [Fact]
    public void OnPlotFinishAutoClick_honours_the_pressed_state_Excel_reports()
    {
        var control = new Mock<IRibbonControl>();

        var autoService = new RibbonStateService();
        autoService.SetCatalogueWriter(null);
        GanttRibbon.OnPlotFinishAutoClick(control.Object, true, autoService);
        Assert.True(autoService.IsPlotFinishAuto());

        var explicitService = new RibbonStateService();
        explicitService.SetCatalogueWriter(null);
        explicitService.SetPlotFinishDate("30-Jun-26");
        GanttRibbon.OnPlotFinishAutoClick(control.Object, false, explicitService);
        Assert.False(explicitService.IsPlotFinishAuto());
    }

    /// <summary>
    /// Positive for the uncheck fix: unchecking AUTO with a stored date
    /// persists Explicit with that date, so the next workbook load cannot
    /// revert the box to checked. The date commit's own write is not counted:
    /// only the uncheck's persist is asserted.
    /// </summary>
    [Fact]
    public void SetPlotStartAuto_uncheck_with_a_stored_date_persists_explicit()
    {
        var log = new Mock<IRollingLog>();
        CommandBoundary.Instance.SetLog(log.Object);
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);
        service.SetPlotStartDate("15-Mar-26");
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        _ = writer
            .Setup(w => w.WriteSettings(It.Is<IReadOnlyDictionary<string, string>>(
                settings =>
                    settings["PlotStartMode"] == "Explicit"
                    && settings["PlotStartDate"] == "15-Mar-26")))
            .Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);

        service.SetPlotStartAuto(false);

        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Once);
        Assert.False(service.IsPlotStartAuto());
        Assert.Equal("15-Mar-26", service.GetPlotStartDate());
    }

    /// <summary>
    /// Positive for the uncheck fix on the finish end: unchecking AUTO with a
    /// stored date persists Explicit with that date. The date commit's own
    /// write is not counted: only the uncheck's persist is asserted.
    /// </summary>
    [Fact]
    public void SetPlotFinishAuto_uncheck_with_a_stored_date_persists_explicit()
    {
        var log = new Mock<IRollingLog>();
        CommandBoundary.Instance.SetLog(log.Object);
        var service = new RibbonStateService();
        service.SetCatalogueWriter(null);
        service.SetPlotFinishDate("30-Jun-26");
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        _ = writer
            .Setup(w => w.WriteSettings(It.Is<IReadOnlyDictionary<string, string>>(
                settings =>
                    settings["PlotFinishMode"] == "Explicit"
                    && settings["PlotFinishDate"] == "30-Jun-26")))
            .Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);

        service.SetPlotFinishAuto(false);

        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Once);
        Assert.False(service.IsPlotFinishAuto());
        Assert.Equal("30-Jun-26", service.GetPlotFinishDate());
    }

    /// <summary>Null dependencies are refused at the boundary, not at first use.</summary>
    [Fact]
    public void OnPlotStartDateChange_throws_for_a_null_state_service()
    {
        var control = new Mock<IRibbonControl>();

        Assert.Throws<ArgumentNullException>(
            () => GanttRibbon.OnPlotStartDateChange(control.Object, "15-Mar-26", null!));
    }

    /// <summary>Null dependencies are refused at the boundary, not at first use.</summary>
    [Fact]
    public void OnPlotFinishDateChange_throws_for_a_null_state_service()
    {
        var control = new Mock<IRibbonControl>();

        Assert.Throws<ArgumentNullException>(
            () => GanttRibbon.OnPlotFinishDateChange(control.Object, "30-Jun-26", null!));
    }

    /// <summary>
    /// Null dependencies are refused at the boundary, not at first use.
    /// </summary>
    [Fact]
    public void GetPlotStartDate_throws_for_a_null_state_service()
    {
        var control = new Mock<IRibbonControl>();

        Assert.Throws<ArgumentNullException>(
            () => GanttRibbon.GetPlotStartDate(control.Object, null!));
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

    /// <summary>
    /// Parses the RibbonX <paramref name="xml"/> and returns the callback method
    /// names (the values of known callback attributes) that do not resolve to a
    /// public instance method on <paramref name="ribbonType"/> with a compatible
    /// signature: zero parameters, one parameter of type
    /// <see cref="IRibbonControl"/> or <see cref="IRibbonUI"/>, the two
    /// parameters <see cref="IRibbonControl"/> plus <see cref="string"/> that
    /// the editBox <c>onChange</c> callback passes, or the two parameters
    /// <see cref="IRibbonControl"/> plus <see cref="bool"/> that the checkBox
    /// <c>onAction</c> callback passes (fix plan ruling 2).
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

        if (parameters.Length == 2)
        {
            // The RibbonX editBox onChange callback passes the control plus the
            // edited text: public void OnChange(IRibbonControl, string).
            if (parameters[0].ParameterType == typeof(IRibbonControl)
                && parameters[1].ParameterType == typeof(string))
            {
                return true;
            }

            // The RibbonX checkBox onAction callback passes the control plus
            // the pressed state: public void OnClick(IRibbonControl, bool).
            return parameters[0].ParameterType == typeof(IRibbonControl)
                && parameters[1].ParameterType == typeof(bool);
        }

        return false;
    }
}
