using System.Globalization;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;

namespace GanttCreator.AddIn;

/// <summary>
/// Excel-DNA RibbonX entry point. Returns a minimal, valid RibbonX document
/// that declares a single Gantt Creator tab.
/// </summary>
/// <remarks>
/// Excel-DNA auto-discovers and registers any non-abstract
/// <see cref="ExcelRibbon"/> descendant (see
/// <c>AssemblyLoader.IsRibbonType</c>), so no .dna file change is required.
/// <see cref="GetCustomUI"/> does not call the base implementation: the base
/// reads <c>DnaLibrary.CustomUIs</c>, which is empty for this add-in and null
/// when not hosted in Excel.
/// </remarks>
[ComVisible(true)]
public class GanttRibbon : ExcelRibbon
{
    /// <summary>
    /// Excel's onLoad callback. Publishes the RibbonUI handle to
    /// <see cref="RibbonStateService"/> — the invalidate mechanism's anchor —
    /// then refreshes, so the dynamic getters are queried with the first
    /// snapshot (which includes the initial invalidation). Never throws: a
    /// throwing <c>onLoad</c> breaks the Ribbon.
    /// </summary>
    /// <param name="ribbon">The RibbonUI handle supplied by Excel.</param>
    public void OnLoad(IRibbonUI ribbon) => OnLoad(ribbon, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="OnLoad(IRibbonUI)"/> against an injected state service.
    /// Internal so contract tests can verify the handoff without the session
    /// singleton.
    /// </summary>
    /// <param name="ribbon">The RibbonUI handle, or null when Excel supplied none.</param>
    /// <param name="stateService">The state service receiving the handle.</param>
    internal static void OnLoad(IRibbonUI? ribbon, RibbonStateService stateService)
    {
        ArgumentNullException.ThrowIfNull(stateService);
        stateService.SetRibbon(ribbon);
        stateService.Refresh();
    }
    /// <summary>
    /// Returns the RibbonX document for the Excel workbook RibbonID, else null.
    /// The document is stored in the RibbonResources.resx resource (a
    /// ResXFileRef to <c>RibbonResources\Ribbon.xml</c>) and read through the
    /// committed <c>RibbonResources.Designer.cs</c> accessor, so the XML never
    /// appears as a C# string literal.
    /// </summary>
    /// <param name="RibbonID">The Ribbon identifier supplied by Excel.</param>
    /// <returns>
    /// The RibbonX document, or null when the RibbonID is not the workbook.
    /// </returns>
    public override string GetCustomUI(string RibbonID)
        => !string.IsNullOrWhiteSpace(RibbonID)
           && RibbonID.Trim().Equals("Microsoft.Excel.Workbook", StringComparison.OrdinalIgnoreCase)
               ? RibbonResources.Ribbon
               : null!;

    /// <summary>
    /// Called when the user clicks the Diagnostics button. The callback is a
    /// thin error boundary: it resolves the command name and delegates to the
    /// project-wide <see cref="CommandBoundary"/>, which invokes the
    /// <see cref="DiagnosticsService"/> command and produces one log record
    /// and one user-safe dialog on failure (docs/02-ARCHITECTURE.md
    /// "Ribbon and commands": no callback contains command logic). After the
    /// boundary run the ribbon state is refreshed and invalidated (work item
    /// R1.5 decision D3: every command run through the ribbon's boundary).
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnDiagnosticsClick(IRibbonControl control)
        => OnDiagnosticsClick(control, CommandBoundary.Instance, RunDiagnostics, NotifyRibbonStateChanged);

    /// <summary>
    /// Runs the Diagnostics command across an injected boundary. Internal so
    /// contract tests can verify the routing with a deterministic boundary
    /// and a stub command instead of the singleton and the real dialog.
    /// <paramref name="onCompleted"/> runs after the boundary returns (the
    /// boundary never throws), and is the test seam for the post-command
    /// ribbon-state refresh.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event, or null when unavailable.</param>
    /// <param name="boundary">The command boundary to run the command across.</param>
    /// <param name="command">The diagnostics command delegate.</param>
    /// <param name="onCompleted">Invoked once after the boundary run, or null to skip it.</param>
    internal static void OnDiagnosticsClick(
        IRibbonControl? control,
        CommandBoundary boundary,
        Action command,
        Action? onCompleted = null)
    {
        ArgumentNullException.ThrowIfNull(boundary);
        ArgumentNullException.ThrowIfNull(command);
        boundary.Run(() => ResolveCommandName(control), command, nameof(OnDiagnosticsClick));
        onCompleted?.Invoke();
    }

    /// <summary>
    /// The Diagnostics command: shows the diagnostics dialog via the
    /// project-wide singleton.
    /// </summary>
    private static void RunDiagnostics() => DiagnosticsService.Instance.ShowDiagnostics();

    /// <summary>
    /// Called when the user clicks the Open log file button. The callback is a
    /// thin error boundary exactly like the Diagnostics callback: the open
    /// command runs across the project-wide <see cref="CommandBoundary"/> and
    /// the ribbon state is refreshed afterwards (work item R1.5 decision D3).
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnOpenLogClick(IRibbonControl control)
        => OnOpenLogClick(control, CommandBoundary.Instance, RunOpenLog, NotifyRibbonStateChanged);

    /// <summary>
    /// Runs the Open-log command across an injected boundary. Internal so
    /// contract tests can verify the routing without the singletons and the
    /// real file open.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event, or null when unavailable.</param>
    /// <param name="boundary">The command boundary to run the command across.</param>
    /// <param name="command">The open-log command delegate.</param>
    /// <param name="onCompleted">Invoked once after the boundary run, or null to skip it.</param>
    internal static void OnOpenLogClick(
        IRibbonControl? control,
        CommandBoundary boundary,
        Action command,
        Action? onCompleted = null)
    {
        ArgumentNullException.ThrowIfNull(boundary);
        ArgumentNullException.ThrowIfNull(command);
        boundary.Run(() => ResolveCommandName(control, nameof(OnOpenLogClick)), command, nameof(OnOpenLogClick));
        onCompleted?.Invoke();
    }

    /// <summary>
    /// The Open-log command: opens the active log file with the system default
    /// handler. A blank log path is a no-op; the control is disabled when the
    /// log is unavailable and the open itself is non-fatal.
    /// </summary>
    private static void RunOpenLog()
    {
        var path = DiagnosticsService.Instance.LogFilePath;
        if (!string.IsNullOrWhiteSpace(path))
        {
            DiagnosticsService.OpenLogFile(path);
        }
    }

    /// <summary>
    /// Called when the user clicks the Initialise sheet button. The callback is
    /// a thin error boundary exactly like the Diagnostics callback: the
    /// initialise command runs across the project-wide
    /// <see cref="CommandBoundary"/> and the ribbon state is refreshed
    /// afterwards (work item R1.5 decision D3). The initialise command itself
    /// surfaces typed refusals; unexpected failures are translated here.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnInitialiseSheetClick(IRibbonControl control)
        => OnInitialiseSheetClick(control, CommandBoundary.Instance, RunInitialiseSheet, NotifyRibbonStateChanged);

    /// <summary>
    /// Runs the Initialise-sheet command across an injected boundary. Internal
    /// so contract tests can verify the routing without the singletons and the
    /// real workbook mutation.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event, or null when unavailable.</param>
    /// <param name="boundary">The command boundary to run the command across.</param>
    /// <param name="command">The initialise command delegate.</param>
    /// <param name="onCompleted">Invoked once after the boundary run, or null to skip it.</param>
    internal static void OnInitialiseSheetClick(
        IRibbonControl? control,
        CommandBoundary boundary,
        Action command,
        Action? onCompleted = null)
    {
        ArgumentNullException.ThrowIfNull(boundary);
        ArgumentNullException.ThrowIfNull(command);
        boundary.Run(() => ResolveCommandName(control, nameof(OnInitialiseSheetClick)), command, nameof(OnInitialiseSheetClick));
        onCompleted?.Invoke();
    }

    /// <summary>
    /// The Initialise-sheet command: runs the workbook initialiser for the
    /// current Excel session via the application command.
    /// </summary>
    private static void RunInitialiseSheet() => InitialiseSheetCommand.RunForExcel();

    /// <summary>
    /// Called when the user clicks the Validate button. The callback is a
    /// thin error boundary exactly like the Initialise-sheet callback: the
    /// validate command runs across the project-wide
    /// <see cref="CommandBoundary"/> and the ribbon state is refreshed
    /// afterwards (work item R1.5 decision D3). The validate command itself
    /// surfaces typed refusals; unexpected failures are translated here.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnValidateSheetClick(IRibbonControl control)
        => OnValidateSheetClick(control, CommandBoundary.Instance, RunValidateSheet, NotifyRibbonStateChanged);

    /// <summary>
    /// Runs the Validate-sheet command across an injected boundary. Internal so
    /// contract tests can verify the routing without the singletons and the real
    /// workbook mutation.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event, or null when unavailable.</param>
    /// <param name="boundary">The command boundary to run the command across.</param>
    /// <param name="command">The validate command delegate.</param>
    /// <param name="onCompleted">Invoked once after the boundary run, or null to skip it.</param>
    internal static void OnValidateSheetClick(
        IRibbonControl? control,
        CommandBoundary boundary,
        Action command,
        Action? onCompleted = null)
    {
        ArgumentNullException.ThrowIfNull(boundary);
        ArgumentNullException.ThrowIfNull(command);
        boundary.Run(() => ResolveCommandName(control, nameof(OnValidateSheetClick)), command, nameof(OnValidateSheetClick));
        onCompleted?.Invoke();
    }

    /// <summary>
    /// The Validate-sheet command: runs the reader → validator → reporter
    /// pipeline for the current Excel session via the application command.
    /// </summary>
    private static void RunValidateSheet() => ValidateSheetCommand.RunForExcel();

    /// <summary>
    /// Called when the user clicks the Refresh chart button. The callback is a thin
    /// error boundary exactly like the Validate one: the refresh command runs across
    /// the project-wide <see cref="CommandBoundary"/> and the ribbon state is
    /// refreshed afterwards (R4.9 D3).
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnRefreshSheetClick(IRibbonControl control)
        => OnRefreshSheetClick(control, CommandBoundary.Instance, RunRefreshSheet, NotifyRibbonStateChanged);

    /// <summary>
    /// Runs the Refresh-sheet command across an injected boundary. Internal so
    /// contract tests can verify the routing without the singletons and the real
    /// workbook mutation.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event, or null when unavailable.</param>
    /// <param name="boundary">The command boundary to run the command across.</param>
    /// <param name="command">The refresh command delegate.</param>
    /// <param name="onCompleted">Invoked once after the boundary run, or null to skip it.</param>
    /// <remarks>
    /// Identical in shape to the Validate routing, deliberately. A refresh that took a
    /// different path — its own try/catch, its own state refresh, its own message box
    /// — would be a second command boundary, which is the one thing the architecture
    /// forbids. The command itself holds the pipeline; this holds only the error
    /// translation.
    /// </remarks>
    internal static void OnRefreshSheetClick(
        IRibbonControl? control,
        CommandBoundary boundary,
        Action command,
        Action? onCompleted = null)
    {
        ArgumentNullException.ThrowIfNull(boundary);
        ArgumentNullException.ThrowIfNull(command);
        boundary.Run(() => ResolveCommandName(control, nameof(OnRefreshSheetClick)), command, nameof(OnRefreshSheetClick));
        onCompleted?.Invoke();
    }

    /// <summary>
    /// The Refresh-sheet command: assembles the live orchestrator and runs one whole
    /// refresh for the current Excel session.
    /// </summary>
    private static void RunRefreshSheet() => RefreshSheetCommand.RunForExcel();

    /// <summary>Called when the user clicks the Repair configuration button.</summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnRepairConfigClick(IRibbonControl control)
        => OnRepairConfigClick(control, CommandBoundary.Instance, RunRepairConfig, NotifyRibbonStateChanged);

    /// <summary>Runs the Repair configuration command across an injected boundary.</summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="boundary">The command boundary.</param>
    /// <param name="command">The repair command.</param>
    /// <param name="onCompleted">The post-command state hook.</param>
    internal static void OnRepairConfigClick(
        IRibbonControl? control,
        CommandBoundary boundary,
        Action command,
        Action? onCompleted = null)
    {
        ArgumentNullException.ThrowIfNull(boundary);
        ArgumentNullException.ThrowIfNull(command);
        boundary.Run(
            () => ResolveCommandName(control, nameof(OnRepairConfigClick)),
            command,
            nameof(OnRepairConfigClick));
        onCompleted?.Invoke();
    }

    /// <summary>Runs the production repair command.</summary>
    private static void RunRepairConfig() => RepairConfigCommand.RunForExcel();

    /// <summary>Called when the user clicks the Add activity button.</summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnAddActivityClick(IRibbonControl control)
        => OnAddActivityClick(control, CommandBoundary.Instance, () => AddRowCommand.RunForExcel(GanttEntityType.AsPlannedActivity), NotifyRibbonStateChanged);

    /// <summary>Called when the user clicks the Add milestone button.</summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnAddMilestoneClick(IRibbonControl control)
        => OnAddMilestoneClick(control, CommandBoundary.Instance, () => AddRowCommand.RunForExcel(GanttEntityType.AsPlannedMilestone), NotifyRibbonStateChanged);

    /// <summary>Called when the user clicks the Add delineator button.</summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnAddDelineatorClick(IRibbonControl control)
        => OnAddDelineatorClick(control, CommandBoundary.Instance, () => AddRowCommand.RunForExcel(GanttEntityType.Delineator), NotifyRibbonStateChanged);

    /// <summary>Runs the Add activity callback across an injected boundary.</summary>
    /// <param name="control">The ribbon control.</param>
    /// <param name="boundary">The command boundary.</param>
    /// <param name="command">The add-row command.</param>
    /// <param name="onCompleted">The post-command state hook.</param>
    internal static void OnAddActivityClick(IRibbonControl? control, CommandBoundary boundary, Action command, Action? onCompleted = null)
        => OnAddRowClick(control, boundary, command, nameof(OnAddActivityClick), onCompleted);

    /// <summary>Runs the Add milestone callback across an injected boundary.</summary>
    /// <param name="control">The ribbon control.</param>
    /// <param name="boundary">The command boundary.</param>
    /// <param name="command">The add-row command.</param>
    /// <param name="onCompleted">The post-command state hook.</param>
    internal static void OnAddMilestoneClick(IRibbonControl? control, CommandBoundary boundary, Action command, Action? onCompleted = null)
        => OnAddRowClick(control, boundary, command, nameof(OnAddMilestoneClick), onCompleted);

    /// <summary>Runs the Add delineator callback across an injected boundary.</summary>
    /// <param name="control">The ribbon control.</param>
    /// <param name="boundary">The command boundary.</param>
    /// <param name="command">The add-row command.</param>
    /// <param name="onCompleted">The post-command state hook.</param>
    internal static void OnAddDelineatorClick(IRibbonControl? control, CommandBoundary boundary, Action command, Action? onCompleted = null)
        => OnAddRowClick(control, boundary, command, nameof(OnAddDelineatorClick), onCompleted);

    private static void OnAddRowClick(
        IRibbonControl? control,
        CommandBoundary boundary,
        Action command,
        string fallbackCommandName,
        Action? onCompleted)
    {
        ArgumentNullException.ThrowIfNull(boundary);
        ArgumentNullException.ThrowIfNull(command);
        boundary.Run(() => ResolveCommandName(control, fallbackCommandName), command, fallbackCommandName);
        onCompleted?.Invoke();
    }

    /// <summary>
    /// Excel's getEnabled callback for the gated controls. A pure read of the
    /// state service's cached snapshot: fast, side-effect-free, and fail-open —
    /// an absent control or a control whose ID cannot be probed stays enabled
    /// (a permanently grey button is a worse failure than a briefly enabled
    /// one, and the command boundary still validates at execution time).
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public bool GetEnabled(IRibbonControl control) => GetEnabled(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetEnabled(IRibbonControl)"/> against an injected state
    /// service. Internal so contract tests can verify the routing without the
    /// session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The control's enabled state.</returns>
    internal static bool GetEnabled(IRibbonControl? control, RibbonStateService stateService)
    {
        ArgumentNullException.ThrowIfNull(stateService);
        // CA1031: the only failure source here is the COM control-ID probe; a
        // probe failure must fail the getter open, never propagate into Excel's
        // getter dispatch.
#pragma warning disable CA1031
        try
        {
            return stateService.GetEnabled(control?.Id);
        }
        catch
        {
            return true;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Excel's getText callback for the start-date edit box. Returns the
    /// effective display date: the explicit date in explicit mode, the
    /// current month-snapped chart bound in automatic mode, or empty when
    /// neither is known. This is also the value the edit box reverts to when
    /// the user enters an unparseable date. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public string GetPlotStartDate(IRibbonControl control) => GetPlotStartDate(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetPlotStartDate(IRibbonControl)"/> against an injected
    /// state service. Internal so contract tests can verify the routing
    /// without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The display start date, or empty when none.</returns>
    internal static string GetPlotStartDate(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
        // CA1031: a probe failure must degrade to empty text rather than
        // propagating into Excel's getter dispatch.
#pragma warning disable CA1031
        try
        {
            return stateService.GetPlotStartDate();
        }
        catch
        {
            return string.Empty;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Excel's getText callback for the finish-date edit box. Returns the
    /// effective display date: the explicit date in explicit mode, the
    /// current month-snapped chart bound in automatic mode, or empty when
    /// neither is known. This is also the value the edit box reverts to when
    /// the user enters an unparseable date. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public string GetPlotFinishDate(IRibbonControl control) => GetPlotFinishDate(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetPlotFinishDate(IRibbonControl)"/> against an injected
    /// state service. Internal so contract tests can verify the routing
    /// without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The display finish date, or empty when none.</returns>
    internal static string GetPlotFinishDate(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.GetPlotFinishDate();
        }
        catch
        {
            return string.Empty;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Excel's getEnabled callback for the start-date edit box. The edit box is
    /// enabled only while the AUTO checkbox is unchecked: when automatic mode
    /// is on the start derives from the data and the user cannot type a date.
    /// Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public bool GetPlotStartDateEnabled(IRibbonControl control) => GetPlotStartDateEnabled(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetPlotStartDateEnabled(IRibbonControl)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>True when the edit box should be enabled.</returns>
    internal static bool GetPlotStartDateEnabled(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.GetEnabled(RibbonControlIds.PlotStartDate);
        }
        catch
        {
            return true;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Excel's getEnabled callback for the finish-date edit box. The edit box is
    /// enabled only while the AUTO checkbox is unchecked: when automatic mode
    /// is on the finish derives from the data and the user cannot type a date.
    /// Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public bool GetPlotFinishDateEnabled(IRibbonControl control) => GetPlotFinishDateEnabled(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetPlotFinishDateEnabled(IRibbonControl)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>True when the edit box should be enabled.</returns>
    internal static bool GetPlotFinishDateEnabled(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.GetEnabled(RibbonControlIds.PlotFinishDate);
        }
        catch
        {
            return true;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Excel's getChecked callback for the start AUTO checkbox: checked when
    /// the start derives from the data. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public bool GetPlotStartAuto(IRibbonControl control) => GetPlotStartAuto(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetPlotStartAuto(IRibbonControl)"/> against an injected
    /// state service. Internal so contract tests can verify the routing
    /// without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>True when the start end is automatic.</returns>
    internal static bool GetPlotStartAuto(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.IsPlotStartAuto();
        }
        catch
        {
            // Fail open: a checked AUTO box matches the initial automatic
            // state, so a probe failure shows the safe default.
            return true;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Excel's getChecked callback for the finish AUTO checkbox: checked when
    /// the finish derives from the data. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public bool GetPlotFinishAuto(IRibbonControl control) => GetPlotFinishAuto(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetPlotFinishAuto(IRibbonControl)"/> against an injected
    /// state service. Internal so contract tests can verify the routing
    /// without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>True when the finish end is automatic.</returns>
    internal static bool GetPlotFinishAuto(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.IsPlotFinishAuto();
        }
        catch
        {
            return true;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Called when the user clicks the Plot-Area AUTO checkbox for the start end.
    /// Excel's checkbox onAction contract passes the pressed state, which this
    /// honours (fix plan ruling 2): the mode is set to what Excel reports,
    /// then persisted and the ribbon state refreshed. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    /// <param name="pressed">True when Excel reports the AUTO box checked.</param>
    public void OnPlotStartAutoClick(IRibbonControl control, bool pressed)
        => OnPlotStartAutoClick(control, pressed, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="OnPlotStartAutoClick(IRibbonControl, bool)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="pressed">True when Excel reports the AUTO box checked.</param>
    /// <param name="stateService">The state service receiving the setting.</param>
    internal static void OnPlotStartAutoClick(IRibbonControl? control, bool pressed, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
        stateService.SetPlotStartAuto(pressed);
    }

    /// <summary>
    /// Called when the user clicks the Plot-Area AUTO checkbox for the finish end.
    /// Excel's checkbox onAction contract passes the pressed state, which this
    /// honours (fix plan ruling 2): the mode is set to what Excel reports,
    /// then persisted and the ribbon state refreshed. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    /// <param name="pressed">True when Excel reports the AUTO box checked.</param>
    public void OnPlotFinishAutoClick(IRibbonControl control, bool pressed)
        => OnPlotFinishAutoClick(control, pressed, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="OnPlotFinishAutoClick(IRibbonControl, bool)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="pressed">True when Excel reports the AUTO box checked.</param>
    /// <param name="stateService">The state service receiving the setting.</param>
    internal static void OnPlotFinishAutoClick(IRibbonControl? control, bool pressed, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
        stateService.SetPlotFinishAuto(pressed);
    }

    /// <summary>
    /// Called when the user commits an edit to the start-date edit box.
    /// The new text is validated and persisted, then the ribbon state is
    /// refreshed. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    /// <param name="text">The edited text Excel passes on commit.</param>
    public void OnPlotStartDateChange(IRibbonControl control, string text)
        => OnPlotStartDateChange(control, text, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="OnPlotStartDateChange(IRibbonControl, string)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="text">The edited text Excel passes on commit, or null in tests.</param>
    /// <param name="stateService">The state service receiving the setting.</param>
    internal static void OnPlotStartDateChange(IRibbonControl? control, string? text, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
        stateService.SetPlotStartDate(text ?? string.Empty);
    }

    /// <summary>
    /// Called when the user commits an edit to the finish-date edit box.
    /// The new text is validated and persisted, then the ribbon state is
    /// refreshed. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    /// <param name="text">The edited text Excel passes on commit.</param>
    public void OnPlotFinishDateChange(IRibbonControl control, string text)
        => OnPlotFinishDateChange(control, text, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="OnPlotFinishDateChange(IRibbonControl, string)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="text">The edited text Excel passes on commit, or null in tests.</param>
    /// <param name="stateService">The state service receiving the setting.</param>
    internal static void OnPlotFinishDateChange(IRibbonControl? control, string? text, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
        stateService.SetPlotFinishDate(text ?? string.Empty);
    }

    /// <summary>
    /// Returns the index of the current plot time scale. The ribbon
    /// dropdown asks for this through getSelectedItemIndex; the answer
    /// comes from the state service snapshot. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public int GetTimeScale(IRibbonControl control) => GetTimeScale(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetTimeScale(IRibbonControl)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The zero-based index of the current time scale.</returns>
    internal static int GetTimeScale(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return (int)stateService.GetTimeScale();
        }
        catch
        {
            return (int)GanttTimeScale.Month;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Called when the user picks a plot time scale from the ribbon
    /// dropdown. Excel's dropDown onAction contract passes the control, the
    /// selected item id, and the selected index; the scale is set from the
    /// id, then persisted. An unknown id is refused by the service rather
    /// than defaulted, so a mistyped selection cannot silently redraw the
    /// period band with the wrong calendar unit. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    /// <param name="selectedId">The selected item id (Month, Quarter, or Week).</param>
    /// <param name="selectedIndex">The selected item index (unused; the id is authoritative).</param>
#pragma warning disable IDE0060 // The index is part of Excel's dropDown onAction contract; Excel always passes it.
    public void OnTimeScaleChange(IRibbonControl control, string selectedId, int selectedIndex)
        => OnTimeScaleChange(control, selectedId, RibbonStateService.Instance);
#pragma warning restore IDE0060

    /// <summary>
    /// Runs <see cref="OnTimeScaleChange(IRibbonControl, string, int)"/>
    /// against an injected state service. Internal so contract tests can
    /// verify the routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="id">The selected item id, or null in tests.</param>
    /// <param name="stateService">The state service receiving the setting.</param>
    internal static void OnTimeScaleChange(IRibbonControl? control, string? id, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
        if (GanttChartSettings.TryParseTimeScale(id, out GanttTimeScale scale))
        {
            stateService.SetTimeScale(scale);
        }
    }

    /// <summary>
    /// Returns the index of the current month period label format. The ribbon
    /// dropdown asks for this through getSelectedItemIndex; the answer comes
    /// from the state service snapshot. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public int GetPeriodLabelFormat(IRibbonControl control) => GetPeriodLabelFormat(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetPeriodLabelFormat(IRibbonControl)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The zero-based index of the current format (MM = 0, MMM = 1).</returns>
    internal static int GetPeriodLabelFormat(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return (int)stateService.GetPeriodLabelFormat();
        }
        catch
        {
            return (int)GanttPeriodLabelFormat.MMM;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Called when the user picks a period label format from the ribbon
    /// dropdown. Excel's dropDown onAction contract passes the control, the
    /// selected item id, and the selected index; the format is set from the
    /// id, then persisted. An unknown id, or one incompatible with the current
    /// scale, is refused by the service rather than defaulted, so a mistyped
    /// selection cannot silently relabel the band. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    /// <param name="selectedId">The selected item id (MM or MMM).</param>
    /// <param name="selectedIndex">The selected item index (unused; the id is authoritative).</param>
#pragma warning disable IDE0060 // The index is part of Excel's dropDown onAction contract; Excel always passes it.
    public void OnPeriodLabelFormatChange(IRibbonControl control, string selectedId, int selectedIndex)
        => OnPeriodLabelFormatChange(control, selectedId, RibbonStateService.Instance);
#pragma warning restore IDE0060

    /// <summary>
    /// Runs <see cref="OnPeriodLabelFormatChange(IRibbonControl, string, int)"/>
    /// against an injected state service. Internal so contract tests can verify
    /// the routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="id">The selected item id, or null in tests.</param>
    /// <param name="stateService">The state service receiving the setting.</param>
    internal static void OnPeriodLabelFormatChange(IRibbonControl? control, string? id, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
        if (GanttChartSettings.TryParsePeriodLabelFormat(id, out GanttPeriodLabelFormat format))
        {
            stateService.SetPeriodLabelFormat(format);
        }
    }

    /// <summary>
    /// Returns the index of the current plot margin preset. The ribbon
    /// dropdown asks for this through getSelectedItemIndex; the answer
    /// comes from the state service snapshot. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public int GetMargin(IRibbonControl control) => GetMargin(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetMargin(IRibbonControl)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The zero-based index of the current margin preset.</returns>
    internal static int GetMargin(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return (int)stateService.GetMargin();
        }
        catch
        {
            return (int)GanttPlotMargins.Default;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Called when the user picks a plot margin from the ribbon
    /// dropdown. Excel's dropDown onAction contract passes the control, the
    /// selected item id, and the selected index; the margin is set from the
    /// id, then persisted. An unknown id is refused by the service rather
    /// than defaulted. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    /// <param name="id">The selected item id (Narrow, Normal, Wide, or Custom).</param>
    /// <param name="selectedIndex">The selected item index (unused; the id is authoritative).</param>
#pragma warning disable IDE0060 // The index is part of Excel's dropDown onAction contract; Excel always passes it.
    public void OnMarginChange(IRibbonControl control, string id, int selectedIndex)
        => OnMarginChange(control, id, RibbonStateService.Instance);
#pragma warning restore IDE0060

    /// <summary>
    /// Runs <see cref="OnMarginChange(IRibbonControl, string, int)"/>
    /// against an injected state service. Internal so contract tests can
    /// verify the routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="id">The selected item id, or null in tests.</param>
    /// <param name="stateService">The state service receiving the setting.</param>
    internal static void OnMarginChange(IRibbonControl? control, string? id, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
        if (GanttPlotMargins.TryParse(id, out GanttPlotMargin margin))
        {
            stateService.SetMargin(margin);
        }
    }

    /// <summary>
    /// Returns the current custom margin in centimetres. The ribbon edit
    /// box asks for this through getText; the answer comes from the
    /// state service snapshot. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public string GetMarginCm(IRibbonControl control) => GetMarginCm(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetMarginCm(IRibbonControl)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The custom margin in centimetres, formatted with the invariant culture.</returns>
    internal static string GetMarginCm(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.GetMarginCm().ToString(CultureInfo.InvariantCulture);
        }
        catch
        {
            return GanttPlotMargins.DefaultCustomCm.ToString(CultureInfo.InvariantCulture);
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Returns whether the custom-margin edit box is enabled. It is
    /// enabled only while the margin is set to Custom, so the box
    /// cannot be edited while a fixed preset is selected. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public bool GetMarginCmEnabled(IRibbonControl control) => GetMarginCmEnabled(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetMarginCmEnabled(IRibbonControl)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>True when the margin is set to Custom.</returns>
    internal static bool GetMarginCmEnabled(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.GetMargin() == GanttPlotMargin.Custom;
        }
        catch
        {
            return false;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Called when the user commits an edit to the custom-margin edit
    /// box. The new text is parsed with the invariant culture and
    /// persisted, then the ribbon state is refreshed; an unparsable
    /// value is a no-op so a partial keystroke cannot corrupt the
    /// stored margin. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    /// <param name="text">The edited text Excel passes on commit.</param>
    public void OnMarginCmChange(IRibbonControl control, string text)
        => OnMarginCmChange(control, text, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="OnMarginCmChange(IRibbonControl, string)"/>
    /// against an injected state service. Internal so contract tests can
    /// verify the routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="text">The edited text Excel passes on commit, or null in tests.</param>
    /// <param name="stateService">The state service receiving the setting.</param>
    internal static void OnMarginCmChange(IRibbonControl? control, string? text, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var cm))
        {
            stateService.SetMarginCm(cm);
        }
    }

    /// <summary>
    /// The enabled getter for the read-only plot-dimension displays (work
    /// item R5.2 D4). Always false: the width/height boxes are outputs, never
    /// inputs, so Excel renders them permanently greyed. A dedicated getter —
    /// not GetEnabled — because the truth table fail-opens unknown IDs, and a
    /// read-only display must never be editable for any workbook state.
    /// Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    /// <returns>Always false.</returns>
    public bool GetPlotDimensionsEnabled(IRibbonControl control)
    {
        _ = control;
        return false;
    }

    /// <summary>
    /// Returns the current plot width in points. The ribbon edit box
    /// asks for this through getText; the answer is the measured panel
    /// grid width, so the displayed figure and the rendered plot cannot
    /// drift. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public string GetPlotWidth(IRibbonControl control) => GetPlotWidth(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetPlotWidth(IRibbonControl)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The plot width in points, formatted with the invariant culture.</returns>
    internal static string GetPlotWidth(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.GetPlotWidth();
        }
        catch
        {
            return string.Empty;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Returns the current plot height in points. The ribbon edit box
    /// asks for this through getText; the answer comes from the state
    /// service snapshot. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public string GetPlotHeight(IRibbonControl control) => GetPlotHeight(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetPlotHeight(IRibbonControl)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The plot height in points, formatted with the invariant culture.</returns>
    internal static string GetPlotHeight(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.GetPlotHeight();
        }
        catch
        {
            return string.Empty;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Called when the user clicks a plot-size preset button. All four
    /// buttons share this handler; Excel's button onAction contract
    /// passes the control, and the control id selects the preset. The
    /// preset is set, then persisted. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnPresetClick(IRibbonControl control)
        => OnPresetClick(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="OnPresetClick(IRibbonControl)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="stateService">The state service receiving the setting.</param>
    internal static void OnPresetClick(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
        // An unknown control id is refused — never defaulted to a preset —
        // so a mistyped or renamed button cannot silently render the chart on
        // the wrong paper.
        SizePreset? preset = control?.Id switch
        {
            RibbonControlIds.PresetA4Portrait => SizePresets.A4Portrait,
            RibbonControlIds.PresetA4Landscape => SizePresets.A4Landscape,
            RibbonControlIds.PresetPresentation16x9 => SizePresets.Presentation16x9,
            RibbonControlIds.PresetPresentation4x3 => SizePresets.Presentation4x3,
            _ => null,
        };
        if (preset is not null)
        {
            stateService.SetPreset(preset);
        }
    }

    /// <summary>
    /// Returns whether a plot-size preset button is pressed. Excel's
    /// button getPressed contract asks for this when a button carries
    /// getPressed; the answer compares the control id against the
    /// current preset so the active preset is visually marked. Never
    /// throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public bool GetPreset(IRibbonControl control) => GetPreset(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetPreset(IRibbonControl)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>True when the control's id names the current preset.</returns>
    internal static bool GetPreset(IRibbonControl? control, RibbonStateService stateService)
    {
        ArgumentNullException.ThrowIfNull(stateService);
        SizePreset? expected = control?.Id switch
        {
            RibbonControlIds.PresetA4Portrait => SizePresets.A4Portrait,
            RibbonControlIds.PresetA4Landscape => SizePresets.A4Landscape,
            RibbonControlIds.PresetPresentation16x9 => SizePresets.Presentation16x9,
            RibbonControlIds.PresetPresentation4x3 => SizePresets.Presentation4x3,
            _ => null,
        };
#pragma warning disable CA1031
        try
        {
            // Records do not overload ==, so compare the durable key
            // rather than relying on the reference identity of the
            // catalogue singletons.
            return expected is not null
                && string.Equals(stateService.GetPreset()?.Key, expected.Key, StringComparison.Ordinal);
        }
        catch
        {
            return false;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// The post-command ribbon-state hook (work item R1.5 decision D3): every
    /// command run through the ribbon's boundary ends with one refresh and
    /// invalidation. Never throws.
    /// </summary>
    private static void NotifyRibbonStateChanged() => RibbonStateService.Instance.Refresh();
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <returns>The command name used in log records and the force-failure hook.</returns>
    internal static string ResolveCommandName(IRibbonControl? control)
        => ResolveCommandName(control, nameof(OnDiagnosticsClick));

    /// <summary>
    /// Resolves the stable command name for a Ribbon callback: the control's
    /// ID when available, otherwise <paramref name="fallbackCommandName"/>.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="fallbackCommandName">The name used when the ID cannot be probed.</param>
    /// <returns>The command name used in log records and the force-failure hook.</returns>
    internal static string ResolveCommandName(IRibbonControl? control, string fallbackCommandName)
    {
        var id = control?.Id;
        return string.IsNullOrWhiteSpace(id) ? fallbackCommandName : id;
    }
}

