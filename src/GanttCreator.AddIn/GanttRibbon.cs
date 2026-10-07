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
    /// Excel's getText callback for the start-date edit box. Returns the last
    /// valid stored value, which is also the value the edit box reverts to when
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
    /// <returns>The stored start date, or empty when none.</returns>
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
    /// Excel's getText callback for the finish-date edit box. Returns the last
    /// valid stored value, which is also the value the edit box reverts to when
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
    /// <returns>The stored finish date, or empty when none.</returns>
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
            return !stateService.IsPlotStartAuto();
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
            return !stateService.IsPlotFinishAuto();
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
    /// Toggles the automatic mode, then persists and refreshes the ribbon state.
    /// Never throws.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnPlotStartAutoClick(IRibbonControl control)
        => OnPlotStartAutoClick(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="OnPlotStartAutoClick(IRibbonControl)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="stateService">The state service receiving the setting.</param>
    internal static void OnPlotStartAutoClick(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
        stateService.TogglePlotStartAuto();
        NotifyRibbonStateChanged();
    }

    /// <summary>
    /// Called when the user clicks the Plot-Area AUTO checkbox for the finish end.
    /// Toggles the automatic mode, then persists and refreshes the ribbon state.
    /// Never throws.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnPlotFinishAutoClick(IRibbonControl control)
        => OnPlotFinishAutoClick(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="OnPlotFinishAutoClick(IRibbonControl)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="stateService">The state service receiving the setting.</param>
    internal static void OnPlotFinishAutoClick(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
        stateService.TogglePlotFinishAuto();
        NotifyRibbonStateChanged();
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
        NotifyRibbonStateChanged();
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
        NotifyRibbonStateChanged();
    }

    /// <summary>
    /// The post-command ribbon-state hook (work item R1.5 decision D3): every
    /// command run through the ribbon's boundary ends with one refresh and
    /// invalidation. Never throws.
    /// </summary>
    private static void NotifyRibbonStateChanged() => RibbonStateService.Instance.Refresh();

    /// <summary>
    /// Resolves the stable command name for a Ribbon callback: the control's
    /// ID when available, otherwise the Diagnostics callback method name.
    /// </summary>
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
