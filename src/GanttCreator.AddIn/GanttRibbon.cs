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
        string? path = DiagnosticsService.Instance.LogFilePath;
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

    /// <summary>
    /// Called when the user clicks Promote to parent. The callback is a thin error
    /// boundary, exactly like the add-row clicks: it resolves the command name and
    /// delegates to <see cref="CommandBoundary"/>, which produces one log record
    /// and one user-safe dialog on an unexpected failure. A typed planner or
    /// mutation refusal is presented by the command itself (R5.3 D6); the ribbon
    /// state is refreshed and invalidated after every run (R1.5 D3).
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnPromoteToParentClick(IRibbonControl control) =>
        OnPromoteToParentClick(control, CommandBoundary.Instance, PromoteToParentCommand.RunForExcel, NotifyRibbonStateChanged);

    /// <summary>
    /// Runs <see cref="OnPromoteToParentClick(IRibbonControl)"/> against an
    /// injected boundary. Internal so contract tests can verify the routing
    /// without the session singleton.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="boundary">The command error boundary.</param>
    /// <param name="command">The promote command to run.</param>
    /// <param name="onCompleted">The post-command state hook.</param>
    internal static void OnPromoteToParentClick(IRibbonControl? control, CommandBoundary boundary, Action command, Action? onCompleted = null)
        => OnAddRowClick(control, boundary, command, nameof(OnPromoteToParentClick), onCompleted);

    /// <summary>
    /// Called when the user clicks Demote to child. The callback is a thin error
    /// boundary, exactly like the add-row clicks: it resolves the command name and
    /// delegates to <see cref="CommandBoundary"/>, which produces one log record
    /// and one user-safe dialog on an unexpected failure. A typed planner or
    /// mutation refusal is presented by the command itself (R5.3 D6); the ribbon
    /// state is refreshed and invalidated after every run (R1.5 D3).
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnDemoteToChildClick(IRibbonControl control) =>
        OnDemoteToChildClick(control, CommandBoundary.Instance, DemoteToChildCommand.RunForExcel, NotifyRibbonStateChanged);

    /// <summary>
    /// Runs <see cref="OnDemoteToChildClick(IRibbonControl)"/> against an
    /// injected boundary. Internal so contract tests can verify the routing
    /// without the session singleton.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="boundary">The command error boundary.</param>
    /// <param name="command">The demote command to run.</param>
    /// <param name="onCompleted">The post-command state hook.</param>
    internal static void OnDemoteToChildClick(IRibbonControl? control, CommandBoundary boundary, Action command, Action? onCompleted = null)
        => OnAddRowClick(control, boundary, command, nameof(OnDemoteToChildClick), onCompleted);

    /// <summary>
    /// Called when the user clicks Move up. The callback is a thin error
    /// boundary, exactly like the promote and demote clicks; the move
    /// command itself owns outcome-to-message translation and stays silent
    /// at the boundary of the selection's set (R5.4 D4). The ribbon state is
    /// refreshed and invalidated after every run (R1.5 D3).
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnMoveUpClick(IRibbonControl control) =>
        OnMoveUpClick(control, CommandBoundary.Instance, MoveUpCommand.RunForExcel, NotifyRibbonStateChanged);

    /// <summary>
    /// Runs <c>OnMoveUpClick</c> against an injected boundary. Internal so
    /// contract tests can verify the routing without the session singleton.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="boundary">The command error boundary.</param>
    /// <param name="command">The move-up command to run.</param>
    /// <param name="onCompleted">The post-command state hook.</param>
    internal static void OnMoveUpClick(IRibbonControl? control, CommandBoundary boundary, Action command, Action? onCompleted = null)
        => OnAddRowClick(control, boundary, command, nameof(OnMoveUpClick), onCompleted);

    /// <summary>
    /// Called when the user clicks Move down: the mirror of
    /// <c>OnMoveUpClick</c>, silent at the bottom of the selection's
    /// set (R5.4 D4).
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnMoveDownClick(IRibbonControl control) =>
        OnMoveDownClick(control, CommandBoundary.Instance, MoveDownCommand.RunForExcel, NotifyRibbonStateChanged);

    /// <summary>
    /// Runs <c>OnMoveDownClick</c> against an injected boundary. Internal so
    /// contract tests can verify the routing without the session singleton.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="boundary">The command error boundary.</param>
    /// <param name="command">The move-down command to run.</param>
    /// <param name="onCompleted">The post-command state hook.</param>
    internal static void OnMoveDownClick(IRibbonControl? control, CommandBoundary boundary, Action command, Action? onCompleted = null)
        => OnAddRowClick(control, boundary, command, nameof(OnMoveDownClick), onCompleted);

    /// <summary>
    /// Called when the user clicks Remove: the destructive remove
    /// command (R5.4 D6) confirms before the mutator runs, so the
    /// callback is the same thin error boundary as the other
    /// mutating clicks. A declined confirmation never reaches the
    /// boundary.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnRemoveClick(IRibbonControl control) =>
        OnRemoveClick(control, CommandBoundary.Instance, RemoveCommand.RunForExcel, NotifyRibbonStateChanged);

    /// <summary>
    /// Runs <see cref="OnRemoveClick(IRibbonControl)"/> against an
    /// injected boundary. Internal so contract tests can verify the
    /// routing without the session singleton.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="boundary">The command error boundary.</param>
    /// <param name="command">The remove command to run.</param>
    /// <param name="onCompleted">The post-command state hook.</param>
    internal static void OnRemoveClick(IRibbonControl? control, CommandBoundary boundary, Action command, Action? onCompleted = null)
        => OnAddRowClick(control, boundary, command, nameof(OnRemoveClick), onCompleted);

    /// <summary>
    /// Called when the user clicks Expand all groups: every parent/child group
    /// in the table is expanded at once (R5.4). The command is silent on
    /// success and on a sheet with no group; only a refusal is presented.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnExpandAllGroupsClick(IRibbonControl control) =>
        OnExpandAllGroupsClick(
            control,
            CommandBoundary.Instance,
            () => CollapseGroupsCommand.RunForExcel(OutlineCollapseDirection.Expand),
            NotifyRibbonStateChanged);

    /// <summary>
    /// Runs <see cref="OnExpandAllGroupsClick(IRibbonControl)"/> against an
    /// injected boundary. Internal so contract tests can verify the routing
    /// without the session singleton.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="boundary">The command error boundary.</param>
    /// <param name="command">The expand-all command to run.</param>
    /// <param name="onCompleted">The post-command state hook.</param>
    internal static void OnExpandAllGroupsClick(IRibbonControl? control, CommandBoundary boundary, Action command, Action? onCompleted = null)
        => OnAddRowClick(control, boundary, command, nameof(OnExpandAllGroupsClick), onCompleted);

    /// <summary>
    /// Called when the user clicks Collapse all groups: every parent/child
    /// group in the table is collapsed at once (R5.4). The mirror of
    /// <see cref="OnExpandAllGroupsClick(IRibbonControl)"/>.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnCollapseAllGroupsClick(IRibbonControl control) =>
        OnCollapseAllGroupsClick(
            control,
            CommandBoundary.Instance,
            () => CollapseGroupsCommand.RunForExcel(OutlineCollapseDirection.Collapse),
            NotifyRibbonStateChanged);

    /// <summary>
    /// Runs <see cref="OnCollapseAllGroupsClick(IRibbonControl)"/> against an
    /// injected boundary. Internal so contract tests can verify the routing
    /// without the session singleton.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="boundary">The command error boundary.</param>
    /// <param name="command">The collapse-all command to run.</param>
    /// <param name="onCompleted">The post-command state hook.</param>
    internal static void OnCollapseAllGroupsClick(IRibbonControl? control, CommandBoundary boundary, Action command, Action? onCompleted = null)
        => OnAddRowClick(control, boundary, command, nameof(OnCollapseAllGroupsClick), onCompleted);

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
    /// Returns the index of the selected row in the plot time-scale dropdown.
    /// The Month scale owns two rows because the month band has two label
    /// forms, so the index is derived from the stored scale/format PAIR, not
    /// from the scale alone. Never throws.
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
    /// <returns>The zero-based index of the selected row.</returns>
    internal static int GetTimeScale(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return TimeScaleSelectionIndex(
                stateService.GetTimeScale(),
                stateService.GetPeriodLabelFormat());
        }
        catch
        {
            return TimeScaleSelectionIndex(GanttTimeScale.Month, GanttPeriodLabelFormat.MMM);
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Called when the user picks a row from the plot time-scale dropdown.
    /// Excel's dropDown onAction contract passes the control, the selected item
    /// id, and the selected index; the id names a scale and, for the month
    /// band, a label form, and both are committed together then persisted. An
    /// unknown or incompatible id is refused by the parsers rather than
    /// defaulted, so a mistyped selection cannot silently redraw the period
    /// band with the wrong calendar unit or label. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    /// <param name="selectedId">
    /// The selected item id (Month-MM, Month-MMM, Quarter, or Week).
    /// </param>
    /// <param name="selectedIndex">The selected item index (unused; the id is authoritative).</param>
#pragma warning disable IDE0060 // The index is part of Excel's dropDown onAction contract; Excel always passes it.
    public void OnTimeScaleChange(IRibbonControl control, string selectedId, int selectedIndex)
        => OnTimeScaleChange(control, selectedId, RibbonStateService.Instance);
#pragma warning restore IDE0060

    /// <summary>
    /// Runs <see cref="OnTimeScaleChange(IRibbonControl, string, int)"/>
    /// against an injected state service. Internal so contract tests can verify
    /// the routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="id">The selected item id, or null in tests.</param>
    /// <param name="stateService">The state service receiving the setting.</param>
    internal static void OnTimeScaleChange(IRibbonControl? control, string? id, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
        if (TryParseTimeScaleSelection(id, out GanttTimeScale scale, out GanttPeriodLabelFormat? format))
        {
            stateService.SetTimeScale(scale, format);
        }
    }

    /// <summary>
    /// Maps the stored scale/format pair to the dropdown's selected index.
    /// This is the exact inverse of <see cref="TryParseTimeScaleSelection"/>
    /// plus the item order declared in Ribbon.xml; both directions are pinned
    /// by contract tests so the getter cannot drift from the document.
    /// </summary>
    private static int TimeScaleSelectionIndex(GanttTimeScale scale, GanttPeriodLabelFormat format) =>
        scale switch
        {
            GanttTimeScale.Month when format == GanttPeriodLabelFormat.MM => 0,
            GanttTimeScale.Month => 1,
            GanttTimeScale.Quarter => 2,
            GanttTimeScale.Week => 3,
            _ => 1,
        };

    /// <summary>
    /// Parses a dropdown item id into the scale and label format it selects.
    /// The id is the scale name, optionally suffixed with the label format
    /// after a hyphen (<c>Month-MM</c>, <c>Month-MMM</c>); a suffix-less id
    /// (<c>Quarter</c>, <c>Week</c>) names the scale only and leaves the format
    /// to the service's canonical rule. Both halves go through the closed Core
    /// parsers, and an incompatible suffix is REFUSED rather than coerced, so
    /// the stored pair is always one the dropdown could have produced. Never
    /// throws.
    /// </summary>
    /// <param name="id">The selected item id, or null in tests.</param>
    /// <param name="scale">The parsed scale when successful.</param>
    /// <param name="format">
    /// The explicitly named format, or null when the id names only the scale.
    /// </param>
    /// <returns><see langword="true"/> when the id names a valid selection.</returns>
    private static bool TryParseTimeScaleSelection(
        string? id,
        out GanttTimeScale scale,
        out GanttPeriodLabelFormat? format)
    {
        scale = default;
        format = null;

        if (string.IsNullOrEmpty(id))
        {
            return false;
        }

        int separator = id.IndexOf('-', StringComparison.Ordinal);
        if (!GanttChartSettings.TryParseTimeScale(
            separator < 0 ? id : id[..separator],
            out scale))
        {
            return false;
        }

        if (separator < 0)
        {
            return true;
        }

        if (!GanttChartSettings.TryParsePeriodLabelFormat(id[(separator + 1)..], out GanttPeriodLabelFormat parsed)
            || !GanttChartSettings.IsCompatible(scale, parsed))
        {
            return false;
        }

        format = parsed;
        return true;
    }

    /// <summary>
    /// Returns how many items the description-label dropdown offers. Excel's
    /// dynamic dropDown contract asks for this through getItemCount; the answer
    /// comes from the state service snapshot (the selected row's projected
    /// positions plus the leading Default item). Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public int GetDescriptionLabelCount(IRibbonControl control) =>
        GetDescriptionLabelCount(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetDescriptionLabelCount(IRibbonControl)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The item count; zero when no single row is selected.</returns>
    internal static int GetDescriptionLabelCount(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.GetDescriptionLabelCount();
        }
        catch
        {
            return 0;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Returns the item ID at the given index for the description-label dropdown
    /// (<c>Default</c> at 0, otherwise the position name). Excel's dynamic
    /// dropDown contract asks for this through getItemID; the answer comes from
    /// the state service snapshot. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    /// <param name="index">The zero-based item index.</param>
    public string GetDescriptionLabelItemID(IRibbonControl control, int index) =>
        GetDescriptionLabelItemID(control, index, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetDescriptionLabelItemID(IRibbonControl, int)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="index">The zero-based item index.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The item ID, or an empty string when the index is out of range.</returns>
    internal static string GetDescriptionLabelItemID(
        IRibbonControl? control,
        int index,
        RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.GetDescriptionLabelItemID(index);
        }
        catch
        {
            return string.Empty;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Returns the user-facing item label at the given index for the
    /// description-label dropdown. Excel's dynamic dropDown contract asks for this
    /// through getItemLabel; the answer comes from the state service snapshot,
    /// so the enum name is never shown verbatim (R5.5 ruling 3). Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    /// <param name="index">The zero-based item index.</param>
    public string GetDescriptionLabelItemLabel(IRibbonControl control, int index) =>
        GetDescriptionLabelItemLabel(control, index, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetDescriptionLabelItemLabel(IRibbonControl, int)"/> against
    /// an injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="index">The zero-based item index.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The item label, or an empty string when the index is out of range.</returns>
    internal static string GetDescriptionLabelItemLabel(
        IRibbonControl? control,
        int index,
        RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.GetDescriptionLabelItemLabel(index);
        }
        catch
        {
            return string.Empty;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Returns the selected index of the description-label dropdown. Excel's
    /// dynamic dropDown contract asks for this through
    /// getSelectedItemIndex; the answer comes from the state service snapshot,
    /// so the dropdown reflects the selected row's stored cell. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public int GetDescriptionLabelSelectedIndex(IRibbonControl control) =>
        GetDescriptionLabelSelectedIndex(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetDescriptionLabelSelectedIndex(IRibbonControl)"/> against
    /// an injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The zero-based selected index; zero (Default) when unknown.</returns>
    internal static int GetDescriptionLabelSelectedIndex(
        IRibbonControl? control,
        RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.GetDescriptionLabelSelectedIndex();
        }
        catch
        {
            return 0;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Returns how many items the date-label dropdown offers. Excel's
    /// dynamic dropDown contract asks for this through getItemCount; the
    /// answer comes from the state service snapshot (the selected row's
    /// allowed date-label choices, which already include Auto). Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public int GetDateLabelCount(IRibbonControl control) =>
        GetDateLabelCount(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetDateLabelCount(IRibbonControl)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The item count; zero when no single row is selected.</returns>
    internal static int GetDateLabelCount(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.GetDateLabelCount();
        }
        catch
        {
            return 0;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Returns the item ID at the given index for the date-label dropdown
    /// (the leading Auto choice carries the Default clear-ID; every other
    /// choice carries its enum name). Excel's dynamic dropDown contract asks
    /// for this through getItemID; the answer comes from the state service
    /// snapshot. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    /// <param name="index">The zero-based item index.</param>
    public string GetDateLabelItemID(IRibbonControl control, int index) =>
        GetDateLabelItemID(control, index, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetDateLabelItemID(IRibbonControl, int)"/> against
    /// an injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="index">The zero-based item index.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The item ID, or an empty string when the index is out of range.</returns>
    internal static string GetDateLabelItemID(
        IRibbonControl? control,
        int index,
        RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.GetDateLabelItemID(index);
        }
        catch
        {
            return string.Empty;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Returns the user-facing item label at the given index for the
    /// date-label dropdown. Excel's dynamic dropDown contract asks for this
    /// through getItemLabel; the answer comes from the state service
    /// snapshot, so the enum name is never shown verbatim. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    /// <param name="index">The zero-based item index.</param>
    public string GetDateLabelItemLabel(IRibbonControl control, int index) =>
        GetDateLabelItemLabel(control, index, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetDateLabelItemLabel(IRibbonControl, int)"/> against
    /// an injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="index">The zero-based item index.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The item label, or an empty string when the index is out of range.</returns>
    internal static string GetDateLabelItemLabel(
        IRibbonControl? control,
        int index,
        RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.GetDateLabelItemLabel(index);
        }
        catch
        {
            return string.Empty;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Returns the selected index of the date-label dropdown. Excel's
    /// dynamic dropDown contract asks for this through
    /// getSelectedItemIndex; the answer comes from the state service
    /// snapshot, so the dropdown reflects the selected row's stored cell.
    /// Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public int GetDateLabelSelectedIndex(IRibbonControl control) =>
        GetDateLabelSelectedIndex(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetDateLabelSelectedIndex(IRibbonControl)"/> against
    /// an injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The zero-based selected index; zero (Auto) when unknown.</returns>
    internal static int GetDateLabelSelectedIndex(
        IRibbonControl? control,
        RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.GetDateLabelSelectedIndex();
        }
        catch
        {
            return 0;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Handles a date-label dropdown selection change. Excel's
    /// dropDown onAction contract passes the control, the selected item
    /// id, and the selected index; the id is authoritative (it names the
    /// position, or <c>Default</c> to clear the cell) and the index is
    /// ignored, exactly as the description-label dropdown works. The
    /// state service parses the id and writes the selected row's
    /// date-label cell; an unknown id is refused rather than defaulted.
    /// Never throws.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    /// <param name="selectedId">The selected item id, or null in tests.</param>
    /// <param name="selectedIndex">The selected item index (unused; the id is authoritative).</param>
#pragma warning disable IDE0060 // The index is part of Excel's dropDown onAction contract; Excel always passes it.
    public void OnDateLabelChange(IRibbonControl control, string selectedId, int selectedIndex)
        => OnDateLabelChange(control, selectedId, RibbonStateService.Instance);
#pragma warning restore IDE0060

    /// <summary>
    /// Runs <see cref="OnDateLabelChange(IRibbonControl, string, int)"/>
    /// against an injected state service. Internal so contract tests can
    /// verify the routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="id">The selected item id, or null in tests.</param>
    /// <param name="stateService">The state service applying the selection.</param>
    internal static void OnDateLabelChange(
        IRibbonControl? control,
        string? id,
        RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            stateService.SetDateLabelPosition(id);
        }
        catch
        {
            // Intentionally empty: the command owns its own refusals and
            // presents them; a failure here degrades to no change.
        }
#pragma warning restore CA1031
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
    /// Called when the user picks a description-label position from the ribbon dropdown.
    /// Excel's dropDown onAction contract passes the control, the selected item
    /// id, and the selected index; the id is authoritative (it names the
    /// position, or <c>Default</c> to clear the cell) and the index is ignored,
    /// exactly as the time-scale and margin dropdowns work. The command writes
    /// the selected row's <c>LabelPosition</c> cell; an unknown id is refused
    /// rather than defaulted. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    /// <param name="selectedId">The selected item id, or null in tests.</param>
    /// <param name="selectedIndex">The selected item index (unused; the id is authoritative).</param>
#pragma warning disable IDE0060 // The index is part of Excel's dropDown onAction contract; Excel always passes it.
    public void OnDescriptionLabelChange(IRibbonControl control, string selectedId, int selectedIndex)
        => OnDescriptionLabelChange(control, selectedId, RibbonStateService.Instance);
#pragma warning restore IDE0060

    /// <summary>
    /// Runs <see cref="OnDescriptionLabelChange(IRibbonControl, string, int)"/>
    /// against an injected state service. Internal so contract tests can verify
    /// the routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="id">The selected item id, or null in tests.</param>
    /// <param name="stateService">The state service applying the setting.</param>
    internal static void OnDescriptionLabelChange(
        IRibbonControl? control,
        string? id,
        RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            stateService.SetDescriptionLabelPosition(id);
        }
        catch
        {
            // Intentionally empty: the command owns its own refusals and
            // presents them; a failure here degrades to no change.
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Returns how many items the custom-style dropdown offers (R5.6a): the
    /// user-created style keys from the state service snapshot. Excel's
    /// dynamic dropDown contract asks for this through getItemCount. Never
    /// throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public int GetCustomStyleCount(IRibbonControl control) =>
        GetCustomStyleCount(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetCustomStyleCount(IRibbonControl)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The item count; zero when no single style-key row is selected.</returns>
    internal static int GetCustomStyleCount(IRibbonControl? control, RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.GetCustomStyleCount();
        }
        catch
        {
            return 0;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Returns the item ID at the given index for the custom-style dropdown
    /// (the style key itself). Excel's dynamic dropDown contract asks for
    /// this through getItemID. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    /// <param name="index">The zero-based item index.</param>
    public string GetCustomStyleItemID(IRibbonControl control, int index) =>
        GetCustomStyleItemID(control, index, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetCustomStyleItemID(IRibbonControl, int)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="index">The zero-based item index.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The item ID, or an empty string when the index is out of range.</returns>
    internal static string GetCustomStyleItemID(
        IRibbonControl? control,
        int index,
        RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.GetCustomStyleItemID(index);
        }
        catch
        {
            return string.Empty;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Returns the user-facing item label at the given index for the
    /// custom-style dropdown (the style key, shown verbatim). Excel's
    /// dynamic dropDown contract asks for this through getItemLabel. Never
    /// throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    /// <param name="index">The zero-based item index.</param>
    public string GetCustomStyleItemLabel(IRibbonControl control, int index) =>
        GetCustomStyleItemLabel(control, index, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetCustomStyleItemLabel(IRibbonControl, int)"/> against
    /// an injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="index">The zero-based item index.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The item label, or an empty string when the index is out of range.</returns>
    internal static string GetCustomStyleItemLabel(
        IRibbonControl? control,
        int index,
        RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.GetCustomStyleItemLabel(index);
        }
        catch
        {
            return string.Empty;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Returns the selected index of the custom-style dropdown. Excel's
    /// dynamic dropDown contract asks for this through
    /// getSelectedItemIndex; the answer comes from the state service
    /// snapshot, so the dropdown reflects the selected row's stored cell.
    /// Never throws.
    /// </summary>
    /// <param name="control">The ribbon control Excel is asking about.</param>
    public int GetCustomStyleSelectedIndex(IRibbonControl control) =>
        GetCustomStyleSelectedIndex(control, RibbonStateService.Instance);

    /// <summary>
    /// Runs <see cref="GetCustomStyleSelectedIndex(IRibbonControl)"/> against
    /// an injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="stateService">The state service answering from its snapshot.</param>
    /// <returns>The zero-based selected index; 0 when no custom style is assigned
    /// or the service cannot answer.</returns>
    internal static int GetCustomStyleSelectedIndex(
        IRibbonControl? control,
        RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            return stateService.GetCustomStyleSelectedIndex();
        }
        catch
        {
            return 0;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Called when the user picks a custom style (R5.6a). The item ID is the
    /// style key; an unknown key is refused by the command at execution
    /// rather than defaulted. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    /// <param name="selectedId">The selected item id (the style key), or null in tests.</param>
    /// <param name="selectedIndex">The selected item index (unused; the id is authoritative).</param>
#pragma warning disable IDE0060 // The index is part of Excel's dropDown onAction contract; Excel always passes it.
    public void OnCustomStyleChange(IRibbonControl control, string selectedId, int selectedIndex)
        => OnCustomStyleChange(control, selectedId, RibbonStateService.Instance);
#pragma warning restore IDE0060

    /// <summary>
    /// Runs <see cref="OnCustomStyleChange(IRibbonControl, string, int)"/>
    /// against an injected state service. Internal so contract tests can verify
    /// the routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="id">The selected item id (the style key), or null in tests.</param>
    /// <param name="stateService">The state service applying the setting.</param>
    internal static void OnCustomStyleChange(
        IRibbonControl? control,
        string? id,
        RibbonStateService stateService)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
#pragma warning disable CA1031
        try
        {
            stateService.SetCustomStyleKey(id);
        }
        catch
        {
            // Intentionally empty: the command owns its own refusals and
            // presents them; a failure here degrades to no change.
        }
#pragma warning restore CA1031
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
        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double cm))
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
    /// Called when the user clicks the Theme settings button. Opens the
    /// theme settings dialog and applies edits if confirmed. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnThemeSettingsClick(IRibbonControl control)
        => OnThemeSettingsClick(control, RibbonStateService.Instance, NotifyRibbonStateChanged);

    /// <summary>
    /// Runs <see cref="OnThemeSettingsClick(IRibbonControl)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="stateService">The state service running the command.</param>
    /// <param name="onCompleted">
    /// The post-command state hook. The Theme settings dialog can create a
    /// custom style, and the ribbon's custom-style dropdown projects its
    /// items from the workbook's style registry, so the ribbon state must be
    /// re-captured and invalidated after the write. Without this hook the
    /// dropdown kept the list it had before the dialog opened, and a style
    /// the user had just created never appeared in it.
    /// </param>
    /// <param name="refresh">
    /// The chart refresh that follows the write. A seam so a contract test
    /// can drive the command without a live Excel host.
    /// </param>
    internal static void OnThemeSettingsClick(
        IRibbonControl? control,
        RibbonStateService stateService,
        Action? onCompleted = null,
        Action? refresh = null)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
        if (stateService.ShowThemeSettingsDialog())
        {
            (refresh ?? RunRefreshSheet)();
            onCompleted?.Invoke();
        }
    }

    /// <summary>
    /// Called when the user clicks the Chart title button. Opens the
    /// chart-title dialog and applies edits if confirmed. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
    public void OnChartTitleClick(IRibbonControl control)
        => OnChartTitleClick(control, RibbonStateService.Instance, NotifyRibbonStateChanged);

    /// <summary>
    /// Runs <see cref="OnChartTitleClick(IRibbonControl)"/> against an
    /// injected state service. Internal so contract tests can verify the
    /// routing without the session singleton. Never throws.
    /// </summary>
    /// <param name="control">The ribbon control, or null when unavailable.</param>
    /// <param name="stateService">The state service running the command.</param>
    /// <param name="onCompleted">
    /// The post-command state hook. The chart title is a setting the ribbon
    /// reads, so the state must be re-captured after the write — the same
    /// contract every other ribbon command follows.
    /// </param>
    /// <param name="refresh">
    /// The chart refresh that follows the write. A seam so a contract test
    /// can drive the command without a live Excel host.
    /// </param>
    internal static void OnChartTitleClick(
        IRibbonControl? control,
        RibbonStateService stateService,
        Action? onCompleted = null,
        Action? refresh = null)
    {
        _ = control;
        ArgumentNullException.ThrowIfNull(stateService);
        if (stateService.ShowChartTitleDialog())
        {
            (refresh ?? RunRefreshSheet)();
            onCompleted?.Invoke();
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
        string? id = control?.Id;
        return string.IsNullOrWhiteSpace(id) ? fallbackCommandName : id;
    }
}

