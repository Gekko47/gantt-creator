using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

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
    public override string GetCustomUI(string RibbonID) => RibbonID != "Microsoft.Excel.Workbook" ? null! : RibbonResources.Ribbon;

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
