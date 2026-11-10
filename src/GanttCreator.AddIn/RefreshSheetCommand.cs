using System.Diagnostics;
using System.Globalization;
using GanttCreator.Core.Scene;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// The Ribbon entry point for Refresh: it checks the orchestrator is available,
/// invokes it exactly once, and presents whatever it returns (R4.8A D1).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this holds no orchestration.</b> The pipeline order, the
/// validate-before-mutate rule, and the step ownership all live in
/// <see cref="IGanttRefreshOrchestrator"/>. A command that also sequenced the steps
/// would be a second copy of that order, and the copy is the one that would drift —
/// the callback is the file people edit when adding a feature, and the service is the
/// file they forget exists.
/// </para>
/// <para>
/// The command's whole job is the boundary work: resolve the live dependencies,
/// call once, and translate the outcome into a user-visible presentation. A
/// presenter failure is contained so a dialog error can never escape into Excel.
/// </para>
/// </remarks>
internal static class RefreshSheetCommand
{
    private static readonly ShapeDeltaCache _shapeDeltaCache = new();

    /// <summary>
    /// The session's arming memory (ADR-0055). It must outlive one orchestrator, because the live
    /// composition builds a new one for every refresh.
    /// </summary>
    private static readonly ProtectionArmingState _protectionArming = ProtectionArmingState.Shared;

    /// <summary>
    /// The session's protection-ownership flag, shared with every adapter's guard so the add-in's
    /// own protection does not block its own writes.
    /// </summary>
    private static readonly ProtectionOwnershipSource _protectionOwnership = ProtectionOwnershipSource.Shared;

    /// <summary>
    /// The process's one release window, so nesting is by depth across commands: Add activity opens
    /// it, the refresh it triggers opens it again, and only the outermost scope re-protects.
    /// </summary>
    private static WorksheetProtectionWindow? _protectionWindow;

    /// <summary>Resolves the release window for this Excel session (ADR-0055).</summary>
    /// <param name="application">The Excel application object.</param>
    /// <returns>The shared window.</returns>
    /// <remarks>
    /// One instance per process because the depth is the mechanism: a per-call window would let the
    /// refresh re-protect the sheet while the command that triggered it still has to write. The
    /// application object is a stable Excel-owned root for the session, so the port built from the
    /// first one stays valid.
    /// </remarks>
    private static WorksheetProtectionWindow ProtectionWindowFor(object? application) =>
        _protectionWindow ??= new WorksheetProtectionWindow(
            // The same sink the window uses, so a live log carries the outline-symbol state a
            // protect found and left next to the window's own open/close line. Both refresh paths
            // share this one window, so the surrounding lines are what separate them.
            new ExcelWorksheetProtectionPort(application, CommandBoundary.Instance.WriteTechnicalRecord),
            _protectionArming,
            CommandBoundary.Instance.WriteTechnicalRecord);

    /// <summary>
    /// Opens the ADR-0055 release window for a command that changes the sheet's structure or shapes.
    /// </summary>
    /// <returns>The scope that re-applies protection when the command finishes.</returns>
    /// <remarks>
    /// The row commands share the refresh's window instance, so a command that inserts a row and then
    /// triggers a render releases once and protects once, at the outermost boundary. Cell-only
    /// commands do not need this: the live measurement shows a protected sheet still accepts cell
    /// writes.
    /// </remarks>
    internal static IDisposable OpenProtectionWindow() =>
        ProtectionWindowFor(ExcelDna.Integration.ExcelDnaUtil.Application).Open();

    internal static void ResetDeltaCache() => _shapeDeltaCache.Invalidate();

    /// <summary>
    /// Forgets that protection is armed, so the next refresh re-arms it.
    /// </summary>
    /// <remarks>
    /// Called when the active workbook changes. The workbook the user switched to may be protected
    /// from an earlier session, whose <c>UserInterfaceOnly</c> did not survive the reopen, so it
    /// refuses this add-in's writes until protection is re-applied. It is also called before the
    /// next refresh after the user opts out, so the release path runs.
    /// </remarks>
    internal static void ResetProtectionArming() => _protectionArming.Invalidate();

    /// <summary>
    /// Re-applies the outline-symbol option on a sheet that is already protected, so the native
    /// <c>+</c>/<c>-</c> controls work as soon as a saved workbook is reopened.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <c>EnableOutlining</c> is refused while a sheet is protected and is not saved with the
    /// worksheet, so a reopened workbook comes up protected with the symbols disabled, and no
    /// protect the add-in performs could restore them until a protect happened to run on an
    /// unprotected sheet. The protect path now cycles protection when the sheet is already
    /// protected, and this is the call that triggers it at the one moment it matters: workbook
    /// activate, before any refresh, because a refresh may refuse early and leave the user looking
    /// at greyed-out buttons with no way to fold a section.
    /// </para>
    /// <para>
    /// <b>It touches nothing when nothing needs repairing.</b> An unprotected sheet returns
    /// immediately: there is no option to restore, and the next protect sets it. A sheet protected
    /// with a password this add-in does not have is refused by the release inside the cycle, which
    /// is caught there, so the sheet is left exactly as the user had it. Nothing here can throw.
    /// </para>
    /// </remarks>
    internal static void RepairOutlineSymbols()
    {
        try
        {
            object? application = ExcelDna.Integration.ExcelDnaUtil.Application;
            if (application is null)
            {
                return;
            }

            // The same sink the window uses, so a repair is visible in the log next to the
            // protection lines it produces.
            ExcelWorksheetProtectionPort port = new(application, CommandBoundary.Instance.WriteTechnicalRecord);
            WorksheetProtectionState? state = port.Read();
            if (state is null || !state.ContentsProtected)
            {
                return;
            }

            // Protect cycles the protection (release -> set the option -> re-protect), which is the
            // only way to set a property the host refuses in the current state.
            _ = port.Protect();
        }
#pragma warning disable CA1031
        catch
        {
            // Workbook activate runs on the Excel UI thread and must never surface a failure;
            // the log already carries whatever the port decided.
        }
#pragma warning restore CA1031
    }

    /// <summary>Runs the production refresh in the current Excel session.</summary>
    internal static void RunForExcel()
    {
        object? application = ExcelDna.Integration.ExcelDnaUtil.Application;

        // ADR-0055 clause 7: an explicit Refresh is the user asking for the chart to be rebuilt from
        // the scene, so it re-asserts every shape's geometry and re-applies the user's toggles. It
        // does NOT reset them — an unlocked chart stays unlocked, because the toggle is the user's
        // durable choice rather than a transient manual-edit window.

        // ADR-0055: the write phase runs with the add-in's protection released, because a protected
        // sheet refuses shape creation and movement (measured; see ProtectionArmingIntegrationTests).
        using IDisposable window = ProtectionWindowFor(application).Open();

        GanttRefreshOrchestrator orchestrator = CreateLiveOrchestrator(application, useDeltaCache: false);
        GanttRefreshOutcome outcome = orchestrator.Refresh();
        long refreshCompleted = Stopwatch.GetTimestamp();
        RibbonStateService.Instance.UpdateCurrentValidationIssues(outcome.ValidationIssues);

        // Store scene warnings in the Ribbon state service so the warnings
        // panel can surface them. Row-level validation issues are reported
        // through row/cell fill colours; scene-level warnings surface here.
        if (outcome.Succeeded)
        {
            RibbonStateService.Instance.UpdateWarnings(outcome.Warnings);
            RibbonStateService.ClearStatusBar();
        }

        if (!outcome.Succeeded)
        {
            Present(CommandErrorDialog.Show, outcome.Message ?? "The chart could not be refreshed.");
        }

        TracePostPass(refreshCompleted);
    }

    /// <summary>
    /// Runs a silent production refresh: status bar only, no modal dialogs.
    /// Used by the on-change refresh scheduler so an automatic refresh never
    /// interrupts the user with a dialog. Follows the same sequencing as the
    /// manual refresh: orchestrator pipeline, warning capture, then targeted
    /// ribbon control invalidation so selection-gated controls
    /// (MoveUp/MoveDown/PromoteToParent/DemoteToChild/AddChild) re-evaluate
    /// against the live selection after the auto-refresh rewrites the sheet.
    /// Does <b>not</b> call <see cref="RibbonStateService.Refresh"/> or
    /// RibbonStateService.Invalidate — those re-read every table row and
    /// re-query every ribbon getter, which is unnecessary when the selection
    /// has not changed.
    /// </summary>
    internal static void RunSilentForExcel()
    {
        object? application = ExcelDna.Integration.ExcelDnaUtil.Application;

        // ADR-0055: same window as the manual path, so a scheduled refresh can create and move shapes
        // on a sheet the add-in owns.
        using IDisposable window = ProtectionWindowFor(application).Open();

        GanttRefreshOrchestrator orchestrator = CreateLiveOrchestrator(application, useDeltaCache: true);
        GanttRefreshOutcome outcome = orchestrator.Refresh();
        long refreshCompleted = Stopwatch.GetTimestamp();
        RibbonStateService.Instance.UpdateCurrentValidationIssues(outcome.ValidationIssues);

        // Store scene warnings in the Ribbon state service — same as RunForExcel.
        if (outcome.Succeeded)
        {
            RibbonStateService.Instance.UpdateWarnings(outcome.Warnings);
            RibbonStateService.ClearStatusBar();
            // Invalidate only the selection-dependent controls so their enabled
            // state re-evaluates against the post-refresh sheet state. Mirrors
            // NotifyRibbonStateChanged() on the manual refresh path, but scoped
            // to the controls that read SelectionBoundaryContext / AddChildContext
            // instead of re-reading every table row and re-querying every ribbon
            // getter.
            RibbonStateService.Instance.InvalidateControl(RibbonControlIds.DescriptionLabel);
            RibbonStateService.Instance.InvalidateControl(RibbonControlIds.DateLabel);
            RibbonStateService.Instance.InvalidateControl(RibbonControlIds.CustomStyle);
            RibbonStateService.Instance.InvalidateControl(RibbonControlIds.MoveUp);
            RibbonStateService.Instance.InvalidateControl(RibbonControlIds.MoveDown);
            RibbonStateService.Instance.InvalidateControl(RibbonControlIds.PromoteToParent);
            RibbonStateService.Instance.InvalidateControl(RibbonControlIds.DemoteToChild);
            RibbonStateService.Instance.InvalidateControl(RibbonControlIds.AddChild);
            TracePostPass(refreshCompleted);
            return;
        }

        Present(RibbonStatusBar.Warn, outcome.Message ?? "The chart could not be refreshed.");
        TracePostPass(refreshCompleted);
    }

    /// <summary>Runs an injected refresh, for tests and for the R4.9 wiring.</summary>
    /// <param name="orchestrator">The refresh service to invoke.</param>
    /// <param name="presenter">The user-facing presenter for a non-success outcome.</param>
    internal static void Run(IGanttRefreshOrchestrator orchestrator, Action<string> presenter)
    {
        ArgumentNullException.ThrowIfNull(orchestrator);
        ArgumentNullException.ThrowIfNull(presenter);

        // Exactly one call. A retry loop here would re-run the whole pipeline against
        // a chart the first attempt already partly wrote, which is how a transient
        // COM refusal becomes a corrupted sheet. The user refreshes again instead.
        GanttRefreshOutcome outcome = orchestrator.Refresh();
        if (outcome.Succeeded)
        {
            return;
        }

        Present(presenter, outcome.Message ?? "The chart could not be refreshed.");
    }

    /// <summary>
    /// Records how long the work AFTER the refresh pipeline took, in milliseconds.
    /// </summary>
    /// <param name="since">The <see cref="Stopwatch"/> timestamp taken when the pipeline returned.</param>
    /// <remarks>
    /// The click-to-plot time the user feels also includes the ribbon invalidation, the
    /// warning and status presentation, and the state restore that follow the pipeline.
    /// Those sit outside the pipeline's own <c>gantt.refresh result</c> line, so without
    /// this figure they are invisible and every later optimisation is attributed to the
    /// pipeline whether or not it belongs there. The sink degrades internally and never
    /// throws, so this cannot abort a refresh that has already succeeded.
    /// </remarks>
    private static void TracePostPass(long since) =>
        CommandBoundary.Instance.WriteTechnicalRecord(string.Create(
            CultureInfo.InvariantCulture,
            $"gantt.refresh post ms={Stopwatch.GetElapsedTime(since).TotalMilliseconds:0.#}"));

    /// <summary>
    /// Shows the message, degrading to no dialog if the host refuses one.
    /// </summary>
    /// <param name="presenter">The user-facing presenter.</param>
    /// <param name="message">The message to show.</param>
    private static void Present(Action<string> presenter, string message)
    {
        // CA1031: a presenter failure is outside command behaviour; it degrades to no
        // dialog. The refresh has already happened, so escaping would surface an error
        // for work that was done.
#pragma warning disable CA1031
        try
        {
            presenter(message);
        }
        catch
        {
            // Intentionally empty: a dialog failure must not escape into Excel.
        }
#pragma warning restore CA1031
    }

    // CA2000: the scope is handed to the orchestrator, which OWNS it and disposes it
    // in a `using` on every path (R4.8A D8). The analyzer cannot see ownership
    // transfer through a constructor, so the disposal it asks for here would be a
    // second, redundant one - and a scope disposed before the refresh ran would
    // restore nothing, which is the opposite of what is wanted.
#pragma warning disable CA2000
    /// <param name="application">The Excel application object.</param>
    /// <param name="useDeltaCache">Whether this scheduled refresh may skip unchanged payload writes.</param>
    private static GanttRefreshOrchestrator CreateLiveOrchestrator(object? application, bool useDeltaCache)
    {
        // One sink for the refresh's delta/geometry decisions and the row-structure
        // writes in the same run. It goes to the rolling log, which degrades internally
        // and never throws, so a logging failure cannot abort a refresh.
        Action<string> trace = CommandBoundary.Instance.WriteTechnicalRecord;

        // Constructed once and shared: the arming policy locks the owned shapes through the same
        // writer the reconciliation uses (ADR-0055), so both must see the same guard and the same
        // trace sink.
        ExcelShapeWriter shapeWriter =
            new(application, new ExcelWorksheetProtectionGuard(application), trace);

        return new GanttRefreshOrchestrator(
            new ExcelGanttTableReader(application),
            new ExcelConfigCatalogueReader(application),
            new ExcelWorksheetProtectionGuard(application),
            new ExcelPanelGridMeasurement(application),
            new ExcelDurationWriter(application),
            new ExcelRowHeightNormaliser(application, trace: trace),
            new ExcelOutlineGroupWriter(application, trace: trace),

            // The text-metrics seam is a real host dependency: Core never measures,
            // so something must. This is now the FONT-BACKED implementation:
            // AptosTextMetrics reads the real advances of the typography families
            // (Aptos / Aptos Narrow, Regular and Bold), which a live Excel probe
            // measured against the previous flat 4pt-per-character table and found
            // wrong in BOTH directions -- 44 percent too small for 'W' and 105
            // percent too large for 'i'. The small end is what clipped labels,
            // because the shape writer writes AutoSize = msoAutoSizeNone and the
            // host cannot recover the difference. The advances are read from the
            // font binaries by scripts/gen-font-advances.py and committed as data,
            // so the measurement stays deterministic and Office-free.
            new ExcelSceneBuildRequestFactory(new AptosTextMetrics()),
            shapeWriter,

            // R4.8A D2's two remaining preflight steps, supplied here so the
            // orchestrator runs them. Both are optional in the constructor, which is
            // what keeps this additive: a caller that omits them gets the pipeline
            // exactly as it behaved before this row.
            new ExcelColumnPresentationRestorer(application),
            new ExcelGanttRowIdentityRepairer(application),

            // ADR-0020's five-setting scope. It is a single object owned by the
            // orchestrator, which disposes it on every path; constructing one per
            // operation would capture-and-restore the same setting several times and
            // restore the wrong value. The technical-record sink is wired here so a
            // host that refuses one of the five writes (ADR-0020: every member is
            // best-effort) is recorded to the log instead of leaving the degradation
            // silent; a null sink was the previous behaviour.
            new ExcelApplicationStateScope(application, CommandBoundary.Instance.WriteTechnicalRecord),

            // The StyleKey column is engine-owned: every Refresh brings column K
            // into line with each row's current Type, so the writer is part of the
            // live composition, not an optional extra.
            new ExcelStyleKeyWriter(application),

            // The row/cell fill highlighter replaces the cell-notes reporter
            // (ExcelGanttValidationReporter). Fills are cosmetic Interior.ColorIndex
            // writes — no SpecialCells scan, no per-cell Comment read, and no stale
            // note left behind when a row is corrected.
            new ExcelRowWarningHighlighter(application),

            // R5.1: heal type-dependent engine-hidden columns on Type change so
            // switching entity type does not leave stale values the validator
            // rejects. Runs after identity repair and before validation.
            new ExcelTypeChangeNormaliser(application),
            shapeDeltaCache: _shapeDeltaCache,
            useDeltaCache: useDeltaCache,
            trace: trace,

            // ADR-0055: the refresh arms the add-in's own protection (release -> lock -> record) and
            // claims ownership only of protection this session applied. Both the ownership flag and
            // the arming memory are session state, because a new orchestrator is built for every
            // pass; the caller's release window re-applies the protection when the writes finish.
            protectionOwnership: _protectionOwnership,
            protectionArming: new WorksheetProtectionArming(
                // The arming path protects too, so it carries the same diagnostic sink: an arm
                // that greys the symbols out must be visible in the log, not only a window close.
                new ExcelWorksheetProtectionPort(application, trace),
                shapeWriter,
                _protectionArming,
                trace));
    }
#pragma warning restore CA2000
}
