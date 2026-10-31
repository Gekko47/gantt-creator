using GanttCreator.Core;
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
    /// <summary>
    /// The shared, long-lived <see cref="ChangeTracker"/> owned by this command.
    /// <see cref="GanttRefreshScheduler"/> populates it on every sheet-change
    /// event; <see cref="RunSilentForExcel"/> consumes and resets it after each
    /// successful refresh. This is what makes change-aware refresh actually
    /// change-aware instead of always running the full pipeline.
    /// </summary>
    private static readonly ChangeTracker _sharedTracker = new();

    /// <summary>
    /// Records one cell change on the shared tracker. Called by
    /// <see cref="GanttRefreshScheduler"/> on every
    /// <c>SheetChange</c> that warrants an auto-refresh. Old value is
    /// <see langword="null"/> because <c>SheetChange</c> fires after the edit;
    /// the new value is read from the cell at call time.
    /// </summary>
    /// <param name="rowNumber">The one-based body-row index.</param>
    /// <param name="column">The 1-based body-column index.</param>
    /// <param name="newValue">The post-edit cell value.</param>
    internal static void RecordChange(int rowNumber, int column, object? newValue) => _sharedTracker.Record(rowNumber, column, null, newValue);

    /// <summary>Runs the production refresh in the current Excel session.</summary>
    internal static void RunForExcel()
    {
        GanttRefreshOrchestrator orchestrator = CreateLiveOrchestrator(ExcelDna.Integration.ExcelDnaUtil.Application);
        GanttRefreshOutcome outcome = orchestrator.Refresh();

        // Store scene warnings in the Ribbon state service so the warnings
        // panel can surface them. Row-level validation issues are reported
        // through row/cell fill colours; scene-level warnings surface here.
        RibbonStateService.Instance.UpdateWarnings(outcome.Warnings);

        if (!outcome.Succeeded)
        {
            Present(CommandErrorDialog.Show, outcome.Message ?? "The chart could not be refreshed.");
        }
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
        // Change-aware auto-refresh: use the shared ChangeTracker owned by
        // GanttRefreshScheduler, which populates it on every SheetChange.
        // The orchestrator targets only affected rows instead of re-reading
        // and re-validating the whole table.
        GanttRefreshOrchestrator orchestrator = CreateLiveOrchestrator(
            ExcelDna.Integration.ExcelDnaUtil.Application, _sharedTracker);
        GanttRefreshOutcome outcome = orchestrator.Refresh();

        // Store scene warnings in the Ribbon state service — same as RunForExcel.
        RibbonStateService.Instance.UpdateWarnings(outcome.Warnings);

        if (outcome.Succeeded)
        {
            // Reset the shared tracker so it only records changes since this
            // successful refresh completed. Until this reset, the tracker still
            // holds the changes that triggered this auto-refresh, so a second
            // immediate SheetChange would re-run the incremental pipeline on
            // the same rows.
            _sharedTracker.Reset();

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
            return;
        }

        Present(RibbonStatusBar.Warn, outcome.Message ?? "The chart could not be refreshed.");
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
    /// <param name="changeTracker">Optional change tracker for change-aware auto-refresh.</param>
    private static GanttRefreshOrchestrator CreateLiveOrchestrator(object? application, ChangeTracker? changeTracker = null) =>
        new(
            new ExcelGanttTableReader(application),
            new ExcelConfigCatalogueReader(application),
            new ExcelWorksheetProtectionGuard(application),
            new ExcelPanelGridMeasurement(application),
            new ExcelDurationWriter(application),
            new ExcelRowHeightNormaliser(application),
            new ExcelOutlineGroupWriter(application),

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
            new ExcelShapeWriter(application, new ExcelWorksheetProtectionGuard(application)),

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

            changeTracker);
#pragma warning restore CA2000
}
