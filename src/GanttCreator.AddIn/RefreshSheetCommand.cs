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
    /// <summary>Runs the production refresh in the current Excel session.</summary>
    internal static void RunForExcel() =>
        Run(
            CreateLiveOrchestrator(ExcelDna.Integration.ExcelDnaUtil.Application),
            CommandErrorDialog.Show);

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

    /// <summary>
    /// Composes the live orchestrator from the Excel application object.
    /// </summary>
    /// <param name="application">The Excel application object.</param>
    /// <returns>The composed orchestrator.</returns>
    /// <remarks>
    /// Every dependency is constructed here and nowhere else, so the composition root
    /// is one readable list rather than something assembled across the command, the
    /// Ribbon, and a static initialiser. Each adapter is the one already landed for
    /// its own row; this method invents no behaviour.
    /// </remarks>
    private static GanttRefreshOrchestrator CreateLiveOrchestrator(object? application) =>
        new(
            new ExcelGanttTableReader(application),
            new ExcelConfigCatalogueReader(application),
            new ExcelWorksheetProtectionGuard(application),
            new ExcelPanelGridMeasurement(application),
            new ExcelDurationWriter(application),
            new ExcelRowHeightNormaliser(application),
            new ExcelOutlineGroupWriter(application),
            new ExcelSceneBuildRequestFactory(),
            new ExcelShapeWriter(application, new ExcelWorksheetProtectionGuard(application)));
}
