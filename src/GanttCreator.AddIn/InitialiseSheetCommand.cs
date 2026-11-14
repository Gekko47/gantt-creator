using System.Globalization;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// The Initialise-sheet application command: runs the workbook initialiser
/// and surfaces typed refusals as one concise, actionable message. Success is
/// silent in the DIALOG sense (work item R2.2 decision D7): the
/// renamed/populated sheet is self-evident, and the command-boundary contract
/// reserves dialogs for failures. Unexpected exceptions do not cross this
/// command — the Ribbon callback runs it inside the project-wide
/// <see cref="CommandBoundary"/>, which translates them.
/// </summary>
/// <remarks>
/// <para>
/// Every attempt writes ONE durable technical record naming the path it took,
/// the sheet label and any refusal. That is deliberately not a dialog: the
/// live symptom that produced it was "Initialise produces a blank sheet" with
/// nothing in the log to say whether the command ran, was refused, or created
/// a sheet the user was not looking at. Refusals already reach the user
/// through the presenter; the record is what makes the same event reportable
/// from the host afterwards.
/// </para>
/// </remarks>
/// <remarks>
/// <para>
/// docs/02-ARCHITECTURE.md "Ribbon and commands": no callback contains command
/// logic. This class owns the Initialise-sheet logic and never touches the
/// Ribbon; the interop boundary stays behind <see cref="IWorkbookInitialiser"/>
/// (no interop type is named in the AddIn compilation — the CS0433 discipline
/// of <c>IExcelApplicationAdapter</c>).
/// </para>
/// <para>
/// The production entry point resolves the live Excel application through
/// Excel-DNA and shows refusals through the error dialog; both are injectable
/// through the internal overload for contract tests.
/// </para>
/// </remarks>
internal static class InitialiseSheetCommand
{
    /// <summary>
    /// Runs the Initialise-sheet command for the current Excel session: the
    /// production initialiser over <c>ExcelDnaUtil.Application</c>, refusals
    /// shown in the error dialog. Runs on the Excel main STA thread (the
    /// Ribbon <c>onAction</c> thread).
    /// </summary>
    internal static void RunForExcel()
    {
        object application = ExcelDna.Integration.ExcelDnaUtil.Application;

        // Inject the SESSION catalogue writer so its host-rejection evidence
        // log (wired by AddInHost) reaches the writer this command actually
        // runs: the initialiser's default writer has no log, so without this
        // injection the CatalogueWriteRejected record never fires (live
        // observation 2026-10-09T17:04Z: refusal with no record). The
        // fallback keeps the command working when the service was not armed,
        // degrading to the previous log-less writer.
        Run(
            new ExcelWorkbookInitialiser(
                application,
                catalogueWriter: RibbonStateService.Instance.GetCatalogueWriter()
                    ?? new ExcelConfigCatalogueWriter(application),

                // The same sink the command writes its own record to, so the mutation
                // sequence is measured on the production path only. "Initialise and the
                // table creation look slow" cannot be diagnosed from the refresh
                // timings: they start after every one of these steps has finished.
                trace: CommandBoundary.Instance.WriteTechnicalRecord),
            CommandErrorDialog.Show,
            RibbonStateService.Instance.ScheduleAutomaticRefresh,
            CommandBoundary.Instance.WriteTechnicalRecord);
    }

    /// <summary>
    /// Runs one initialise attempt against an injected initialiser and
    /// presenter. A refusal surfaces exactly one presenter message; success is
    /// silent in the dialog sense. Both outcomes write one technical record
    /// through <paramref name="trace"/> when it is supplied.
    /// </summary>
    /// <param name="initialiser">The workbook-initialiser port.</param>
    /// <param name="presenter">Receives the translated refusal message.</param>
    /// <param name="onInitialised">Schedules the one-time initial empty render after creation succeeds.</param>
    /// <param name="trace">
    /// Durable technical-record sink. Written BEFORE the presenter so a modal
    /// dialog cannot hide the evidence, and never through the presenter, which
    /// is reserved for the user-facing refusal message.
    /// </param>
    /// <exception cref="ArgumentNullException">A dependency is <see langword="null"/>.</exception>
    internal static void Run(
        IWorkbookInitialiser initialiser,
        Action<string> presenter,
        Action? onInitialised = null,
        Action<string>? trace = null)
    {
        ArgumentNullException.ThrowIfNull(initialiser);
        ArgumentNullException.ThrowIfNull(presenter);

        WorkbookInitialiseOutcome outcome = initialiser.Initialise();
        trace?.Invoke(Describe(outcome));
        if (outcome.Succeeded)
        {
            onInitialised?.Invoke();
            return;
        }

        // CA1031: the presenter failure is outside this command's contract —
        // the boundary pattern degrades to no message rather than propagating
        // into Excel.
#pragma warning disable CA1031
        try
        {
            presenter(TranslateRefusal(outcome.Refusal!.Value, outcome.SheetName));
        }
        catch
        {
            // Intentionally empty: a dialog failure degrades to no dialog.
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Formats the durable record for one attempt: the path taken, the final
    /// sheet label, and the refusal reason when there is one.
    /// </summary>
    /// <param name="outcome">The initialiser's outcome.</param>
    /// <returns>One invariant-culture technical record.</returns>
    private static string Describe(WorkbookInitialiseOutcome outcome) => string.Create(
        CultureInfo.InvariantCulture,
        $"gantt.initialise path={outcome.Path} sheet={outcome.SheetName ?? "none"} refusal={outcome.Refusal?.ToString() ?? "none"}");

    /// <summary>
    /// Translates a refusal to the user-safe message. The messages are
    /// actionable and name the situation the user can act on; they never
    /// contain schedule content or technical detail.
    /// </summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <param name="_sheetName">
    /// Reserved for future use. Not interpolated in the current messages.
    /// </param>
    /// <returns>The user-safe message.</returns>
#pragma warning disable IDE0060 // Remove unused parameter
    internal static string TranslateRefusal(InitialiseRefusalReason refusal, string? _sheetName) => refusal switch
    {
        InitialiseRefusalReason.NoActiveWorkbook =>
            "Gantt Creator needs an open workbook. Open or create a workbook, then try Initialise sheet again.",
        InitialiseRefusalReason.TableExists =>
            "A table named tblGanttData already exists somewhere in this workbook, "
            + "so Initialise sheet left it unchanged.",
        InitialiseRefusalReason.ConfigSheetExists =>
            "This workbook already contains Gantt Creator configuration (_GanttCreatorConfig), "
            + "so Initialise sheet left it unchanged. Use Repair configuration to inspect or repair it.",
        InitialiseRefusalReason.TargetProtected =>
            "The target worksheet or workbook structure is protected, so Initialise sheet cannot populate it. "
            + "Check both target worksheet protection and workbook structure protection, then try again.",
        InitialiseRefusalReason.CatalogueDrift =>
            "Initialise sheet could not verify the Gantt Creator configuration. Repair the configuration and try again.",
        InitialiseRefusalReason.CatalogueWriteFailed =>
            "Excel refused to write the Gantt Creator configuration, so nothing was changed. "
            + "Close any open dialogs, check that Excel is not busy, then try again.",
        _ => "Initialise sheet could not run. Try again; if it keeps failing, see the Diagnostics dialog.",
    };
#pragma warning restore IDE0060 // Remove unused parameter
}
