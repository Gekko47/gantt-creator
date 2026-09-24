using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// The Initialise-sheet application command: runs the workbook initialiser
/// and surfaces typed refusals as one concise, actionable message. Success is
/// silent (work item R2.2 decision D7): the renamed/populated sheet is
/// self-evident, and the command-boundary contract reserves dialogs for
/// failures. Unexpected exceptions do not cross this command — the Ribbon
/// callback runs it inside the project-wide <see cref="CommandBoundary"/>,
/// which translates them.
/// </summary>
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
        => Run(
            new ExcelWorkbookInitialiser(ExcelDna.Integration.ExcelDnaUtil.Application),
            CommandErrorDialog.Show);

    /// <summary>
    /// Runs one initialise attempt against an injected initialiser and
    /// presenter. A refusal surfaces exactly one presenter message and nothing
    /// is logged; success is silent.
    /// </summary>
    /// <param name="initialiser">The workbook-initialiser port.</param>
    /// <param name="presenter">Receives the translated refusal message.</param>
    /// <exception cref="ArgumentNullException">A dependency is <see langword="null"/>.</exception>
    internal static void Run(IWorkbookInitialiser initialiser, Action<string> presenter)
    {
        ArgumentNullException.ThrowIfNull(initialiser);
        ArgumentNullException.ThrowIfNull(presenter);

        WorkbookInitialiseOutcome outcome = initialiser.Initialise();
        if (outcome.Succeeded)
        {
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
            "The target worksheet already contains a table named tblGanttData, so Initialise sheet left it unchanged.",
        InitialiseRefusalReason.ConfigSheetExists =>
            "This workbook already contains Gantt Creator configuration (_GanttCreatorConfig), "
            + "so Initialise sheet left it unchanged. Use Repair configuration to inspect or repair it.",
        InitialiseRefusalReason.TargetProtected =>
            "The target worksheet is protected, so Initialise sheet cannot populate it. "
            + "Unprotect the worksheet and try again.",
        InitialiseRefusalReason.CatalogueDrift =>
            "Initialise sheet could not verify the Gantt Creator configuration. Repair the configuration and try again.",
        _ => "Initialise sheet could not run. Try again; if it keeps failing, see the Diagnostics dialog.",
    };
#pragma warning restore IDE0060 // Remove unused parameter
}
