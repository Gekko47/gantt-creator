using GanttCreator.Core;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// The Validate-sheet application command: reads the visible <c>tblGanttData</c>
/// table, validates every row, and surfaces each finding as a cell note on its
/// offending cell — without mutating any cell value or format. The R2.4 read and
/// R2.5 validate pipeline is the first consumer of the row-level error report
/// (work item R2.6).
/// </summary>
/// <remarks>
/// <para>
/// docs/02-ARCHITECTURE.md "Ribbon and commands": no callback contains command
/// logic. This class owns the Validate logic and never touches the Ribbon; the
/// interop boundary stays behind <see cref="IGanttTableReader"/> and
/// <see cref="IGanttValidationReporter"/> (no interop type is named in the
/// AddIn compilation — the CS0433 discipline of
/// <see cref="IExcelApplicationAdapter"/>).
/// </para>
/// <para>
/// The production entry point resolves the live Excel application through
/// Excel-DNA and uses <see cref="CommandErrorDialog.Show"/> for the one summary
/// dialog; both are injectable through the internal overload for contract
/// tests.
/// </para>
/// </remarks>
internal static class ValidateSheetCommand
{
    /// <summary>
    /// Runs the Validate command for the current Excel session: the production
    /// reader and reporter over <c>ExcelDnaUtil.Application</c>, and the typed
    /// refusals shown in the error dialog. Runs on the Excel main STA thread (the
    /// Ribbon <c>onAction</c> thread).
    /// </summary>
    internal static void RunForExcel()
        => Run(
            new ExcelGanttTableReader(ExcelDna.Integration.ExcelDnaUtil.Application),
            new ExcelGanttValidationReporter(
                ExcelDna.Integration.ExcelDnaUtil.Application,
                new ExcelWorksheetProtectionGuard(ExcelDna.Integration.ExcelDnaUtil.Application)),
            CommandErrorDialog.Show);

    /// <summary>
    /// Runs one validate pass against injected reader/reporter ports and a
    /// presenter. A read refusal surfaces exactly one actionable message; a
    /// successful pass surfaces exactly one summary message; a failed presenter
    /// degrades to no dialog (CA1031-guarded, identical to
    /// <see cref="InitialiseSheetCommand"/>).
    /// </summary>
    /// <param name="reader">The table-read port.</param>
    /// <param name="reporter">The validation-report port.</param>
    /// <param name="presenter">Receives the translated message.</param>
    /// <exception cref="ArgumentNullException">A dependency is <see lang="null"/>.</exception>
    internal static void Run(
        IGanttTableReader reader,
        IGanttValidationReporter reporter,
        Action<string> presenter)
    {
        ArgumentNullException.ThrowIfNull(reader);
        ArgumentNullException.ThrowIfNull(reporter);
        ArgumentNullException.ThrowIfNull(presenter);

        GanttTableReadOutcome read = reader.Read();
        if (!read.Succeeded)
        {
            SafePresent(presenter, TranslateReadRefusal(read.Refusal!.Value));
            return;
        }

        GanttValidationOutcome validation = GanttRowValidator.Validate(read.Rows);
        GanttValidationReportOutcome report = reporter.Report(validation.Issues);
        if (!report.Succeeded)
        {
            SafePresent(presenter, TranslateReportRefusal(report.Refusal!.Value));
            return;
        }

        SafePresent(presenter, SummaryMessage(validation.Issues, report.NotesWritten));
    }

    /// <summary>
    /// Invokes the presenter inside the command's CA1031 boundary. A dialog
    /// failure degrades to no dialog rather than propagating into the Ribbon
    /// callback; the underlying validation/report work is unaffected.
    /// </summary>
    /// <param name="presenter">The presenter delegate.</param>
    /// <param name="message">The message to present.</param>
    private static void SafePresent(Action<string> presenter, string message)
    {
#pragma warning disable CA1031
        try
        {
            presenter(message);
        }
        catch
        {
            // Intentionally empty: a dialog failure degrades to no dialog.
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Translates a table-read refusal to a user-safe message. The message is
    /// actionable and names the situation the user can act on; it never contains
    /// schedule content or technical detail.
    /// </summary>
    /// <param name="refusal">The read refusal reason.</param>
    /// <returns>The user-safe message.</returns>
    internal static string TranslateReadRefusal(GanttTableReadRefusalReason refusal) => refusal switch
    {
        GanttTableReadRefusalReason.NoActiveWorkbook =>
            "Gantt Creator needs an open workbook. Open a workbook, then run Validate again.",
        GanttTableReadRefusalReason.TableMissing =>
            "Validate could not find a table named tblGanttData on the active sheet. Run Initialise sheet first.",
        GanttTableReadRefusalReason.DateSystemUnsupported =>
            "This workbook uses the 1904 date system, which Gantt Creator does not support. Use a 1900-date workbook and run Validate again.",
        _ => "Validate could not run. Try again; if it keeps failing, see the Diagnostics dialog.",
    };

    /// <summary>
    /// Translates a report-write refusal to a user-safe message.
    /// </summary>
    /// <param name="refusal">The report refusal reason.</param>
    /// <returns>The user-safe message.</returns>
    internal static string TranslateReportRefusal(GanttValidationReportRefusalReason refusal) => refusal switch
    {
        GanttValidationReportRefusalReason.NoActiveWorkbook =>
            "Gantt Creator needs an active workbook to write validation notes. Open a workbook and try again.",
        GanttValidationReportRefusalReason.TableMissing =>
            "Validate could not locate tblGanttData to write onto. Run Initialise sheet first.",
        GanttValidationReportRefusalReason.TargetProtected =>
            "The worksheet or workbook is protected, so validation notes were not changed. Remove protection and run Validate again.",
        _ => "Validate could not write the notes. Try again; if it keeps failing, see the Diagnostics dialog.",
    };

    /// <summary>
    /// Builds the single summary message shown after a successful validation +
    /// report pass: counts only, never schedule content.
    /// </summary>
    /// <param name="issues">The issues reported (for counts).</param>
    /// <param name="notesWritten">The number of cell notes written.</param>
    /// <returns>The summary message.</returns>
    internal static string SummaryMessage(IReadOnlyList<GanttValidationIssue> issues, int notesWritten)
    {
        var errors = 0;
        var warnings = 0;
        foreach (GanttValidationIssue issue in issues)
        {
            if (issue.Severity == GanttValidationSeverity.Error)
            {
                errors++;
            }
            else
            {
                warnings++;
            }
        }

        return errors == 0 && warnings == 0
            ? "Validation complete: all rows valid."
            : string.Format(
                System.Globalization.CultureInfo.CurrentCulture,
                "Validation complete: {0} error(s) and {1} warning(s) on tblGanttData; {2} note(s) written.",
                errors,
                warnings,
                notesWritten);
    }
}
