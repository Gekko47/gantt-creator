namespace GanttCreator.Office;

/// <summary>
/// Why a validation-report attempt refused to write notes.
/// </summary>
/// <remarks>
/// Every refusal is a typed, expected outcome — routine, not an exception
/// (docs/02-ARCHITECTURE.md "Error handling"). Unexpected COM failures are
/// not refusals and propagate to the command boundary.
/// </remarks>
public enum GanttValidationReportRefusalReason
{
    /// <summary>The Excel application object or the active workbook was absent.</summary>
    NoActiveWorkbook = 0,

    /// <summary>No table named <c>tblGanttData</c> was found, so there is no sheet to annotate.</summary>
    TableMissing = 1,
}

/// <summary>
/// The typed outcome of one validation-report attempt: the number of cell notes
/// written, or the refusal reason. On a refusal nothing was mutated.
/// </summary>
/// <param name="Succeeded">Whether notes were written.</param>
/// <param name="NotesWritten">The count of cell notes written; <c>0</c> on a refusal.</param>
/// <param name="Refusal">
/// The refusal reason when <paramref name="Succeeded"/> is <see langword="false"/>;
/// otherwise <see langword="null"/>.
/// </param>
public sealed record GanttValidationReportOutcome(
    bool Succeeded,
    int NotesWritten,
    GanttValidationReportRefusalReason? Refusal)
{
    /// <summary>
    /// Creates a success outcome.
    /// </summary>
    /// <param name="notesWritten">The count of cell notes written.</param>
    /// <returns>The success outcome.</returns>
    public static GanttValidationReportOutcome Ok(int notesWritten) => new(true, notesWritten, null);

    /// <summary>
    /// Creates a refusal outcome. Nothing was mutated.
    /// </summary>
    /// <param name="refusal">Why no notes were written.</param>
    /// <returns>The refusal outcome.</returns>
    public static GanttValidationReportOutcome Refused(GanttValidationReportRefusalReason refusal) =>
        new(false, 0, refusal);
}
