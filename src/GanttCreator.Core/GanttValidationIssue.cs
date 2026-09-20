namespace GanttCreator.Core;

/// <summary>
/// One actionable validation finding for a single <c>tblGanttData</c> body row.
/// Blocking <see cref="GanttValidationSeverity.Error"/>s exclude the row from
/// the validated events; <see cref="GanttValidationSeverity.Warning"/>s permit
/// rendering. Immutable.
/// </summary>
/// <param name="RowNumber">The one-based body-row index from <see cref="GanttRowDto.RowNumber"/>.</param>
/// <param name="Field">The workbook column the finding concerns (e.g. <c>Start</c>, <c>Type</c>).</param>
/// <param name="Code">The stable machine-readable code (see <see cref="GanttValidationCodes"/>).</param>
/// <param name="Severity">Whether the finding blocks the row.</param>
/// <param name="Message">The concise human-readable explanation.</param>
public sealed record GanttValidationIssue(
    int RowNumber,
    string Field,
    string Code,
    GanttValidationSeverity Severity,
    string Message);
