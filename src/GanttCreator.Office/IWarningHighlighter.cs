namespace GanttCreator.Office;

/// <summary>
/// Highlights validation-warning cells on the visible <c>tblGanttData</c> body with
/// <see cref="Core.GanttCatalogues.WarningFill"/>, and clears the fill
/// from cells that no longer warn. Never mutates cell values; never touches the
/// VeryHidden configuration sheet.
/// </summary>
public interface IWarningHighlighter
{
    /// <summary>
    /// Applies <see cref="Core.GanttCatalogues.WarningFill"/> to cells that
    /// carry validation warnings, and clears it from cells that no longer do.
    /// </summary>
    /// <param name="issues">
    /// Validation issues in <c>Severity → Row → Field → Code</c> order. Only
    /// <see cref="Core.GanttValidationSeverity.Warning"/> issues are
    /// highlighted; errors are left to the caller's refusal path.
    /// </param>
    /// <returns>
    /// The count of cells highlighted, or a typed refusal. On a refusal nothing
    /// was mutated.
    /// </returns>
    GanttValidationReportOutcome HighlightWarnings(IReadOnlyList<Core.GanttValidationIssue> issues);
}
