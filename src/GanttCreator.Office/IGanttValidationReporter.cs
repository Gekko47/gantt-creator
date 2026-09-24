namespace GanttCreator.Office;

/// <summary>
/// Narrow port over the R2.6 row-level error-reporting operation: writes
/// add-in-owned cell notes for every <see cref="Core.GanttValidationIssue"/> in the
/// supplied outcome, or refuses with a typed result.
/// </summary>
/// <remarks>
/// <para>
/// The surface is deliberately primitive — no interop type crosses the port —
/// so the AddIn compilation never names
/// <c>Microsoft.Office.Interop.Excel</c> (the CS0433 duplicate-type hazard
/// documented on <see cref="IExcelApplicationAdapter"/>).
/// </para>
/// <para>
/// The port is strictly additive over cells: it never writes or formats a cell
/// value, mutates only add-in-owned cell notes (tagged by a text sentinel), and
/// clears only notes it owns. Every expected conflict returns a typed refusal
/// and leaves the sheet exactly as it was.
/// </para>
/// </remarks>
public interface IGanttValidationReporter
{
    /// <summary>
    /// Annotates the offending cells of <c>tblGanttData</c> with cell notes for
    /// each issue, clearing prior owned notes first so the report is idempotent.
    /// </summary>
    /// <param name="issues">
    /// The validation findings to report. Must already be in the deterministic
    /// <c>Severity → Row → Field → Code</c> order produced by
    /// <see cref="Core.GanttRowValidator.Validate(IReadOnlyList{Core.GanttRowDto})"/>.
    /// </param>
    /// <returns>
    /// The typed outcome: the count of notes written on success, or the refusal
    /// reason with <c>0</c> written. On a refusal nothing was mutated.
    /// </returns>
    GanttValidationReportOutcome Report(IReadOnlyList<Core.GanttValidationIssue> issues);
}
