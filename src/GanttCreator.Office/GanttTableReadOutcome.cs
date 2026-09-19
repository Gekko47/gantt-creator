namespace GanttCreator.Office;

/// <summary>
/// Why a table-read attempt refused to return rows.
/// </summary>
/// <remarks>
/// Every refusal is a typed, expected outcome — routine, not an exception
/// (docs/02-ARCHITECTURE.md "Error handling"). Unexpected COM failures are
/// not refusals and propagate to the command boundary.
/// </remarks>
public enum GanttTableReadRefusalReason
{
    /// <summary>The Excel application object or the active workbook was absent.</summary>
    NoActiveWorkbook = 0,

    /// <summary>No table named <c>tblGanttData</c> with the required headers was found.</summary>
    TableMissing = 1,

    /// <summary>The workbook uses the 1904 date system, which R2.4 explicitly rejects (see ADR-0006).</summary>
    DateSystemUnsupported = 2,
}

/// <summary>
/// The typed outcome of one table-read attempt: the rows read, or the
/// refusal reason. On a refusal nothing was mutated.
/// </summary>
/// <param name="Succeeded">Whether rows were returned.</param>
/// <param name="Rows">The rows in body order; empty on a refusal.</param>
/// <param name="Refusal">The refusal reason when <paramref name="Succeeded"/> is <see langword="false"/>; otherwise <see langword="null"/>.</param>
public sealed record GanttTableReadOutcome(
    bool Succeeded,
    IReadOnlyList<Core.GanttRowDto> Rows,
    GanttTableReadRefusalReason? Refusal)
{
    /// <summary>
    /// Creates a success outcome.
    /// </summary>
    /// <param name="rows">The rows in body order.</param>
    /// <returns>The success outcome.</returns>
    public static GanttTableReadOutcome Ok(IReadOnlyList<Core.GanttRowDto> rows) => new(true, rows, null);

    /// <summary>
    /// Creates a refusal outcome. Nothing was mutated.
    /// </summary>
    /// <param name="refusal">Why no rows were returned.</param>
    /// <returns>The refusal outcome.</returns>
    public static GanttTableReadOutcome Refused(GanttTableReadRefusalReason refusal) =>
        new(false, [], refusal);
}
