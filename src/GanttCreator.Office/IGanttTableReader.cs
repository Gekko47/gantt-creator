namespace GanttCreator.Office;

/// <summary>
/// Narrow port over the table-read operation: bulk-reads every body row of
/// the visible <c>tblGanttData</c> table into neutral Core DTOs, or refuses
/// with a typed result.
/// </summary>
/// <remarks>
/// <para>
/// The surface is deliberately primitive — no interop type crosses the port —
/// so the AddIn compilation never names
/// <c>Microsoft.Office.Interop.Excel</c> (the CS0433 duplicate-type hazard
/// documented on <see cref="IExcelApplicationAdapter"/>).
/// </para>
/// <para>
/// The port is read-only: it never writes cells, creates sheets, or changes
/// application state. Expected conflicts return typed refusals and mutate
/// nothing; unexpected COM failures propagate to the command boundary.
/// </para>
/// </remarks>
public interface IGanttTableReader
{
    /// <summary>
    /// Reads every body row of <c>tblGanttData</c> in body order.
    /// </summary>
    /// <returns>
    /// The typed outcome: the rows on success, or the refusal reason with an
    /// empty row list. On a refusal nothing was mutated.
    /// </returns>
    GanttTableReadOutcome Read();
}
