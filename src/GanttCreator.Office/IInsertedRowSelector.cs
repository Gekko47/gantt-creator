namespace GanttCreator.Office;

/// <summary>
/// Moves the user's selection onto a body row of the visible Gantt table, or onto a
/// single cell of one.
/// </summary>
/// <remarks>
/// <para>
/// The surface is deliberately primitive — no interop type crosses the port —
/// so the AddIn compilation never names <c>Microsoft.Office.Interop.Excel</c>
/// (the CS0433 duplicate-type hazard documented on
/// <see cref="IExcelApplicationAdapter"/>).
/// </para>
/// <para>
/// The port lives in the Office project rather than in the AddIn because its
/// implementation is a COM adapter: the Excel proxy chain it walks (workbook,
/// sheets, list objects, list rows, range) is Office-layer work, and a COM
/// implementation in the AddIn would both name the PIA and hold proxies the
/// AddIn does not own. The AddIn keeps only the command.
/// </para>
/// </remarks>
public interface IInsertedRowSelector
{
    /// <summary>Selects the given one-based body row of the visible Gantt table.</summary>
    /// <param name="bodyIndex">The one-based body-row index to select.</param>
    void SelectBodyRow(int bodyIndex);

    /// <summary>
    /// Selects the cell in the user's CURRENT column of the given one-based body row.
    /// </summary>
    /// <param name="bodyIndex">The one-based body-row index whose cell is selected.</param>
    /// <remarks>
    /// <para>
    /// The commands that relocate a row (promote, demote, move up, move down) rewrite
    /// cell contents, so the address the user was on now holds a different activity.
    /// Selecting the row's full width would keep the row but lose the column; this keeps
    /// both, which is what "the selected cell follows the row" means for a user working
    /// in one column.
    /// </para>
    /// <para>
    /// The column comes from the host's active cell at the moment the selection runs.
    /// A row relocation writes cells and never moves the active cell, so that column is
    /// still the user's, and nothing has to be captured before the mutation.
    /// </para>
    /// </remarks>
    void SelectBodyCell(int bodyIndex);
}
