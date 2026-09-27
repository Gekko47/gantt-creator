namespace GanttCreator.Office;

/// <summary>
/// Moves the user's selection onto a body row of the visible Gantt table.
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
}
