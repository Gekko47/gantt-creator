using GanttCreator.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="IInsertedRowSelector"/> over the Excel application object.
/// </summary>
/// <param name="application">
/// The Excel application object (for example <c>ExcelDnaUtil.Application</c>), or
/// <see langword="null"/> when the host supplied none.
/// </param>
/// <remarks>
/// <para>
/// The target workbook is captured in the constructor, not resolved when the
/// selection runs. The selector is built before <c>AddRowCommand.Run</c>
/// performs the write, so the capture records the workbook that is active
/// before any handler runs; re-reading <c>Application.ActiveWorkbook</c> after
/// the write would follow whatever a cell-change or workbook-activate handler
/// left active, which may be a different workbook from the one that received
/// the row.
/// </para>
/// <para>
/// The <c>internal virtual</c> accessors isolate the Excel COM parameterised
/// properties (indexers). Expression trees cannot contain indexed properties
/// (CS0855), so contract tests substitute these seams and every other member
/// through Moq; the real indexer and <c>Select</c>/<c>Activate</c> behaviour is
/// exercised by the tagged live-Office integration test.
/// </para>
/// <para>
/// COM ownership: every proxy walked here is an Excel-owned shared root reached
/// from the captured workbook. This adapter owns none of them, never calls
/// <c>FinalReleaseComObject</c>, and force-releases nothing, and each is held in
/// a local and used without chained member expressions.
/// </para>
/// <para>
/// This adapter writes nothing: it activates the Gantt worksheet and selects a
/// range. It is therefore classified read-only by
/// <c>ProtectionGuardFirstTests</c> — it mutates no cell, table, sheet, or name.
/// </para>
/// </remarks>
public class ExcelInsertedRowSelector(object? application) : IInsertedRowSelector
{
    private readonly Workbook? _workbook = CaptureWorkbook(application);

    /// <summary>
    /// The application object, held only to read the ACTIVE CELL's column when a cell is
    /// selected. Nothing else here reads the application: the workbook is captured so a
    /// handler that activates another workbook cannot redirect the selection.
    /// </summary>
    private readonly Application? _application = application as Application;

    /// <summary>
    /// Captures the currently active workbook, degrading to "no selection change"
    /// when the host refuses.
    /// </summary>
    /// <param name="application">The Excel application object, or <see langword="null"/>.</param>
    /// <returns>The captured workbook, or <see langword="null"/>.</returns>
    /// <remarks>
    /// CA1031: this read happens before the row exists, so a host failure must
    /// degrade to "no selection change" rather than escape the command and
    /// prevent the insert the user asked for.
    /// </remarks>
    private static Workbook? CaptureWorkbook(object? application)
    {
        Application? excel = application as Application;

        // CA1031: a best-effort read whose only failure mode is "the cursor stays
        // put" must not propagate into Excel.
#pragma warning disable CA1031
        try
        {
            return excel?.ActiveWorkbook;
        }
        catch
        {
            // Intentionally empty: the row is still added; only the cursor stays put.
            return null;
        }
#pragma warning restore CA1031
    }

    /// <inheritdoc />
    public void SelectBodyRow(int bodyIndex)
    {
        Workbook? workbook = _workbook;
        if (workbook is null)
        {
            return;
        }

        // Every COM call this method makes is best-effort. The row has already been
        // added by the time this runs, so a host failure anywhere in the walk must
        // degrade to "the cursor stays put" rather than escape: letting one out
        // would report the insert as failed to the user when it in fact succeeded.
        // The walk itself is inside the guard for the same reason as the two
        // selection calls below - a refused sheet, table, or list-row lookup is
        // still a failure to select, not a failure of the insert.
        //
        // CA1031: these are all best-effort selection changes whose only failure
        // mode is "the cursor stays put". That must not propagate into Excel and
        // undo the command the user asked for.
#pragma warning disable CA1031
        try
        {
            SelectRowCore(workbook, bodyIndex, worksheetColumn: null);
        }
        catch
        {
            // Intentionally empty: the row is added either way; only the cursor
            // stays put.
        }
#pragma warning restore CA1031
    }

    /// <inheritdoc />
    public void SelectBodyCell(int bodyIndex)
    {
        Workbook? workbook = _workbook;
        if (workbook is null)
        {
            return;
        }

        // Best-effort for the same reason as SelectBodyRow: the row has already moved,
        // and a refused activation or selection must degrade to "the cursor stays
        // where the host left it" rather than surface a failed command that in fact
        // succeeded. CA1031: these are selection changes whose only failure mode is
        // the cursor not moving.
#pragma warning disable CA1031
        try
        {
            SelectRowCore(workbook, bodyIndex, ReadActiveColumn());
        }
        catch
        {
            // Intentionally empty: the row moved either way; only the cursor stays put.
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Walks the workbook to the Gantt table and selects one body row, or one cell of it.
    /// </summary>
    /// <param name="workbook">The workbook captured before the insert or move.</param>
    /// <param name="bodyIndex">The one-based body-row index to select.</param>
    /// <param name="worksheetColumn">
    /// The worksheet column whose cell is selected, or <see langword="null"/> to select
    /// the row's full width.
    /// </param>
    /// <remarks>
    /// Split out of the public members so a single best-effort guard covers the sheet,
    /// table, and row lookups as well as <c>Activate</c> and <c>Select</c>. Every proxy
    /// is held in a local and used without chained member expressions.
    /// </remarks>
    private void SelectRowCore(Workbook workbook, int bodyIndex, int? worksheetColumn)
    {
        Sheets sheets = workbook.Sheets;
        int count = sheets.Count;
        for (int sheetIndex = 1; sheetIndex <= count; sheetIndex++)
        {
            object sheet = GetSheetAt(sheets, sheetIndex);
            if (sheet is not Worksheet worksheet)
            {
                continue;
            }

            ListObjects listObjects = worksheet.ListObjects;
            int tableCount = listObjects.Count;
            for (int tableIndex = 1; tableIndex <= tableCount; tableIndex++)
            {
                ListObject table = GetTableAt(listObjects, tableIndex);
                if (!GanttChartContract.IsDataTableName(table.Name))
                {
                    continue;
                }

                // ListRows is table-relative, so body row N is ListRows[N]
                // whatever worksheet row the table starts on. Indexing
                // table.Range.Rows by an absolute worksheet row only agrees with
                // the body index when the table happens to start at row 1.
                //
                // The ListRows proxy is held in a local rather than passed inline
                // so the ownership rule holds: every proxy is named once and used
                // through that name. It is Excel-owned and shared, reached from the
                // captured workbook, so it is not force-released - see the class
                // remarks.
                ListRows listRows = GetListRows(table);
                ListRow targetRow = GetListRowAt(listRows, bodyIndex);

                Excel.Range rowRange = GetRowRange(targetRow);

                // A supplied column selects ONE CELL of the row rather than its full
                // width. The address is built from the worksheet, not an offset into the
                // row range, so a column outside the table still selects the cell the
                // user was working in -- which is where they expect the cursor to be
                // after the row moved.
                Excel.Range target = worksheetColumn is { } column
                    ? GetCell(worksheet, GetRangeRow(rowRange), column)
                    : rowRange;

                // Range.Select only works on the active sheet, so the Gantt
                // worksheet is activated first. A refused activation or selection
                // degrades to no selection change for the reason the enclosing
                // guard records.
                worksheet.Activate();
                target.Select();
                return;
            }
        }
    }

    /// <summary>Returns the sheet object at the one-based index.</summary>
    /// <param name="sheets">The workbook sheet collection.</param>
    /// <param name="index">The one-based sheet index.</param>
    /// <returns>The sheet object.</returns>
    internal virtual object GetSheetAt(Sheets sheets, int index) => sheets[index];

    /// <summary>Returns the list object at the one-based index.</summary>
    /// <param name="listObjects">The worksheet list-object collection.</param>
    /// <param name="index">The one-based list-object index.</param>
    /// <returns>The list object.</returns>
    internal virtual ListObject GetTableAt(ListObjects listObjects, int index) => listObjects[index];

    /// <summary>Returns the table's list rows.</summary>
    /// <param name="table">The Gantt table.</param>
    /// <returns>The list-row collection.</returns>
    internal virtual ListRows GetListRows(ListObject table) => table.ListRows;

    /// <summary>Returns the one-based body row of the table.</summary>
    /// <param name="rows">The table's list rows.</param>
    /// <param name="bodyIndex">The one-based body-row index.</param>
    /// <returns>The list row.</returns>
    internal virtual ListRow GetListRowAt(ListRows rows, int bodyIndex) => rows[bodyIndex];

    /// <summary>Returns the range spanning a list row.</summary>
    /// <param name="row">The list row.</param>
    /// <returns>The row's range.</returns>
    internal virtual Excel.Range GetRowRange(ListRow row) => row.Range;

    /// <summary>Returns the worksheet row of a range. Test seam over the COM property.</summary>
    /// <param name="range">The range.</param>
    /// <returns>The one-based worksheet row.</returns>
    internal virtual int GetRangeRow(Excel.Range range) => range.Row;

    /// <summary>Returns one cell of a worksheet. Test seam over the COM indexer.</summary>
    /// <param name="worksheet">The worksheet.</param>
    /// <param name="worksheetRow">The one-based worksheet row.</param>
    /// <param name="worksheetColumn">The one-based worksheet column.</param>
    /// <returns>The cell's range.</returns>
    internal virtual Excel.Range GetCell(Worksheet worksheet, int worksheetRow, int worksheetColumn)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        Excel.Range cells = worksheet.Cells;
        return cells[worksheetRow, worksheetColumn];
    }

    /// <summary>
    /// Reads the active cell's worksheet column, or <see langword="null"/> when the host
    /// reports no active cell.
    /// </summary>
    /// <returns>The one-based column, or <see langword="null"/>.</returns>
    /// <remarks>
    /// Read at selection time, not captured before the mutation: a row relocation writes
    /// cells and never moves the active cell, so this value is still the user's column.
    /// A host that reports nothing falls back to selecting the row's full width, which is
    /// the behaviour this port had before cell selection existed.
    /// </remarks>
    internal virtual int? ReadActiveColumn()
    {
        Application? application = _application;
        Excel.Range? activeCell = application?.ActiveCell;
        return activeCell?.Column;
    }
}
