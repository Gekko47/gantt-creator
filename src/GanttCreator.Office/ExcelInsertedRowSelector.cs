using GanttCreator.Core;
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
    private readonly Excel.Workbook? _workbook = CaptureWorkbook(application);

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
    private static Excel.Workbook? CaptureWorkbook(object? application)
    {
        var excel = application as Excel.Application;

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
        Excel.Workbook? workbook = _workbook;
        if (workbook is null)
        {
            return;
        }

        Excel.Sheets sheets = workbook.Sheets;
        var count = sheets.Count;
        for (var sheetIndex = 1; sheetIndex <= count; sheetIndex++)
        {
            var sheet = GetSheetAt(sheets, sheetIndex);
            if (sheet is not Excel.Worksheet worksheet)
            {
                continue;
            }

            Excel.ListObjects listObjects = worksheet.ListObjects;
            var tableCount = listObjects.Count;
            for (var tableIndex = 1; tableIndex <= tableCount; tableIndex++)
            {
                Excel.ListObject table = GetTableAt(listObjects, tableIndex);
                if (!string.Equals(table.Name, GanttTableSchema.TableName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                // ListRows is table-relative, so body row N is ListRows[N]
                // whatever worksheet row the table starts on. Indexing
                // table.Range.Rows by an absolute worksheet row only agrees with
                // the body index when the table happens to start at row 1.
                Excel.ListRow targetRow = GetListRowAt(GetListRows(table), bodyIndex);

                // Range.Select only works on the active sheet, so the Gantt
                // worksheet is activated first. A refused activation or selection
                // degrades to no selection change: the row has already been added
                // by the time this runs, so letting a host failure escape would
                // report the insert as failed when it in fact succeeded.
                //
                // CA1031: both calls are best-effort selection changes whose only
                // failure mode is "the cursor stays put". That must not propagate
                // into Excel and undo the command the user asked for.
#pragma warning disable CA1031
                try
                {
                    worksheet.Activate();
                    Excel.Range target = GetRowRange(targetRow);
                    target.Select();
                }
                catch
                {
                    // Intentionally empty: the row is added either way; only the
                    // cursor stays put.
                }
#pragma warning restore CA1031
                return;
            }
        }
    }
    /// <summary>Returns the sheet object at the one-based index.</summary>
    /// <param name="sheets">The workbook sheet collection.</param>
    /// <param name="index">The one-based sheet index.</param>
    /// <returns>The sheet object.</returns>
    internal virtual object GetSheetAt(Excel.Sheets sheets, int index) => sheets[index];

    /// <summary>Returns the list object at the one-based index.</summary>
    /// <param name="listObjects">The worksheet list-object collection.</param>
    /// <param name="index">The one-based list-object index.</param>
    /// <returns>The list object.</returns>
    internal virtual Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index) => listObjects[index];

    /// <summary>Returns the table's list rows.</summary>
    /// <param name="table">The Gantt table.</param>
    /// <returns>The list-row collection.</returns>
    internal virtual Excel.ListRows GetListRows(Excel.ListObject table) => table.ListRows;

    /// <summary>Returns the one-based body row of the table.</summary>
    /// <param name="rows">The table's list rows.</param>
    /// <param name="bodyIndex">The one-based body-row index.</param>
    /// <returns>The list row.</returns>
    internal virtual Excel.ListRow GetListRowAt(Excel.ListRows rows, int bodyIndex) => rows[bodyIndex];

    /// <summary>Returns the range spanning a list row.</summary>
    /// <param name="row">The list row.</param>
    /// <returns>The row's range.</returns>
    internal virtual Excel.Range GetRowRange(Excel.ListRow row) => row.Range;
}
