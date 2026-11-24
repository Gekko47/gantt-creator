using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// The one single-cell write path shared by the label, date-label, and style-key
/// mutators.
/// </summary>
/// <remarks>
/// <para>
/// The adapters keep their <c>internal virtual</c> seams and pass them as method
/// groups, so a contract-test override still fires. The helper owns only the
/// part the three copies shared: the schema column lookup, the one-based row and
/// cell addressing, and the single <see cref="COMException"/> containment.
/// </para>
/// <para>
/// An optional column that the workbook does not carry returns
/// <see cref="CellWriteResult.ColumnMissing"/> rather than aiming the write at a
/// negative index. A required column cannot reach that arm: the table read
/// refuses before this point.
/// </para>
/// </remarks>
internal static class ExcelSingleCellWriter
{
    /// <summary>The closed result of one single-cell write.</summary>
    internal enum CellWriteResult
    {
        /// <summary>The cell was written.</summary>
        Ok,

        /// <summary>The host refused the write.</summary>
        WriteFailed,

        /// <summary>The optional column is absent from this workbook.</summary>
        ColumnMissing,
    }

    /// <summary>
    /// Writes one cell through the caller's seams.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="columnMap">The schema-column to table-column index map.</param>
    /// <param name="rowIndex">The zero-based body-row index.</param>
    /// <param name="columnName">The schema column name.</param>
    /// <param name="value">The value to write.</param>
    /// <param name="columnOptional">Whether the column may be absent.</param>
    /// <param name="getListRows">The adapter's <c>ListObject.ListRows</c> seam.</param>
    /// <param name="getListRowAt">The adapter's <c>ListRows.Item</c> seam.</param>
    /// <param name="getListRowRange">The adapter's <c>ListRow.Range</c> seam.</param>
    /// <param name="getCellAt">The adapter's <c>Range.Cells</c> seam.</param>
    /// <param name="setRangeValues">The adapter's <c>Range.Value2</c> setter seam.</param>
    /// <returns>The typed result.</returns>
    internal static CellWriteResult Write(
        ListObject table,
        int[] columnMap,
        int rowIndex,
        string columnName,
        object? value,
        bool columnOptional,
        Func<ListObject, ListRows> getListRows,
        Func<ListRows, int, ListRow> getListRowAt,
        Func<ListRow, Excel.Range> getListRowRange,
        Func<Excel.Range, int, Excel.Range> getCellAt,
        Action<Excel.Range, object?> setRangeValues)
    {
        ArgumentNullException.ThrowIfNull(table);
        ArgumentNullException.ThrowIfNull(columnMap);
        ArgumentNullException.ThrowIfNull(columnName);
        ArgumentNullException.ThrowIfNull(getListRows);
        ArgumentNullException.ThrowIfNull(getListRowAt);
        ArgumentNullException.ThrowIfNull(getListRowRange);
        ArgumentNullException.ThrowIfNull(getCellAt);
        ArgumentNullException.ThrowIfNull(setRangeValues);

        int tableIndex = columnMap[ExcelGanttRowProjection.SchemaIndexOf(columnName)];
        if (tableIndex < 1)
        {
            return columnOptional ? CellWriteResult.ColumnMissing : CellWriteResult.WriteFailed;
        }

        try
        {
            ListRow row = getListRowAt(getListRows(table), rowIndex + 1);
            Excel.Range cell = getCellAt(getListRowRange(row), tableIndex);
            setRangeValues(cell, value);
        }
        catch (COMException)
        {
            return CellWriteResult.WriteFailed;
        }

        return CellWriteResult.Ok;
    }
}
