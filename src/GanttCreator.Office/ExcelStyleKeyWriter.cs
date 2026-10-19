using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using GanttCreator.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live adapter that writes the engine-owned <c>StyleKey</c> column (column K) in
/// ONE bulk pass over the column range.
/// </summary>
public class ExcelStyleKeyWriter(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null) : IStyleKeyWritePort
{
    private readonly Application? _application = application as Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <inheritdoc />
    public StyleKeyWriteOutcome Write(StyleKeyWritePlan plan)
    {
        ArgumentNullException.ThrowIfNull(plan);

        if (plan.WriteCount == 0)
        {
            return StyleKeyWriteOutcome.Ok(0, 0);
        }

        Application? application = _application;
        Workbook? workbook = application?.ActiveWorkbook;
        if (workbook is null)
        {
            return StyleKeyWriteOutcome.Refused(StyleKeyWriteRefusalReason.NoActiveWorkbook);
        }

        if (!TryFindTable(workbook.Sheets, out Worksheet? worksheet, out ListObject? table)
            || worksheet is null
            || table is null)
        {
            return StyleKeyWriteOutcome.Refused(StyleKeyWriteRefusalReason.TableMissing);
        }

        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(worksheet);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return StyleKeyWriteOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? StyleKeyWriteRefusalReason.NoActiveWorkbook
                    : StyleKeyWriteRefusalReason.TargetProtected);
        }

        if (!TryGetStyleKeyColumnRange(table, out Excel.Range? column) || column is null)
        {
            return StyleKeyWriteOutcome.Refused(StyleKeyWriteRefusalReason.TableMissing);
        }

        object? raw = GetRangeValues(column);
        List<object?[]> current = raw is Array
            ? ExcelValue2Matrix.ReadRows(raw)
            : [RawValueRow(raw)];
        if (current.Count == 0)
        {
            return StyleKeyWriteOutcome.Refused(StyleKeyWriteRefusalReason.WriteFailed);
        }

        object?[] merged = new object?[current.Count];
        for (int index = 0; index < current.Count; index++)
        {
            merged[index] = current[index].Length > 0 ? current[index][0] : null;
        }

        foreach (StyleKeyWrite write in plan.Writes)
        {
            int target = write.RowNumber - 1;
            if (target < 0 || target >= merged.Length)
            {
                return StyleKeyWriteOutcome.Refused(StyleKeyWriteRefusalReason.WriteFailed);
            }

            merged[target] = write.NewStyleKey;
        }

        try
        {
            SetRangeValues(column, ToTwoDimensional(merged));
        }
        catch (COMException)
        {
            return StyleKeyWriteOutcome.Refused(StyleKeyWriteRefusalReason.WriteFailed);
        }

        return StyleKeyWriteOutcome.Ok(plan.WriteCount, plan.WriteCount);
    }


    /// <summary>
    /// Wraps a scalar <c>Value2</c> payload as the one-row, one-cell shape the merge
    /// below reads, so a single-body-row table takes the same path as a many-row one.
    /// </summary>
    /// <param name="value">The scalar cell value Excel returned.</param>
    /// <returns>A single row carrying that one value.</returns>
    private static object?[] RawValueRow(object? value) => [value];

    /// <summary>
    /// Converts the merged column into the rectangular payload Excel's
    /// <c>Value2</c> assignment requires.
    /// </summary>
    /// <param name="values">The merged column values, one per body row.</param>
    /// <returns>A one-based <c>object[,]</c> for assignment.</returns>
    [SuppressMessage(
        "Performance",
        "CA1814:MultiDimArrays",
        Justification = "Excel's Value2 is a SAFEARRAY and is written through a multidimensional array. A jagged array cannot express it, so the analyzer's preferred shape is not available here.")]
    [SuppressMessage(
        "Performance",
        "CA1859:UseConcreteTypesWhenPossible",
        Justification = "The payload is assigned to a COM property typed as object, so the declared return type must stay object.")]
    private static object ToTwoDimensional(object?[] values)
    {
        object[,] payload = new object[values.Length, 1];
        for (int row = 0; row < values.Length; row++)
        {
            payload[row, 0] = values[row]!;
        }

        return payload;
    }

    /// <summary>
    /// Locates the <c>StyleKey</c> column's body range. Test seam over the COM
    /// parameterised properties.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="column">The body range of the <c>StyleKey</c> column.</param>
    /// <returns>Whether the column was found and has a body.</returns>
    internal virtual bool TryGetStyleKeyColumnRange(
        ListObject table,
        out Excel.Range? column)
    {
        ArgumentNullException.ThrowIfNull(table);

        column = null;
        ListColumns? columns = table.ListColumns;
        if (columns is null)
        {
            return false;
        }

        int columnCount = columns.Count;
        for (int index = 1; index <= columnCount; index++)
        {
            ListColumn candidate = GetColumnAt(columns, index);
            if (string.Equals(candidate.Name, "StyleKey", StringComparison.Ordinal))
            {
                column = candidate.DataBodyRange;
                return column is not null;
            }
        }

        return false;
    }

    /// <summary>Reads the range's bulk <c>Value2</c> payload. Test seam.</summary>
    /// <param name="range">The range to read.</param>
    /// <returns>The raw payload.</returns>
    internal virtual object? GetRangeValues(Excel.Range range)
    {
        ArgumentNullException.ThrowIfNull(range);
        return range.Value2;
    }

    /// <summary>Assigns the range's bulk <c>Value2</c> payload. Test seam.</summary>
    /// <param name="range">The range to write.</param>
    /// <param name="payload">The two-dimensional payload.</param>
    internal virtual void SetRangeValues(Excel.Range range, object payload)
    {
        ArgumentNullException.ThrowIfNull(range);
        range.Value2 = payload;
    }

    /// <summary>Finds the Gantt worksheet and its table. Test seam.</summary>
    /// <param name="sheets">The workbook's sheet collection.</param>
    /// <param name="worksheet">The resolved worksheet.</param>
    /// <param name="table">The resolved table.</param>
    /// <returns>Whether both were found.</returns>
    internal virtual bool TryFindTable(
        Sheets sheets,
        out Worksheet? worksheet,
        out ListObject? table)
    {
        worksheet = null;
        table = null;
        foreach (object? entry in sheets)
        {
            if (entry is not Worksheet candidate)
            {
                // A chart sheet or any other non-worksheet entry: it cannot carry the
                // Gantt table, so it is skipped rather than cast.
                continue;
            }

            ListObjects listObjects = candidate.ListObjects;
            foreach (ListObject candidateTable in listObjects)
            {
                if (string.Equals(candidateTable.Name, GanttTableSchema.TableName, StringComparison.OrdinalIgnoreCase))
                {
                    worksheet = candidate;
                    table = candidateTable;
                    return true;
                }
            }
        }

        return false;
    }

    /// <summary>Returns the list column at the one-based index. Test seam.</summary>
    /// <param name="columns">The table's list columns.</param>
    /// <param name="index">The one-based column index.</param>
    /// <returns>The list column.</returns>
    internal virtual ListColumn GetColumnAt(ListColumns columns, int index) => columns[index];
}
