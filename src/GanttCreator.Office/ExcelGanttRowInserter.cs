using System.Reflection;
using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live guarded row inserter for the visible <c>tblGanttData</c> table.
/// Appends one scaffold row and never renders or changes any other workbook
/// state.
/// </summary>
public class ExcelGanttRowInserter(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null,
    ITypeOptionsMaterialiser? typeOptionsMaterialiser = null) : IGanttRowInserter
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);
    private readonly ITypeOptionsMaterialiser _typeOptionsMaterialiser =
        typeOptionsMaterialiser ?? new ExcelTypeOptionsMaterialiser(application);

    /// <inheritdoc />
    public GanttRowInsertOutcome Insert(GanttEntityType type, Func<GanttRowId> nextId)
    {
        // ADR-0008 D4: this is intentionally the first operation in the
        // mutating entry point, before any application/workbook/table access.
        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return GanttRowInsertOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? GanttRowInsertRefusalReason.NoActiveWorkbook
                    : GanttRowInsertRefusalReason.TargetProtected);
        }

        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (workbook is null)
        {
            return GanttRowInsertOutcome.Refused(GanttRowInsertRefusalReason.NoActiveWorkbook);
        }

        Excel.Sheets sheets = workbook.Sheets;
        if (!TryFindTable(sheets, out Excel.ListObject? table) || table is null
            || !TryBuildColumnMap(table, out var columnMap) || columnMap is null)
        {
            return GanttRowInsertOutcome.Refused(GanttRowInsertRefusalReason.TableMissing);
        }

        IReadOnlyList<object?> values = GanttRowDefaults.Build(type, nextId);
        Excel.ListRows rows = GetListRows(table);
        Excel.Range? reusableRow = null;
        Excel.ListRow? newRow = null;
        if (GetListRowCount(rows) == 0)
        {
            reusableRow = GetReusableInitialBlankRow(table);
        }

        Excel.Range rowRange;
        if (reusableRow is not null)
        {
            rowRange = reusableRow;
        }
        else
        {
            newRow = AddRow(rows);
            rowRange = GetRowRange(newRow);
        }

        WriteRow(rowRange, values, columnMap);
        TypeOptionsMaterialiseOutcome typeOptions = _typeOptionsMaterialiser.Materialise();
        if (!typeOptions.Succeeded)
        {
            if (newRow is not null)
            {
                DeleteRow(newRow);
            }
            else
            {
                ClearRange(rowRange);
            }

            return GanttRowInsertOutcome.Refused(GanttRowInsertRefusalReason.TypeOptionsUnavailable);
        }

        return GanttRowInsertOutcome.Ok(newRow is null ? 1 : GetRowIndex(newRow));
    }

    private static bool IsBlankValue(object? value) =>
        value is null
        || value is Missing
        || value is DBNull
        || (value is string text && string.IsNullOrWhiteSpace(text));

    private Excel.Range? GetReusableInitialBlankRow(Excel.ListObject table)
    {
        Excel.Range tableRange = GetTableRange(table);
        Excel.Range tableRows = GetRangeRows(tableRange);
        var rowCount = GetRangeRowCount(tableRows);
        if (rowCount <= 1)
        {
            return null;
        }

        Excel.Range candidate = GetRangeAt(tableRows, rowCount);
        List<object?[]> values = ExcelValue2Matrix.ReadRows(GetRangeValue2(candidate));
        return values.Count == 1
            && values[0].Length > 0
            && values[0].All(IsBlankValue)
            ? candidate
            : null;
    }

    private bool TryFindTable(Excel.Sheets sheets, out Excel.ListObject? table)
    {
        table = null;
        var sheetCount = sheets.Count;
        for (var sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
        {
            if (GetSheetAt(sheets, sheetIndex) is not Excel.Worksheet worksheet)
            {
                continue;
            }

            Excel.ListObjects listObjects = worksheet.ListObjects;
            var tableCount = listObjects.Count;
            for (var tableIndex = 1; tableIndex <= tableCount; tableIndex++)
            {
                Excel.ListObject candidate = GetTableAt(listObjects, tableIndex);
                if (string.Equals(
                    candidate.Name,
                    GanttTableSchema.TableName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    table = candidate;
                    return true;
                }
            }
        }

        return false;
    }

    private bool TryBuildColumnMap(Excel.ListObject table, out int[]? columnMap)
    {
        columnMap = null;
        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        var map = new int[schema.Count];
        var byName = new Dictionary<string, int>(StringComparer.Ordinal);
        Excel.ListColumns columns = table.ListColumns;
        var columnCount = columns.Count;
        for (var index = 1; index <= columnCount; index++)
        {
            Excel.ListColumn column = GetColumnAt(columns, index);
            byName[column.Name] = index;
        }

        for (var schemaIndex = 0; schemaIndex < schema.Count; schemaIndex++)
        {
            GanttTableColumn expected = schema[schemaIndex];
            if (!byName.TryGetValue(expected.Name, out var tableIndex))
            {
                if (expected.IsRequired)
                {
                    return false;
                }

                map[schemaIndex] = -1;
            }
            else
            {
                map[schemaIndex] = tableIndex;
            }
        }

        columnMap = map;
        return true;
    }

    private static void WriteRow(
        Excel.Range rowRange,
        IReadOnlyList<object?> values,
        int[] columnMap)
    {
#pragma warning disable CA1814 // Excel Range.Value2 requires a rectangular SAFEARRAY.
        var matrix = new object[1, columnMap.Length];
        for (var schemaIndex = 0; schemaIndex < values.Count; schemaIndex++)
        {
            var tableIndex = columnMap[schemaIndex];
            if (tableIndex > 0)
            {
                matrix[0, tableIndex - 1] = values[schemaIndex] ?? string.Empty;
            }
        }

        rowRange.Value2 = matrix;
#pragma warning restore CA1814
    }

    internal virtual Excel.Worksheet GetSheetAt(Excel.Sheets sheets, int index) => sheets[index];

    internal virtual Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index) =>
        listObjects[index];

    internal virtual Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index) =>
        columns[index];

    internal virtual Excel.ListRows GetListRows(Excel.ListObject table) => table.ListRows;

    internal virtual int GetListRowCount(Excel.ListRows rows) => rows.Count;

    internal virtual Excel.Range GetTableRange(Excel.ListObject table) => table.Range;

    internal virtual Excel.Range GetRangeRows(Excel.Range range)
    {
        Excel.Range rows = range.Rows;
        return rows;
    }

    internal virtual int GetRangeRowCount(Excel.Range rows)
    {
        var count = rows.Count;
        return count;
    }

    internal virtual Excel.Range GetRangeAt(Excel.Range rows, int index)
    {
        Excel.Range row = rows[index];
        return row;
    }

    internal virtual object? GetRangeValue2(Excel.Range range) => range.Value2;

    internal virtual void ClearRange(Excel.Range range) => range.ClearContents();

    internal virtual Excel.ListRow AddRow(Excel.ListRows rows) => rows.Add(Type.Missing);

    internal virtual int GetRowIndex(Excel.ListRow row) => row.Index;

    internal virtual Excel.Range GetRowRange(Excel.ListRow row) => row.Range;

    internal virtual void DeleteRow(Excel.ListRow row) => row.Delete();
}
