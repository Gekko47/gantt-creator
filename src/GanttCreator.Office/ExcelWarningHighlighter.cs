using GanttCreator.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="IWarningHighlighter"/> over the Excel application object.
/// Highlights validation-warning cells on the visible <c>tblGanttData</c> body with
/// <see cref="GanttCatalogues.WarningFill"/> and clears it from cells that no longer warn.
/// </summary>
/// <param name="application">
/// The Excel application object, or <see langword="null"/> when the host supplied none.
/// </param>
/// <param name="protectionGuard">The shared read-only workbook-protection guard.</param>
public class ExcelWarningHighlighter(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null) : IWarningHighlighter
{
    private readonly Application? _application = application as Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <summary>
    /// Applies <see cref="GanttCatalogues.WarningFill"/> to cells that carry
    /// validation warnings, and clears it from cells that no longer do.
    /// </summary>
    /// <param name="issues">
    /// Validation issues in <c>Severity → Row → Field → Code</c> order. Only
    /// <see cref="GanttValidationSeverity.Warning"/> issues are highlighted;
    /// errors are left to the caller's refusal path.
    /// </param>
    /// <returns>
    /// The count of cells highlighted, or a typed refusal. On a refusal nothing
    /// was mutated.
    /// </returns>
    public GanttValidationReportOutcome HighlightWarnings(
        IReadOnlyList<GanttValidationIssue> issues)
    {
        Application? application = _application;
        if (application is null)
        {
            return GanttValidationReportOutcome.Refused(
                GanttValidationReportRefusalReason.NoActiveWorkbook);
        }

        ProtectionGuardOutcome activeProtection = _protectionGuard.Query();
        if (activeProtection != ProtectionGuardOutcome.NotProtected)
        {
            return GanttValidationReportOutcome.Refused(
                activeProtection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? GanttValidationReportRefusalReason.NoActiveWorkbook
                    : GanttValidationReportRefusalReason.TargetProtected);
        }

        Workbook? workbook = application.ActiveWorkbook;
        if (workbook is null)
        {
            return GanttValidationReportOutcome.Refused(
                GanttValidationReportRefusalReason.NoActiveWorkbook);
        }

        Sheets sheets = workbook.Sheets;
        if (!TryFindTable(sheets, out Worksheet? worksheet, out ListObject? table) ||
            worksheet is null || table is null)
        {
            return GanttValidationReportOutcome.Refused(
                GanttValidationReportRefusalReason.TableMissing);
        }

        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(worksheet);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return GanttValidationReportOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? GanttValidationReportRefusalReason.NoActiveWorkbook
                    : GanttValidationReportRefusalReason.TargetProtected);
        }

        if (!TryBuildColumnMap(table, out int[]? columnMap) || columnMap is null)
        {
            return GanttValidationReportOutcome.Refused(
                GanttValidationReportRefusalReason.TableMissing);
        }

        Excel.Range? body = table.DataBodyRange;

        if (body is not null)
        {
            ClearWarningFill(body);
        }

        List<GanttValidationIssue> warnings = [.. issues.Where(issue => issue.Severity == GanttValidationSeverity.Warning)];

        if (warnings.Count == 0 || body is null)
        {
            return GanttValidationReportOutcome.Ok(0);
        }

        int highlighted = 0;
        foreach (GanttValidationIssue issue in warnings)
        {
            int columnIndex = ResolveColumnIndex(columnMap, issue.Field);
            if (columnIndex < 1)
            {
                continue;
            }

            Excel.Range? cell = GetCellRange(body, issue.RowNumber, columnIndex);
            if (cell is null)
            {
                continue;
            }

#pragma warning disable CA1031
            try
            {
                cell.Interior.Color = GanttCatalogues.WarningFill.ARGB;
                highlighted++;
            }
            catch
            {
            }
#pragma warning restore CA1031
        }

        return GanttValidationReportOutcome.Ok(highlighted);
    }

    private static void ClearWarningFill(Excel.Range body)
    {
#pragma warning disable CA1031
        try
        {
            foreach (Excel.Range cell in body.Cells)
            {
                if (cell.Interior.Color == GanttCatalogues.WarningFill.ARGB)
                {
                    cell.Interior.ColorIndex = XlColorIndex.xlColorIndexNone;
                }
            }
        }
        catch
        {
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Finds the data table by walking the sheets and each sheet's tables by
    /// one-based index through <see cref="GetSheetAt"/> and
    /// <see cref="GetTableAt(ListObjects, int)"/>.
    /// </summary>
    /// <remarks>
    /// Index traversal keeps every Excel collection access behind a seam the
    /// contract tests can substitute; a COM collection enumerator cannot be stubbed
    /// through an expression tree.
    /// </remarks>
    /// <param name="sheets">The workbook's sheets collection.</param>
    /// <param name="worksheet">The sheet holding the data table, when found.</param>
    /// <param name="table">The data table, when found.</param>
    /// <returns><see langword="true"/> when the data table was found.</returns>
    private bool TryFindTable(
        Sheets sheets,
        out Worksheet? worksheet,
        out ListObject? table)
    {
        worksheet = null;
        table = null;

        try
        {
            int sheetCount = sheets.Count;
            for (int sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
            {
                if (GetSheetAt(sheets, sheetIndex) is not Worksheet ws)
                {
                    continue;
                }

                ListObjects listObjects = ws.ListObjects;
                int tableCount = listObjects.Count;
                for (int tableIndex = 1; tableIndex <= tableCount; tableIndex++)
                {
                    ListObject listObject = GetTableAt(listObjects, tableIndex);
                    if (GanttChartContract.IsDataTableName(listObject.Name))
                    {
                        worksheet = ws;
                        table = listObject;
                        return true;
                    }
                }
            }
        }
#pragma warning disable CA1031
        catch
        {
        }
#pragma warning restore CA1031

        return false;
    }

    private bool TryBuildColumnMap(
        ListObject table,
        out int[]? columnMap)
    {
        columnMap = null;
        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        int[] map = new int[schema.Count];
        Dictionary<string, int> byName = new(StringComparer.Ordinal);

        ListColumns columns = table.ListColumns;
        int columnCount = columns.Count;
        for (int index = 1; index <= columnCount; index++)
        {
            ListColumn column = GetTableAt(columns, index);
            byName[column.Name] = index;
        }

        for (int schemaIndex = 0; schemaIndex < schema.Count; schemaIndex++)
        {
            GanttTableColumn expected = schema[schemaIndex];
            if (!byName.TryGetValue(expected.Name, out int tableIndex))
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

    private static int ResolveColumnIndex(int[] columnMap, string fieldName)
    {
        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        for (int i = 0; i < schema.Count; i++)
        {
            if (string.Equals(schema[i].Name, fieldName, StringComparison.Ordinal))
            {
                return columnMap[i];
            }
        }

        return -1;
    }

    /// <summary>Returns the sheet at the one-based index. Test seam over the COM parameterised <c>Sheets.Item</c> property.</summary>
    internal virtual object GetSheetAt(Sheets sheets, int index) => sheets[index];

    /// <summary>Returns the list object at the one-based index. Test seam over the COM parameterised <c>ListObjects.Item</c> property.</summary>
    internal virtual ListObject GetTableAt(ListObjects listObjects, int index) => listObjects[index];

    /// <summary>Returns the list column at the one-based index. Test seam over the COM parameterised <c>ListColumns.Item</c> property.</summary>
    internal virtual ListColumn GetTableAt(ListColumns columns, int index) => columns[index];

    /// <summary>Returns the cell at the one-based body row and one-based table-column index. Test seam over the COM <c>Range.Cells[row, column]</c> indexer.</summary>
    internal virtual Excel.Range? GetCellRange(Excel.Range? body, int rowNumber, int columnIndex)
    {
        if (body is null)
        {
            return null;
        }

        try
        {
            return body.Cells[rowNumber, columnIndex];
        }
#pragma warning disable CA1031
        catch
        {
            return null;
        }
#pragma warning restore CA1031
    }
}
