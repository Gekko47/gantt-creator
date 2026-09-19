using GanttCreator.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="IGanttTableReader"/> over the Excel application object
/// supplied by the host at add-in load. Bulk-reads <c>tblGanttData</c> via
/// one <c>DataBodyRange.Value2</c> call; never writes cells, creates sheets,
/// or changes application state.
/// </summary>
/// <param name="application">
/// The Excel application object (for example <c>ExcelDnaUtil.Application</c>), or
/// <see langword="null"/> when the host supplied none (unit tests, non-Excel
/// host). A foreign object fails the interface cast and degrades to the
/// no-active-workbook refusal with no mutation.
/// </param>
/// <remarks>
/// <para>
/// COM ownership: the <c>Application</c>, <c>Workbook</c>, <c>Worksheet</c>,
/// <c>Range</c>, and <c>ListObject</c> objects reached here are Excel-owned
/// shared roots. This adapter takes no ownership of them, never calls
/// <c>FinalReleaseComObject</c>, and force-releases nothing. Every proxy is
/// held in a local and used without chained member expressions.
/// </para>
/// <para>
/// The <c>internal virtual</c> accessors isolate the Excel COM parameterised
/// properties (indexers). Expression trees cannot contain indexed properties
/// (CS0855), so contract tests substitute these seams and every other member
/// through Moq; the real indexer behaviour is exercised by the tagged
/// live-Office integration test.
/// </para>
/// </remarks>
public class ExcelGanttTableReader(object? application) : IGanttTableReader
{
    private readonly Application? _application = application as Application;

    /// <inheritdoc />
    public GanttTableReadOutcome Read()
    {
        Application? application = _application;
        if (application is null)
        {
            return GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.NoActiveWorkbook);
        }

        Workbook? workbook = application.ActiveWorkbook;
        if (workbook is null)
        {
            return GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.NoActiveWorkbook);
        }

        if (GetDate1904(workbook))
        {
            return GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.DateSystemUnsupported);
        }

        Sheets sheets = workbook.Sheets;
        if (!TryFindTable(sheets, out ListObject? table) || table is null)
        {
            return GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.TableMissing);
        }

        if (!TryBuildColumnMap(table, out var columnMap) || columnMap is null)
        {
            return GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.TableMissing);
        }

        Excel.Range? body = table.DataBodyRange;
        if (body is null)
        {
            return GanttTableReadOutcome.Ok([]);
        }

        var raw = GetBodyValues(body);
        List<GanttRowDto> rows = ConvertBody(raw, columnMap);
        return GanttTableReadOutcome.Ok(rows);
    }

    /// <summary>
    /// Finds <c>tblGanttData</c> on any worksheet. Chart sheets carry no
    /// list objects and are skipped via the worksheet cast.
    /// </summary>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <param name="table">The table when found.</param>
    /// <returns><see langword="true"/> when the table was found.</returns>
    private bool TryFindTable(Sheets sheets, out ListObject? table)
    {
        table = null;
        var count = sheets.Count;
        for (var index = 1; index <= count; index++)
        {
            var sheet = GetSheetAt(sheets, index);
            if (sheet is Worksheet worksheet)
            {
                ListObjects listObjects = worksheet.ListObjects;
                var tableCount = listObjects.Count;
                for (var tableIndex = 1; tableIndex <= tableCount; tableIndex++)
                {
                    ListObject candidate = GetTableAt(listObjects, tableIndex);
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
        }

        return false;
    }

    /// <summary>
    /// Builds the schema-column to table-column index map from live header
    /// names (Ordinal), so a reordered table still maps correctly.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="columnMap">Per-schema-column one-based table indexes, or <see langword="null"/>.</param>
    /// <returns><see langword="true"/> when every required header resolved.</returns>
    private bool TryBuildColumnMap(ListObject table, out int[]? columnMap)
    {
        columnMap = null;
        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        var map = new int[schema.Count];
        var byName = new Dictionary<string, int>(StringComparer.Ordinal);
        ListColumns columns = table.ListColumns;
        var columnCount = columns.Count;
        for (var index = 1; index <= columnCount; index++)
        {
            ListColumn column = GetColumnAt(columns, index);
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


    /// <summary>
    /// Converts the bulk <c>Value2</c> payload into DTOs in body order.
    /// An empty body yields no rows. Excel error cells are
    /// <see cref="int"/> and map to <see langword="null"/> fields with the
    /// row preserved.
    /// </summary>
    /// <param name="raw">The <c>Value2</c> payload.</param>
    /// <param name="columnMap">The schema-column to table-column index map.</param>
    /// <returns>The rows in body order.</returns>
    private static List<GanttRowDto> ConvertBody(object? raw, int[] columnMap)
    {
        if (raw is not Array matrix)
        {
            return [];
        }

        var rowCount = matrix.GetLength(0);
        var rows = new List<GanttRowDto>(rowCount);
        for (var row = 1; row <= rowCount; row++)
        {
            var cells = new object?[columnMap.Length];
            for (var schemaIndex = 0; schemaIndex < columnMap.Length; schemaIndex++)
            {
                var tableIndex = columnMap[schemaIndex];
                cells[schemaIndex] = tableIndex < 1 ? null : CoerceCell(matrix.GetValue(row, tableIndex));
            }

            rows.Add(ConvertRow(row, cells));
        }

        return rows;
    }

    /// <summary>
    /// Coerces one matrix cell: COM empty markers become <see langword="null"/>.
    /// </summary>
    /// <param name="cell">The raw matrix cell.</param>
    /// <returns>The coerced payload.</returns>
    private static object? CoerceCell(object? cell) => cell is null or System.Reflection.Missing ? null : cell;


    /// <summary>
    /// Converts one body row of cell payloads into a DTO. Text is trimmed;
    /// dates, stack index, and visibility go through
    /// <see cref="ExcelCellConverter"/>; <c>SortOrder</c> stays raw text.
    /// </summary>
    /// <param name="rowNumber">The one-based body-row index.</param>
    /// <param name="cells">The per-schema-column cell payloads.</param>
    /// <returns>The neutral DTO.</returns>
    private static GanttRowDto ConvertRow(int rowNumber, object?[] cells)
    {
        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        var byName = new Dictionary<string, object?>(StringComparer.Ordinal);
        for (var index = 0; index < schema.Count; index++)
        {
            byName[schema[index].Name] = cells[index];
        }

        _ = ExcelCellConverter.TryConvertDate(byName["Start"], out DateOnly? start);
        _ = ExcelCellConverter.TryConvertDate(byName["Finish"], out DateOnly? finish);
        _ = ExcelCellConverter.TryConvertStackIndex(byName["StackIndex"], out var stackIndex);
        _ = ExcelCellConverter.TryConvertVisible(byName["Visible"], out var visible);

        return new GanttRowDto(
            rowNumber,
            ExcelCellConverter.ToText(byName["Id"]),
            ExcelCellConverter.ToText(byName["LaneId"]),
            stackIndex,
            ExcelCellConverter.ToText(byName["Type"]),
            ExcelCellConverter.ToText(byName["Description"]),
            start,
            finish,
            ExcelCellConverter.ToText(byName["ParentId"]),
            ExcelCellConverter.ToText(byName["StyleKey"]),
            ExcelCellConverter.ToText(byName["LabelPosition"]),
            ExcelCellConverter.ToText(byName["FillColour"]),
            ExcelCellConverter.ToText(byName["StrokeColour"]),
            visible,
            ExcelCellConverter.ToText(byName["SortOrder"]));
    }

    /// <summary>
    /// Reads the workbook date-system flag. Test seam over the COM property.
    /// </summary>
    /// <param name="workbook">The active workbook.</param>
    /// <returns><see langword="true"/> when the workbook uses the 1904 date system.</returns>
    internal virtual bool GetDate1904(Workbook workbook) => workbook.Date1904;

    /// <summary>
    /// Returns the sheet at the one-based index. Test seam over the COM
    /// parameterised <c>Sheets.Item</c> property.
    /// </summary>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <param name="index">The one-based sheet index.</param>
    /// <returns>The sheet at the index.</returns>
    internal virtual object GetSheetAt(Sheets sheets, int index) => sheets[index];

    /// <summary>
    /// Returns the list object at the one-based index. Test seam over the COM
    /// parameterised <c>ListObjects.Item</c> property.
    /// </summary>
    /// <param name="listObjects">The worksheet's list objects.</param>
    /// <param name="index">The one-based table index.</param>
    /// <returns>The list object at the index.</returns>
    internal virtual ListObject GetTableAt(ListObjects listObjects, int index) => listObjects[index];

    /// <summary>
    /// Returns the list column at the one-based index. Test seam over the COM
    /// parameterised <c>ListColumns.Item</c> property.
    /// </summary>
    /// <param name="columns">The table's list columns.</param>
    /// <param name="index">The one-based column index.</param>
    /// <returns>The list column at the index.</returns>
    internal virtual ListColumn GetColumnAt(ListColumns columns, int index) => columns[index];

    /// <summary>
    /// Returns the bulk <c>Value2</c> payload of the table body. Test seam so
    /// contract tests inject <c>object[,]</c> without COM.
    /// </summary>
    /// <param name="body">The table body range.</param>
    /// <returns>The <c>Value2</c> payload.</returns>
    internal virtual object? GetBodyValues(Excel.Range body) => body.Value2;
}
