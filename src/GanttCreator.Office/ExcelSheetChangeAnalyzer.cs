using System.Text.RegularExpressions;
using GanttCreator.Core;
using Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="ISheetChangeAnalyzer"/> over the Excel application object.
/// Maps a changed range's address to a <c>tblGanttData</c> column and
/// reports whether it warrants an automatic refresh.
/// </summary>
/// <param name="application">
/// The Excel application object, or <see langword="null"/> when unavailable.
/// </param>
/// <remarks>
/// <para>
/// COM ownership follows the same policy as <see cref="ExcelGanttTableReader"/>:
/// all proxies are Excel-owned shared roots, held in locals, used without
/// chained member expressions, and never force-released.
/// </para>
/// <para>
/// Any chart-affecting table body column triggers automatic refresh.
/// <c>Duration</c> is derived and excluded. Header cells and changes outside
/// the active workbook's Gantt table are ignored.
/// </para>
/// </remarks>
public partial class ExcelSheetChangeAnalyzer(object? application) : ISheetChangeAnalyzer
{
    private readonly Application? _application = application as Application;

    /// <inheritdoc />
    public bool ShouldAutoRefresh(string? changedAddress)
    {
        if (string.IsNullOrEmpty(changedAddress))
        {
            return false;
        }

        Application? application = _application;
        if (application is null)
        {
            return false;
        }

        Workbook? workbook = application.ActiveWorkbook;
        if (workbook is null)
        {
            return false;
        }

        int bang = changedAddress.LastIndexOf('!');
        if (bang < 0)
        {
            return false;
        }

        string owner = changedAddress[..bang].Trim('"', '\'');
        string cellRange = changedAddress[(bang + 1)..];
        int separator = owner.LastIndexOf('|');
        if (separator < 0
            || !string.Equals(owner[..separator], workbook.Name, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        Sheets sheets = workbook.Sheets;
        if (!TryFindTable(sheets, out ListObject? table, out string? tableSheetName) || table is null
            || !string.Equals(owner[(separator + 1)..], tableSheetName, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (!TryBuildColumnMap(table, out int[]? columnMap) || columnMap is null)
        {
            return false;
        }

        Microsoft.Office.Interop.Excel.Range? body = table.DataBodyRange;
        if (body is null || !TryParseRange(cellRange, out int firstRow, out int lastRow, out int firstColumn, out int lastColumn))
        {
            return false;
        }

        int bodyFirstRow = body.Row;
        int bodyLastRow = bodyFirstRow + body.Rows.Count - 1;
        int bodyFirstColumn = body.Column;
        int bodyLastColumn = bodyFirstColumn + table.ListColumns.Count - 1;
        return RangeIntersectsBody(firstRow, lastRow, firstColumn, lastColumn,
                bodyFirstRow, bodyLastRow, bodyFirstColumn, bodyLastColumn)
            && HasChartAffectingColumn(
                GanttTableSchema.Default.Columns, columnMap, bodyFirstColumn, firstColumn, lastColumn);
    }

    internal static bool RangeIntersectsBody(
        int firstRow, int lastRow, int firstColumn, int lastColumn,
        int bodyFirstRow, int bodyLastRow, int bodyFirstColumn, int bodyLastColumn) =>
        lastRow >= bodyFirstRow && firstRow <= bodyLastRow
        && lastColumn >= bodyFirstColumn && firstColumn <= bodyLastColumn;

    internal static bool HasChartAffectingColumn(
        IReadOnlyList<GanttTableColumn> schema,
        IReadOnlyList<int> columnMap,
        int bodyFirstColumn,
        int firstChangedColumn,
        int lastChangedColumn)
    {
        ArgumentNullException.ThrowIfNull(schema);
        ArgumentNullException.ThrowIfNull(columnMap);
        if (columnMap.Count != schema.Count)
        {
            return false;
        }

        for (int index = 0; index < schema.Count; index++)
        {
            if (schema[index].Name.Equals("Duration", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            int localColumn = columnMap[index];
            if (localColumn > 0
                && IsColumnIncluded(bodyFirstColumn + localColumn - 1, firstChangedColumn, lastChangedColumn))
            {
                return true;
            }
        }

        return false;
    }

    private static bool TryFindTable(
        Sheets sheets,
        out ListObject? table,
        out string? sheetName)
    {
        table = null;
        sheetName = null;
        try
        {
            foreach (object? sheet in sheets)
            {
                if (sheet is not Worksheet ws)
                {
                    continue;
                }

                foreach (ListObject listObject in ws.ListObjects)
                {
                    if (GanttChartContract.IsDataTableName(listObject.Name))
                    {
                        table = listObject;
                        sheetName = ws.Name;
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

    private static bool TryBuildColumnMap(
        ListObject table,
        out int[]? columnMap)
    {
        columnMap = null;
        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        int[] map = new int[schema.Count];

        Dictionary<string, int>? byName = TryReadHeaderNamesInOneCall(table);
        if (byName is null)
        {
            // A header the bulk payload cannot express as text (a numeric or blank header
            // cell, where the table's own column name and the cell value can differ) falls
            // back to the column names -- exactly the map this built before.
            byName = new Dictionary<string, int>(StringComparer.Ordinal);
            ListColumns columns = table.ListColumns;
            int columnCount = columns.Count;
            for (int index = 1; index <= columnCount; index++)
            {
                ListColumn column = columns[index];
                byName[column.Name] = index;
            }
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

    /// <summary>
    /// Reads the table's header row in ONE bulk <c>Value2</c> call and maps each header
    /// text to its one-based table-column index.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <returns>
    /// The map, or <see langword="null"/> when the header cannot be read this way and the
    /// caller must read the table's column names individually.
    /// </returns>
    /// <remarks>
    /// <para>
    /// <b>Why one call.</b> This runs on every worksheet change, because the change's
    /// answer decides whether an automatic refresh is scheduled. The per-column
    /// <c>ListColumn.Name</c> reads it replaces cost one COM round trip each, so a paste
    /// that raises many change events paid for the whole column collection each time.
    /// The header row IS the column names, so the map is built from the same authority.
    /// </para>
    /// <para>
    /// A payload that is not a single row of non-empty text returns
    /// <see langword="null"/> rather than a partial map: the caller then reads the
    /// column names, so a sheet Excel names differently from its header cells is handled
    /// by the old path instead of being guessed at.
    /// </para>
    /// </remarks>
    private static Dictionary<string, int>? TryReadHeaderNamesInOneCall(ListObject table)
    {
        ArgumentNullException.ThrowIfNull(table);

        try
        {
            Microsoft.Office.Interop.Excel.Range? header = table.HeaderRowRange;
            if (header is null)
            {
                return null;
            }

            object? raw = header.Value2;
            List<object?[]> rows = ExcelValue2Matrix.ReadRows(raw);
            if (rows.Count != 1 || rows[0].Length == 0)
            {
                return null;
            }

            Dictionary<string, int> byName = new(StringComparer.Ordinal);
            for (int index = 0; index < rows[0].Length; index++)
            {
                if (rows[0][index] is not string name || name.Length == 0)
                {
                    return null;
                }

                byName[name] = index + 1;
            }

            return byName;
        }
#pragma warning disable CA1031
        catch
        {
            // A dead or disconnected proxy is the same answer as an unreadable header:
            // fall back to the per-column read, which fails the same way if the host is
            // genuinely gone. The caller treats a null map as "names unavailable".
            return null;
        }
#pragma warning restore CA1031
    }

    internal static bool TryParseRange(string address, out int firstRow, out int lastRow, out int firstColumn, out int lastColumn)
    {
        firstRow = lastRow = firstColumn = lastColumn = 0;
        string[] areas = address.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (areas.Length == 0)
        {
            return false;
        }

        foreach (string area in areas)
        {
            Match match = RangeRegex().Match(area);
            if (!match.Success
                || !int.TryParse(match.Groups[2].Value, out int row1)
                || !int.TryParse(match.Groups[4].Success ? match.Groups[4].Value : match.Groups[2].Value, out int row2))
            {
                firstRow = lastRow = firstColumn = lastColumn = 0;
                return false;
            }

            int areaFirstColumn = ColumnNumber(match.Groups[1].Value);
            int areaLastColumn = match.Groups[3].Success ? ColumnNumber(match.Groups[3].Value) : areaFirstColumn;
            int areaFirstRow = Math.Min(row1, row2);
            int areaLastRow = Math.Max(row1, row2);
            if (areaFirstRow < 1 || areaFirstColumn < 1 || areaLastColumn < 1)
            {
                firstRow = lastRow = firstColumn = lastColumn = 0;
                return false;
            }

            firstRow = firstRow == 0 ? areaFirstRow : Math.Min(firstRow, areaFirstRow);
            lastRow = Math.Max(lastRow, areaLastRow);
            firstColumn = firstColumn == 0 ? areaFirstColumn : Math.Min(firstColumn, areaFirstColumn);
            lastColumn = Math.Max(lastColumn, areaLastColumn);
        }

        return true;
    }

    internal static bool IsColumnIncluded(int column, int firstColumn, int lastColumn) =>
        column >= firstColumn && column <= lastColumn;

    [GeneratedRegex(@"^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex RangeRegex();

    private static int ColumnNumber(string letters) => letters.Aggregate(0, (value, character) => (value * 26) + char.ToUpperInvariant(character) - 'A' + 1);
}
