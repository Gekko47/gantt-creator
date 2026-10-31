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
/// Only <c>Id</c> (column 1) and <c>Type</c> (column 4) changes trigger
/// auto-refresh, per the R5.8 decision Q3: type/id changes refresh
/// automatically; start/finish require a manual Refresh; duration is
/// engine-driven.
/// </para>
/// </remarks>
public class ExcelSheetChangeAnalyzer(object? application) : ISheetChangeAnalyzer
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

        Sheets sheets = workbook.Sheets;
        if (!TryFindTable(sheets, out ListObject? table) || table is null)
        {
            return false;
        }

        if (!TryBuildColumnMap(table, out int[]? columnMap) || columnMap is null)
        {
            return false;
        }

        // Parse the Excel address ($C$5 → column index 3). The address may
        // refer to a multi-cell range; use the top-left cell's column.
        int columnIndex = ParseColumnIndex(changedAddress);
        if (columnIndex < 1)
        {
            return false;
        }

        // Schema column ordinals are 0-based; table column indices are 1-based.
        // Id is schema[0], Type is schema[3].
        bool isIdColumn = columnMap[0] == columnIndex;
        bool isTypeColumn = columnMap[3] == columnIndex;
        return isIdColumn || isTypeColumn;
    }

    private static bool TryFindTable(
        Sheets sheets,
        out ListObject? table)
    {
        table = null;
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
                    if (string.Equals(listObject.Name, GanttTableSchema.TableName,
                            StringComparison.OrdinalIgnoreCase))
                    {
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

    private static bool TryBuildColumnMap(
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
            ListColumn column = columns[index];
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

    /// <summary>
    /// Parses the one-based column index from an Excel-style address such as
    /// <c>"$C$5"</c> or <c>"C5"</c>. Returns 0 when unparseable.
    /// </summary>
    private static int ParseColumnIndex(string address)
    {
        try
        {
            int start = 0;
            if (address.Length > 0 && address[0] == '$')
            {
                start = 1;
            }

            int end = start;
            while (end < address.Length && char.IsLetter(address[end]))
            {
                end++;
            }

            if (end == start)
            {
                return 0;
            }

            string letters = address[start..end];
            int column = 0;
            foreach (char c in letters)
            {
                column = (column * 26) + (char.ToUpperInvariant(c) - 'A') + 1;
            }

            return column;
        }
#pragma warning disable CA1031
        catch
        {
            return 0;
        }
#pragma warning restore CA1031
    }
}
