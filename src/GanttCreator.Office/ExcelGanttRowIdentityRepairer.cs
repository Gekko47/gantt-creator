using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>Excel identity repair adapter for the visible Gantt table.</summary>
public class ExcelGanttRowIdentityRepairer(object? application, IWorksheetProtectionGuard? protectionGuard = null)
    : IGanttRowIdentityRepairer
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IWorksheetProtectionGuard _protectionGuard = protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <inheritdoc />
    public GanttRowIdentityRepairOutcome Repair()
    {
        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (workbook is null || !TryFindTable(workbook.Sheets, out Excel.Worksheet? worksheet, out Excel.ListObject? table) || worksheet is null || table is null)
        {
            return GanttRowIdentityRepairOutcome.Refused(GanttRowIdentityRepairRefusalReason.TableMissing);
        }

        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(worksheet);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return GanttRowIdentityRepairOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? GanttRowIdentityRepairRefusalReason.NoActiveWorkbook
                    : GanttRowIdentityRepairRefusalReason.TargetProtected
            );
        }

        Excel.Range? body = GetTableBody(table);
        if (body is null)
        {
            return GanttRowIdentityRepairOutcome.Ok(0);
        }

        List<object?[]> rows = ExcelValue2Matrix.ReadRows(GetBodyValues(body));
        var idIndex = FindColumnIndex(table, "Id");
        var parentIdIndex = FindColumnIndex(table, "ParentId");
        if (idIndex < 0)
        {
            return GanttRowIdentityRepairOutcome.Refused(GanttRowIdentityRepairRefusalReason.TableMissing);
        }

        var rowCount = rows.Count;
        var originalIds = new string[rowCount];
        var canonicalIndexes = new Dictionary<string, int>(StringComparer.Ordinal);
        var idOccurrences = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            originalIds[rowIndex] = ToText(rows[rowIndex].ElementAtOrDefault(idIndex)).Trim();
            _ = canonicalIndexes.TryAdd(originalIds[rowIndex], rowIndex);
            _ = idOccurrences.TryGetValue(originalIds[rowIndex], out var occurrences);
            idOccurrences[originalIds[rowIndex]] = occurrences + 1;
        }

        var replacements = new Dictionary<int, GanttRowId>();
        var used = new HashSet<string>(StringComparer.Ordinal);
        for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            var original = originalIds[rowIndex];
            if (GanttRowId.TryParse(original, out _))
            {
                _ = used.Add(original);
            }
        }
        for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            var original = originalIds[rowIndex];
            var malformed = !GanttRowId.TryParse(original, out _);
            var duplicate = !canonicalIndexes.TryGetValue(original, out var canonicalIndex) || canonicalIndex != rowIndex;
            if (malformed || duplicate)
            {
                GanttRowId replacement = NewUniqueId(used);
                _ = used.Add(replacement.Value);
                replacements[rowIndex] = replacement;
            }
        }

        if (parentIdIndex >= 0)
        {
            for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                var parentText = ToText(rows[rowIndex].ElementAtOrDefault(parentIdIndex)).Trim();
                if (parentText.Length == 0 || !replacements.Any(pair => string.Equals(originalIds[pair.Key], parentText, StringComparison.Ordinal)))
                {
                    continue;
                }

                if (idOccurrences[parentText] > 1)
                {
                    continue;
                }
            }
        }

        var repaired = 0;
        foreach (KeyValuePair<int, GanttRowId> pair in replacements.OrderBy(pair => pair.Key))
        {
            GetCell(body, pair.Key + 1, idIndex + 1).Value2 = pair.Value.Value;
            repaired++;
        }

        if (parentIdIndex >= 0)
        {
            for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                var parentText = ToText(rows[rowIndex].ElementAtOrDefault(parentIdIndex)).Trim();
                if (canonicalIndexes.TryGetValue(parentText, out var parentRowIndex)
                    && replacements.TryGetValue(parentRowIndex, out GanttRowId? replacement)
                    && replacement is not null)
                {
                    GetCell(body, rowIndex + 1, parentIdIndex + 1).Value2 = replacement.Value;
                    repaired++;
                }
            }
        }

        return GanttRowIdentityRepairOutcome.Ok(repaired);
    }

    private static GanttRowId NewUniqueId(HashSet<string> used)
    {
        GanttRowId id;
        do
        {
            id = GanttRowId.New();
        } while (used.Contains(id.Value));

        return id;
    }

    private bool TryFindTable(
        Excel.Sheets sheets,
        out Excel.Worksheet? worksheet,
        out Excel.ListObject? table)
    {
        worksheet = null;
        table = null;
        var sheetCount = sheets.Count;
        for (var sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
        {
            if (GetSheetAt(sheets, sheetIndex) is not Excel.Worksheet candidate)
            {
                continue;
            }

            Excel.ListObjects objects = candidate.ListObjects;
            var objectCount = objects.Count;
            for (var objectIndex = 1; objectIndex <= objectCount; objectIndex++)
            {
                Excel.ListObject candidateTable = GetTableAt(objects, objectIndex);
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

    private static int FindColumnIndex(Excel.ListObject table, string name)
    {
        Excel.ListColumns columns = table.ListColumns;
        var count = columns.Count;
        for (var index = 1; index <= count; index++)
        {
            if (string.Equals(columns[index].Name, name, StringComparison.Ordinal))
            {
                return index - 1;
            }
        }

        return -1;
    }

    private static string ToText(object? value) =>
        Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;

    private static Excel.Range GetCell(Excel.Range body, int row, int column)
    {
        Excel.Range cells = body.Cells;
        return cells[row, column];
    }

    internal virtual Excel.Worksheet GetSheetAt(Excel.Sheets sheets, int index) => sheets[index];

    internal virtual Excel.ListObject GetTableAt(Excel.ListObjects objects, int index) => objects[index];

    internal virtual Excel.Range? GetTableBody(Excel.ListObject table) => table.DataBodyRange;

    internal virtual object? GetBodyValues(Excel.Range body) => body.Value2;
}
