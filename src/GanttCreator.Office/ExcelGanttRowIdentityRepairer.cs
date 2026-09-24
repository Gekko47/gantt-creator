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
        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return GanttRowIdentityRepairOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? GanttRowIdentityRepairRefusalReason.NoActiveWorkbook
                    : GanttRowIdentityRepairRefusalReason.TargetProtected
            );
        }

        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (workbook is null || !TryFindTable(workbook.Sheets, out Excel.ListObject? table) || table is null)
        {
            return GanttRowIdentityRepairOutcome.Refused(GanttRowIdentityRepairRefusalReason.TableMissing);
        }

        Excel.Range? body = GetTableBody(table);
        if (body is null)
        {
            return GanttRowIdentityRepairOutcome.Ok(0);
        }

        List<object?[]> rows = ExcelValue2Matrix.ReadRows(GetBodyValues(body));
        var idIndex = FindColumnIndex(table, "Id");
        if (idIndex < 0)
        {
            return GanttRowIdentityRepairOutcome.Refused(GanttRowIdentityRepairRefusalReason.TableMissing);
        }

        var seen = new HashSet<string>(StringComparer.Ordinal);
        var used = new HashSet<string>(StringComparer.Ordinal);
        var repaired = 0;
        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var text = ToText(rows[rowIndex].ElementAtOrDefault(idIndex)).Trim();
            if (!GanttRowId.TryParse(text, out GanttRowId? parsed) || parsed is null || !seen.Add(text))
            {
                GanttRowId replacement = NewUniqueId(used);
                _ = used.Add(replacement.Value);
                GetCell(body, rowIndex + 1, idIndex + 1).Value2 = replacement.Value;
                repaired++;
                continue;
            }

            _ = used.Add(text);
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

    private bool TryFindTable(Excel.Sheets sheets, out Excel.ListObject? table)
    {
        table = null;
        var count = sheets.Count;
        for (var index = 1; index <= count; index++)
        {
            if (GetSheetAt(sheets, index) is not Excel.Worksheet worksheet)
            {
                continue;
            }

            Excel.ListObjects objects = worksheet.ListObjects;
            var objectCount = objects.Count;
            for (var objectIndex = 1; objectIndex <= objectCount; objectIndex++)
            {
                Excel.ListObject candidate = GetTableAt(objects, objectIndex);
                if (string.Equals(candidate.Name, GanttTableSchema.TableName, StringComparison.OrdinalIgnoreCase))
                {
                    table = candidate;
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
