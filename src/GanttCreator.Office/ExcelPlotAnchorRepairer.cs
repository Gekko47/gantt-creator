using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>Excel plot-anchor repair adapter.</summary>
public class ExcelPlotAnchorRepairer(object? application, IWorksheetProtectionGuard? protectionGuard = null) : IPlotAnchorRepairer
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IWorksheetProtectionGuard _protectionGuard = protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <inheritdoc />
    public PlotAnchorRepairOutcome Repair()
    {
        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (workbook is null)
        {
            return PlotAnchorRepairOutcome.Refused(PlotAnchorRepairRefusalReason.NoActiveWorkbook);
        }

        if (!TryFindGanttTable(workbook.Sheets, out Excel.Worksheet? gantt, out Excel.ListObject? table) || gantt is null || table is null)
        {
            return PlotAnchorRepairOutcome.Refused(PlotAnchorRepairRefusalReason.TableMissing);
        }

        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(gantt);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return PlotAnchorRepairOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? PlotAnchorRepairRefusalReason.NoActiveWorkbook
                    : PlotAnchorRepairRefusalReason.TargetProtected
            );
        }

        try
        {
            Excel.Names names = gantt.Names;
            var refersTo = BuildRefersTo(gantt.Name, table.ListColumns.Count + 1);
            Excel.Name? existing = FindName(names, GanttWorkbookContract.PlotAnchorDefinedName);
            if (existing is null)
            {
                _ = names.Add(GanttWorkbookContract.PlotAnchorDefinedName, refersTo);
            }
            else
            {
                existing.RefersTo = refersTo;
            }

            return PlotAnchorRepairOutcome.Ok();
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            return PlotAnchorRepairOutcome.Refused(PlotAnchorRepairRefusalReason.NameWriteFailed);
        }
    }

    private bool TryFindGanttTable(
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

    private static Excel.Name? FindName(Excel.Names names, string nameText)
    {
        var count = names.Count;
        for (var index = 1; index <= count; index++)
        {
            Excel.Name candidate = names.Item(index);
            if (string.Equals(candidate.Name, nameText, StringComparison.OrdinalIgnoreCase))
            {
                return candidate;
            }
        }

        return null;
    }

    private static string BuildRefersTo(string sheetName, int columnIndex)
    {
        var escaped = sheetName.Replace("'", "''", StringComparison.Ordinal);
        return $"='{escaped}'!${ToA1Column(columnIndex)}$1";
    }

    private static string ToA1Column(int columnIndex)
    {
        var builder = new System.Text.StringBuilder();
        var remaining = columnIndex;
        while (remaining > 0)
        {
            var digit = (remaining - 1) % 26;
            _ = builder.Insert(0, (char)('A' + digit));
            remaining = (remaining - 1) / 26;
        }

        return builder.ToString();
    }

    internal virtual Excel.Worksheet GetSheetAt(Excel.Sheets sheets, int index) => sheets[index];

    internal virtual Excel.ListObject GetTableAt(Excel.ListObjects objects, int index) => objects[index];
}
