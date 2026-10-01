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
            var refersTo = GanttSheetLayout.BuildPlotAnchorRefersTo(gantt.Name);
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

    internal virtual Excel.Worksheet GetSheetAt(Excel.Sheets sheets, int index) => sheets[index];

    internal virtual Excel.ListObject GetTableAt(Excel.ListObjects objects, int index) => objects[index];
}
