using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>Excel adapter for configuration-sheet visibility repair.</summary>
public class ExcelConfigSheetVisibilityRepairer(object? application, IWorksheetProtectionGuard? protectionGuard = null)
    : IConfigSheetVisibilityRepairer
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IWorksheetProtectionGuard _protectionGuard = protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <inheritdoc />
    public ConfigSheetVisibilityRepairOutcome Repair()
    {
        ProtectionGuardOutcome activeProtection = _protectionGuard.Query();
        if (activeProtection != ProtectionGuardOutcome.NotProtected)
        {
            return ConfigSheetVisibilityRepairOutcome.Refused(
                activeProtection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? ConfigSheetVisibilityRepairRefusalReason.NoActiveWorkbook
                    : ConfigSheetVisibilityRepairRefusalReason.TargetProtected);
        }

        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (workbook is null)
        {
            return ConfigSheetVisibilityRepairOutcome.Refused(ConfigSheetVisibilityRepairRefusalReason.NoActiveWorkbook);
        }

        Excel.Sheets sheets = workbook.Sheets;
        var count = sheets.Count;
        for (var index = 1; index <= count; index++)
        {
            if (GetSheetAt(sheets, index) is Excel.Worksheet sheet
                && string.Equals(sheet.Name, GanttWorkbookContract.ConfigSheetName, StringComparison.OrdinalIgnoreCase))
            {
                ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(sheet);
                if (protection != ProtectionGuardOutcome.NotProtected)
                {
                    return ConfigSheetVisibilityRepairOutcome.Refused(
                        protection == ProtectionGuardOutcome.NoActiveWorkbook
                            ? ConfigSheetVisibilityRepairRefusalReason.NoActiveWorkbook
                            : ConfigSheetVisibilityRepairRefusalReason.TargetProtected);
                }

                sheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                return ConfigSheetVisibilityRepairOutcome.Ok();
            }
        }

        return ConfigSheetVisibilityRepairOutcome.Refused(ConfigSheetVisibilityRepairRefusalReason.ConfigSheetMissing);
    }

    internal virtual Excel.Worksheet GetSheetAt(Excel.Sheets sheets, int index) => sheets[index];
}
