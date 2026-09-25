using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>Live Excel adapter for the TypeOptions name and Type-column list rule.</summary>
public class ExcelTypeOptionsMaterialiser(
    object? application,
    IConfigCatalogueReader? catalogueReader = null,
    IWorksheetProtectionGuard? protectionGuard = null) : ITypeOptionsMaterialiser
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IConfigCatalogueReader _catalogueReader =
        catalogueReader ?? new ExcelConfigCatalogueReader(application);
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <inheritdoc />
    public TypeOptionsMaterialiseOutcome Materialise()
    {
        ProtectionGuardOutcome activeProtection = _protectionGuard.Query();
        if (activeProtection != ProtectionGuardOutcome.NotProtected)
        {
            return TypeOptionsMaterialiseOutcome.Refused(
                activeProtection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? TypeOptionsRefusalReason.NoActiveWorkbook
                    : TypeOptionsRefusalReason.TargetProtected);
        }

        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (workbook is null)
        {
            return TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.NoActiveWorkbook);
        }

        ConfigReadOutcome catalogue = _catalogueReader.Read();
        if (!catalogue.Succeeded)
        {
            return TypeOptionsMaterialiseOutcome.Refused(
                catalogue.Refusal == ConfigReadRefusalReason.ConfigSheetMissing
                    ? TypeOptionsRefusalReason.ConfigSheetMissing
                    : catalogue.Refusal == ConfigReadRefusalReason.TableMissing
                        ? TypeOptionsRefusalReason.TableMissing
                        : TypeOptionsRefusalReason.CatalogueHashMismatch);
        }

        Excel.Sheets sheets = workbook.Sheets;
        Excel.Worksheet? config = FindSheet(sheets, GanttWorkbookContract.ConfigSheetName);
        Excel.Worksheet? gantt = FindTableSheet(sheets, GanttTableSchema.TableName);
        if (config is null || gantt is null)
        {
            return TypeOptionsMaterialiseOutcome.Refused(
                config is null
                    ? TypeOptionsRefusalReason.ConfigSheetMissing
                    : TypeOptionsRefusalReason.TableMissing);
        }

        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(config);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return TypeOptionsMaterialiseOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? TypeOptionsRefusalReason.NoActiveWorkbook
                    : TypeOptionsRefusalReason.TargetProtected);
        }

        Excel.ListObject? types = FindTable(config, GanttCatalogues.TypesTableName);
        Excel.ListObject? data = FindTable(gantt, GanttTableSchema.TableName);
        if (types is null || data is null)
        {
            return TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.TableMissing);
        }

        Excel.ListColumn? typeColumn = FindColumn(data, "Type");
        Excel.ListColumn? displayNameColumn = FindColumn(types, "DisplayName");
        if (typeColumn is null || displayNameColumn is null)
        {
            return TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.TableMissing);
        }

        Excel.Range? optionsRange = displayNameColumn.DataBodyRange;
        if (optionsRange is null)
        {
            return TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.CatalogueHashMismatch);
        }

        Excel.Names names = workbook.Names;
        Excel.Name? existingName = FindName(names, GanttWorkbookContract.TypeOptionsDefinedName);
        var expectedRefersTo = BuildRefersTo(config.Name, optionsRange);
        if (existingName is not null
            && !string.Equals(existingName.RefersTo, expectedRefersTo, StringComparison.Ordinal))
        {
            return TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.NameTargetInvalid);
        }

        SetName(names, GanttWorkbookContract.TypeOptionsDefinedName, expectedRefersTo);
        Excel.Range? typeRange = typeColumn.DataBodyRange;
        if (typeRange is not null)
        {
            Excel.Validation validation = typeRange.Validation;
            validation.Delete();
            validation.Add(
                Excel.XlDVType.xlValidateList,
                Excel.XlDVAlertStyle.xlValidAlertStop,
                Type.Missing,
                $"={GanttWorkbookContract.TypeOptionsDefinedName}",
                Type.Missing);
            validation.InCellDropdown = true;
            validation.IgnoreBlank = false;
            validation.ShowError = true;
        }

        return TypeOptionsMaterialiseOutcome.Ok();
    }

    private static string BuildRefersTo(string sheetName, Excel.Range range)
    {
        var address = range.Address[true, true]
            ?? throw new InvalidOperationException("TypeOptions range has no address.");
        return $"='{sheetName.Replace("'", "''", StringComparison.Ordinal)}'!{address}";
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

    private static void SetName(Excel.Names names, string nameText, string refersTo)
    {
        Excel.Name? existing = FindName(names, nameText);
        if (existing is null)
        {
            _ = names.Add(nameText, refersTo);
        }
        else
        {
            existing.RefersTo = refersTo;
        }
    }

    private static Excel.Worksheet? FindSheet(Excel.Sheets sheets, string name)
    {
        var count = sheets.Count;
        for (var index = 1; index <= count; index++)
        {
            if (sheets[index] is Excel.Worksheet sheet
                && string.Equals(sheet.Name, name, StringComparison.Ordinal))
            {
                return sheet;
            }
        }

        return null;
    }

    private static Excel.Worksheet? FindTableSheet(Excel.Sheets sheets, string tableName)
    {
        var count = sheets.Count;
        for (var index = 1; index <= count; index++)
        {
            if (sheets[index] is Excel.Worksheet sheet && FindTable(sheet, tableName) is not null)
            {
                return sheet;
            }
        }

        return null;
    }

    private static Excel.ListObject? FindTable(Excel.Worksheet sheet, string name)
    {
        Excel.ListObjects objects = sheet.ListObjects;
        var count = objects.Count;
        for (var index = 1; index <= count; index++)
        {
            Excel.ListObject table = objects[index];
            if (string.Equals(table.Name, name, StringComparison.Ordinal))
            {
                return table;
            }
        }

        return null;
    }

    private static Excel.ListColumn? FindColumn(Excel.ListObject table, string name)
    {
        Excel.ListColumns columns = table.ListColumns;
        var count = columns.Count;
        for (var index = 1; index <= count; index++)
        {
            Excel.ListColumn column = columns[index];
            if (string.Equals(column.Name, name, StringComparison.Ordinal))
            {
                return column;
            }
        }

        return null;
    }
}
