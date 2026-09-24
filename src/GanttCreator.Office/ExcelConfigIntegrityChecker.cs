using GanttCreator.Core;
using GanttCreator.Core.ConfigIntegrity;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>Read-only Excel implementation of <see cref="IConfigIntegrityChecker"/>.</summary>
public class ExcelConfigIntegrityChecker(
    object? application,
    IConfigCatalogueReader? catalogueReader = null,
    IGanttTableReader? tableReader = null
) : IConfigIntegrityChecker
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IConfigCatalogueReader _catalogueReader = catalogueReader ?? new ExcelConfigCatalogueReader(application);
    private readonly IGanttTableReader _tableReader = tableReader ?? new ExcelGanttTableReader(application);

    /// <inheritdoc />
    public ConfigIntegrityCheckOutcome Check()
    {
        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (application is null || workbook is null)
        {
            return ConfigIntegrityCheckOutcome.NoActiveWorkbook();
        }

        List<ConfigIntegrityFinding> findings = [];
        Excel.Sheets sheets = workbook.Sheets;
        Excel.Worksheet? config = FindConfigSheet(sheets);
        if (config is null)
        {
            return ConfigIntegrityCheckOutcome.Ok([
                new ConfigIntegrityFinding(ConfigIntegrityFindingKind.ConfigSheetMissing, "Configuration worksheet is missing."),
            ]);
        }

        if (config.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden)
        {
            findings.Add(
                new ConfigIntegrityFinding(ConfigIntegrityFindingKind.WrongVisibility, "Configuration worksheet is not VeryHidden.")
            );
        }

        var worksheetCount = sheets.Count;
        for (var index = 1; index <= worksheetCount; index++)
        {
            if (
                GetSheetAt(sheets, index) is Excel.Worksheet sheet
                && string.Equals(sheet.Name, GanttWorkbookContract.ConfigSheetName, StringComparison.OrdinalIgnoreCase)
                && !ReferenceEquals(sheet, config)
            )
            {
                findings.Add(
                    new ConfigIntegrityFinding(
                        ConfigIntegrityFindingKind.SecondHelperSheet,
                        "More than one configuration worksheet exists."
                    )
                );
            }
        }

        var ganttFound =
            TryFindGanttTable(sheets, out Excel.Worksheet? gantt, out Excel.ListObject? ganttTable)
            && gantt is not null
            && ganttTable is not null;
        if (ganttFound)
        {
            var expectedAnchor = BuildPlotAnchorRefersTo(gantt!.Name, ganttTable!);
            Excel.Names sheetNames = gantt.Names;
            Excel.Name? anchor = FindName(sheetNames, GanttWorkbookContract.PlotAnchorDefinedName);
            if (anchor is null || !string.Equals(anchor.RefersTo, expectedAnchor, StringComparison.Ordinal))
            {
                findings.Add(
                    new ConfigIntegrityFinding(
                        ConfigIntegrityFindingKind.PlotAnchorDisagreement,
                        "The stored plot anchor does not match the live Gantt table."
                    )
                );
            }

            Excel.Names workbookNames = workbook.Names;
            if (FindName(workbookNames, GanttWorkbookContract.TypeOptionsDefinedName) is null)
            {
                findings.Add(
                    new ConfigIntegrityFinding(ConfigIntegrityFindingKind.TypeOptionsMissing, "The TypeOptions defined name is missing.")
                );
            }
        }

        ConfigReadOutcome catalogue = _catalogueReader.Read();
        if (!catalogue.Succeeded)
        {
            AddCatalogueFinding(findings, catalogue.Refusal!.Value);
        }

        GanttTableReadOutcome rows = _tableReader.Read();
        if (rows.Succeeded)
        {
            GanttValidationOutcome validation = GanttRowValidator.Validate(rows.Rows);
            foreach (GanttValidationIssue issue in validation.Issues.Where(issue =>
                    issue.Code is GanttValidationCodes.IdMissingOrMalformed or GanttValidationCodes.DuplicateId
                )
            )
            {
                findings.Add(
                    new ConfigIntegrityFinding(
                        ConfigIntegrityFindingKind.IdentityDamage,
                        $"Row identity requires repair at row {issue.RowNumber}.",
                        RowNumber: issue.RowNumber
                    )
                );
            }
        }

        return ConfigIntegrityCheckOutcome.Ok([.. ConfigIntegrityPlan.Build(findings).Findings]);
    }

    private static void AddCatalogueFinding(List<ConfigIntegrityFinding> findings, ConfigReadRefusalReason refusal)
    {
        ConfigIntegrityFindingKind kind = refusal switch
        {
            ConfigReadRefusalReason.NoActiveWorkbook => ConfigIntegrityFindingKind.Unknown,
            ConfigReadRefusalReason.ConfigSheetMissing => ConfigIntegrityFindingKind.ConfigSheetMissing,
            ConfigReadRefusalReason.CatalogueHashMismatch => ConfigIntegrityFindingKind.CatalogueHashMismatch,
            ConfigReadRefusalReason.TableMissing => ConfigIntegrityFindingKind.CatalogueTableMissing,
            ConfigReadRefusalReason.HeaderMismatch => ConfigIntegrityFindingKind.CatalogueTableCorrupt,
            ConfigReadRefusalReason.RowCountMismatch => ConfigIntegrityFindingKind.CatalogueTableCorrupt,
            ConfigReadRefusalReason.CatalogueMismatch => ConfigIntegrityFindingKind.CatalogueTableCorrupt,
            ConfigReadRefusalReason.ValueOutOfRange => ConfigIntegrityFindingKind.CatalogueTableCorrupt,
            ConfigReadRefusalReason.BadColourFormat => ConfigIntegrityFindingKind.CatalogueTableCorrupt,
            _ => ConfigIntegrityFindingKind.Unknown,
        };
        findings.Add(new ConfigIntegrityFinding(
            kind,
            $"Catalogue reader refused: {refusal}."));
    }

    private bool TryFindGanttTable(Excel.Sheets sheets, out Excel.Worksheet? worksheet, out Excel.ListObject? table)
    {
        worksheet = null;
        table = null;
        var count = sheets.Count;
        for (var sheetIndex = 1; sheetIndex <= count; sheetIndex++)
        {
            if (GetSheetAt(sheets, sheetIndex) is not Excel.Worksheet candidate)
            {
                continue;
            }

            Excel.ListObjects objects = candidate.ListObjects;
            var objectCount = objects.Count;
            for (var objectIndex = 1; objectIndex <= objectCount; objectIndex++)
            {
                Excel.ListObject listObject = objects[objectIndex];
                if (string.Equals(listObject.Name, GanttTableSchema.TableName, StringComparison.OrdinalIgnoreCase))
                {
                    worksheet = candidate;
                    table = listObject;
                    return true;
                }
            }
        }

        return false;
    }

    private static string BuildPlotAnchorRefersTo(string sheetName, Excel.ListObject table)
    {
        var columnIndex = table.ListColumns.Count + 1;
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

    private Excel.Worksheet? FindConfigSheet(Excel.Sheets sheets)
    {
        var count = sheets.Count;
        for (var index = 1; index <= count; index++)
        {
            if (
                GetSheetAt(sheets, index) is Excel.Worksheet worksheet
                && string.Equals(worksheet.Name, GanttWorkbookContract.ConfigSheetName, StringComparison.Ordinal)
            )
            {
                return worksheet;
            }
        }

        return null;
    }

    internal virtual Excel.Worksheet GetSheetAt(Excel.Sheets sheets, int index) => sheets[index];
}
