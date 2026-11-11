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

        int worksheetCount = sheets.Count;
        for (int index = 1; index <= worksheetCount; index++)
        {
            if (
                GetSheetAt(sheets, index) is Excel.Worksheet sheet
                && GanttWorkbookContract.IsConfigSheetName(sheet.Name)
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

        // Schedule data on the configuration worksheet is refused, never repaired
        // (ADR-0011): the five catalogue tables are all that sheet may carry, so any
        // other table there is either a copied schedule or the user's own data
        // standing where the add-in expects a catalogue. Detected by TABLE rather
        // than by cell, because a stray cell is a note and a table is data — and
        // because a cell-level rule would report every scratch pad as damage.
        foreach (string foreignTable in FindForeignConfigTables(config))
        {
            findings.Add(
                new ConfigIntegrityFinding(
                    ConfigIntegrityFindingKind.ScheduleDataOnConfig,
                    $"The configuration worksheet carries a table named '{foreignTable}', which is not one of the five add-in catalogue tables. Move that data to a visible worksheet and remove the table; this state is reported and never repaired automatically.",
                    TableName: foreignTable
                )
            );
        }

        bool ganttFound =
            TryFindGanttTable(sheets, out Excel.Worksheet? gantt, out Excel.ListObject? ganttTable)
            && gantt is not null
            && ganttTable is not null;
        if (ganttFound)
        {
            string expectedAnchor = GanttSheetLayout.BuildPlotAnchorRefersTo(gantt!.Name);
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
        int count = sheets.Count;
        for (int sheetIndex = 1; sheetIndex <= count; sheetIndex++)
        {
            if (GetSheetAt(sheets, sheetIndex) is not Excel.Worksheet candidate)
            {
                continue;
            }

            Excel.ListObjects objects = candidate.ListObjects;
            int objectCount = objects.Count;
            for (int objectIndex = 1; objectIndex <= objectCount; objectIndex++)
            {
                Excel.ListObject listObject = objects[objectIndex];
                if (GanttChartContract.IsDataTableName(listObject.Name))
                {
                    worksheet = candidate;
                    table = listObject;
                    return true;
                }
            }
        }

        return false;
    }

    private static Excel.Name? FindName(Excel.Names names, string nameText)
    {
        int count = names.Count;
        for (int index = 1; index <= count; index++)
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
        int count = sheets.Count;
        for (int index = 1; index <= count; index++)
        {
            if (
                GetSheetAt(sheets, index) is Excel.Worksheet worksheet
                && GanttWorkbookContract.IsConfigSheetName(worksheet.Name)
            )
            {
                return worksheet;
            }
        }

        return null;
    }

    internal virtual Excel.Worksheet GetSheetAt(Excel.Sheets sheets, int index) => sheets[index];

    /// <summary>
    /// Names of the tables on the configuration worksheet that are not contract
    /// catalogue tables.
    /// </summary>
    /// <param name="config">The configuration worksheet.</param>
    /// <returns>The foreign table names in collection order.</returns>
    private static List<string> FindForeignConfigTables(Excel.Worksheet config)
    {
        List<string> foreign = [];
        Excel.ListObjects tables = config.ListObjects;
        int count = tables.Count;
        for (int index = 1; index <= count; index++)
        {
            Excel.ListObject table = tables[index];
            string name = table.Name;
            // Case-insensitive because Excel treats table names that way, so a
            // catalogue renamed to different case is still the catalogue.
            if (
                !GanttCatalogues.ContractTableNames.Any(contract =>
                    string.Equals(contract, name, StringComparison.OrdinalIgnoreCase)
                )
            )
            {
                foreign.Add(name);
            }
        }

        return foreign;
    }
}
