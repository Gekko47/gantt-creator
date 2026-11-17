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
            // The chart is a labelled worksheet plus its table. A table on any
            // other worksheet is a state the product does not describe: the
            // reader follows the table by name, but recovery and the plot anchor
            // repair target the labelled sheet, so no path owns it. The
            // configuration sheet is excluded because it is already reported by
            // the stronger ScheduleDataOnConfig finding above — one table, one
            // statement about what is wrong with it.
            bool onLabelledSheet = GanttWorkbookContract.IsGanttSheetLabel(gantt!.Name);
            bool onConfigurationSheet = ReferenceEquals(gantt, config);
            if (!onLabelledSheet && !onConfigurationSheet)
            {
                findings.Add(
                    new ConfigIntegrityFinding(
                        ConfigIntegrityFindingKind.DataTableOnUnexpectedSheet,
                        $"The data table '{ganttTable!.Name}' sits on worksheet '{gantt.Name}', which does not carry the standard Gantt label '{GanttWorkbookContract.GanttSheetLabel}'. Move the table to the labelled worksheet; this state is reported and never repaired automatically.",
                        TableName: ganttTable.Name
                    )
                );
            }

            // Extra user columns beside the schema do not stop a render, but the
            // plot anchor, the title column and the export width are all derived
            // from the schema width, so a column the schema never defined silently
            // moves the chart. Reported and never deleted: the column is the
            // analyst's data, and removing it is their decision.
            List<string> unexpectedColumns = FindUnexpectedColumns(ganttTable!);
            if (unexpectedColumns.Count > 0)
            {
                findings.Add(
                    new ConfigIntegrityFinding(
                        ConfigIntegrityFindingKind.ExtraTableColumns,
                        $"The data table carries {unexpectedColumns.Count} column(s) the schema does not define: {string.Join(", ", unexpectedColumns)}. Remove or rename them; the add-in never deletes a column.",
                        TableName: ganttTable!.Name
                    )
                );
            }

            // A table that starts anywhere but the contract origin still renders,
            // because the reader follows it by name, but the title row, the period
            // band and the plot anchor are all stated against fixed rows and columns.
            // Reported and never moved: moving the table is a schema decision.
            Excel.Range tableRange = ganttTable!.Range;
            int headerRow = tableRange.Row;
            int firstColumn = tableRange.Column;
            if (!GanttSheetLayout.IsAtContractOrigin(headerRow, firstColumn))
            {
                findings.Add(
                    new ConfigIntegrityFinding(
                        ConfigIntegrityFindingKind.TableOffContractOrigin,
                        $"The data table starts at row {headerRow}, column {firstColumn}; the chart layout expects row {GanttSheetLayout.HeaderRowIndex}, column {GanttSheetLayout.TableOriginColumnIndex}. Rows inserted above the table, or a table moved sideways, shift the chart against its title and period band; this is reported and never moved automatically.",
                        TableName: ganttTable.Name
                    )
                );
            }

            // A sheet-scoped anchor means something only on the worksheet that is
            // meant to carry the chart. Checking it elsewhere reports a missing
            // anchor as if the chart were merely stale, and the anchor repair
            // writes to whichever sheet holds the table — a name on a sheet no
            // repair path owns.
            if (onLabelledSheet)
            {
                string expectedAnchor = GanttSheetLayout.BuildPlotAnchorRefersTo(gantt.Name);
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

                // A shape whose name is one this writer creates, but whose ownership
                // tag is absent or no longer matches that name, is the orphaned half
                // of the ownership contract: ListOwned cannot see it, so no refresh
                // reconciles or deletes it and every reconciliation scan treats it as
                // unowned. Reported and never deleted — a shape without a matching tag
                // cannot be proved owned, so deleting it because its name looks like
                // ours would breach the unowned-content invariant. The repair matrix
                // for ownership damage is R9.4's, and it consumes this finding rather
                // than re-detecting the same state.
                foreach (string orphanedShape in FindShapesThatCannotBeProvedOwned(gantt))
                {
                    findings.Add(
                        new ConfigIntegrityFinding(
                            ConfigIntegrityFindingKind.OwnedShapeWithoutValidTag,
                            $"The shape named '{orphanedShape}' on worksheet '{gantt.Name}' is named like one the add-in creates, but its ownership tag is missing or does not match that name. A shape without a matching tag cannot be proved owned, so the add-in never deletes it; this state is reported and never repaired automatically."
                        )
                    );
                }
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
    /// Names of the live table columns the schema does not define.
    /// </summary>
    /// <remarks>
    /// Answered only when every required schema column is present. A renamed or
    /// deleted column is structure damage with its own repair owner, and calling
    /// the rename an extra column would report one injury twice. The comparison
    /// is ordinal because the table reader resolves headers that way; any other
    /// rule would call a case-shifted header "extra" while the reader calls it
    /// missing.
    /// </remarks>
    /// <param name="table">The Gantt data table.</param>
    /// <returns>The unexpected column names in table order, or an empty list.</returns>
    private static List<string> FindUnexpectedColumns(Excel.ListObject table)
    {
        List<string> live = [];
        Excel.ListColumns columns = table.ListColumns;
        int count = columns.Count;
        for (int index = 1; index <= count; index++)
        {
            live.Add(columns[index].Name);
        }

        bool schemaComplete = GanttTableSchema
            .Default.Columns.Where(column => column.IsRequired)
            .All(column => live.Contains(column.Name, StringComparer.Ordinal));

        return schemaComplete
            ? [.. live.Where(name => !GanttTableSchema.Default.TryGetColumn(name, out _))]
            : [];
    }

    /// <summary>
    /// Names of the shapes on a worksheet that are named like add-in shapes but whose
    /// ownership tag cannot prove it.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Two tests, and both are needed. The name test is
    /// <see cref="ExcelShapeWriter.CouldBeOwnedName"/>, reused rather than restated so
    /// this scan and the ownership scan cannot disagree about which names this writer
    /// can create; it is a name test, not an ownership proof.
    /// </para>
    /// <para>
    /// The tag test reports only evidence of a <em>damaged</em> tag: no alternative
    /// text at all, or text that still carries the ownership prefix but is not the tag
    /// this name would produce (the "renamed" half, including a valid tag for some
    /// other identifier). A shape whose alternative text is a user's own description is
    /// left alone. "Could not be proved owned" is true of every drawing a user ever
    /// made, and a colon in a shape name must not turn one into a finding: this is a
    /// finding that refuses the repair plan, so a false positive is not free.
    /// </para>
    /// </remarks>
    /// <param name="worksheet">The worksheet that carries the chart.</param>
    /// <returns>The shape names in collection order, or an empty list.</returns>
    private static List<string> FindShapesThatCannotBeProvedOwned(Excel.Worksheet worksheet)
    {
        List<string> damaged = [];
        Excel.Shapes shapes = worksheet.Shapes;
        int count = shapes.Count;
        for (int index = 1; index <= count; index++)
        {
            Excel.Shape shape = shapes.Item(index);
            string name = shape.Name;
            if (string.IsNullOrWhiteSpace(name) || !ExcelShapeWriter.CouldBeOwnedName(name))
            {
                continue;
            }

            string? alternativeText = shape.AlternativeText;
            if (
                string.Equals(
                    alternativeText,
                    ShapeOwnershipTag.ForPrimitiveId(name),
                    StringComparison.Ordinal
                )
            )
            {
                continue;
            }

            if (
                string.IsNullOrWhiteSpace(alternativeText)
                || ShapeOwnershipTag.HasOwnershipPrefix(alternativeText)
            )
            {
                damaged.Add(name);
            }
        }

        return damaged;
    }

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
