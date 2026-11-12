using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using GanttCreator.Core.ConfigIntegrity;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Depth tests for <see cref="ExcelConfigIntegrityChecker"/>: the duplicate-sheet
/// scan, the workbook-name checks, and the catalogue-refusal mapping.
/// </summary>
/// <remarks>
/// <see cref="ConfigIntegrityCheckerTests"/> covers the missing-configuration-sheet
/// and catalogue-drift paths. This file exercises the branches that need a
/// resolvable Gantt table and a live defined-name collection, which is the half
/// the checker exists for.
/// </remarks>
public sealed class ConfigIntegrityCheckerDepthTests
{
    /// <summary>The checker with the sheet indexer substituted, so the scans run.</summary>
    private sealed class TestableChecker(
        Excel.Application application,
        IConfigCatalogueReader catalogueReader,
        IGanttTableReader tableReader,
        IReadOnlyList<Excel._Worksheet?> sheetObjects)
        : ExcelConfigIntegrityChecker(application, catalogueReader, tableReader)
    {
        internal override Excel.Worksheet GetSheetAt(Excel.Sheets sheets, int index) =>
            (Excel.Worksheet)sheetObjects[index - 1]!;
    }

    // The contract label, not an abbreviation: "Gantt" would model the very state
    // finding 5 reports — a data table on a worksheet the add-in never labelled.
    private const string GanttSheetName = GanttWorkbookContract.GanttSheetLabel;

    private static Mock<Excel.Name> DefinedName(string name, string refersTo)
    {
        Mock<Excel.Name> defined = new();
        _ = defined.SetupGet(n => n.Name).Returns(name);
        _ = defined.SetupGet(n => n.RefersTo).Returns(refersTo);
        return defined;
    }

    private static Mock<Excel.Names> NamesCollection(params Excel.Name[] entries)
    {
        Mock<Excel.Names> names = new();
        _ = names.SetupGet(n => n.Count).Returns(entries.Length);
        for (int index = 1; index <= entries.Length; index++)
        {
            Excel.Name entry = entries[index - 1];
            _ = names.Setup(n => n.Item(It.IsAny<object>())).Returns(entry);
        }

        return names;
    }

    /// <summary>A worksheet range address, which the host always supplies.</summary>
    private static Excel.Range RangeAt(int row, int column)
    {
        Mock<Excel.Range> range = new();
        _ = range.SetupGet(r => r.Row).Returns(row);
        _ = range.SetupGet(r => r.Column).Returns(column);
        return range.Object;
    }

    /// <summary>A worksheet's shape collection, modelling the host's always-present list.</summary>
    private static Excel.Shapes ShapesCollection(params (string Name, string AlternativeText)[] entries)
    {
        Mock<Excel.Shapes> shapes = new();
        _ = shapes.SetupGet(s => s.Count).Returns(entries.Length);
        for (int index = 1; index <= entries.Length; index++)
        {
            int captured = index;
            Mock<Excel.Shape> shape = new();
            _ = shape.SetupGet(s => s.Name).Returns(entries[captured - 1].Name);
            _ = shape.SetupGet(s => s.AlternativeText).Returns(entries[captured - 1].AlternativeText);
            _ = shapes.Setup(s => s.Item(captured)).Returns(shape.Object);
        }

        return shapes.Object;
    }

    /// <summary>A table's column collection, modelling the host's always-present list.</summary>
    private static Excel.ListColumns ColumnsCollection(params string[] names)
    {
        Mock<Excel.ListColumns> columns = new();
        _ = columns.SetupGet(c => c.Count).Returns(names.Length);
        for (int index = 1; index <= names.Length; index++)
        {
            int captured = index;
            Mock<Excel.ListColumn> column = new();
            _ = column.SetupGet(c => c.Name).Returns(names[captured - 1]);
            _ = columns.Setup(c => c[captured]).Returns(column.Object);
        }

        return columns.Object;
    }

    private static Excel.Worksheet ConfigSheet(
        Excel.XlSheetVisibility visibility = Excel.XlSheetVisibility.xlSheetVeryHidden,
        params string[] tableNames)
    {
        // An empty table collection (the default): the Gantt-table scan walks EVERY
        // sheet, so a configuration sheet reporting no ListObjects at all makes that
        // scan throw rather than skip it. A caller that supplies names is modelling
        // a configuration sheet that carries tables, which is the state the
        // ScheduleDataOnConfig finding exists for.
        Mock<Excel.ListObjects> tables = new();
        _ = tables.SetupGet(o => o.Count).Returns(tableNames.Length);
        for (int index = 1; index <= tableNames.Length; index++)
        {
            int captured = index;
            Mock<Excel.ListObject> table = new();
            _ = table.SetupGet(t => t.Name).Returns(tableNames[captured - 1]);
            // A real table always carries at least one column, and a table on the
            // configuration sheet is not the Gantt table: the column scan must never
            // see the schema in it.
            _ = table.SetupGet(t => t.ListColumns).Returns(ColumnsCollection("Value"));
            // Any address will do here: the sheet itself is what is wrong, and the
            // checker only asks where the table it found actually starts.
            _ = table.SetupGet(t => t.Range).Returns(RangeAt(1, 1));
            _ = tables.Setup(t => t[captured]).Returns(table.Object);
        }

        Mock<Excel.Worksheet> sheet = new();
        _ = sheet.SetupGet(s => s.Name).Returns(GanttWorkbookContract.ConfigSheetName);
        _ = sheet.SetupGet(s => s.Visible).Returns(visibility);
        _ = sheet.SetupGet(s => s.ListObjects).Returns(tables.Object);
        // Every real worksheet has a Names collection, and a configuration sheet
        // carrying a foreign table is also the sheet the Gantt-table scan finds
        // first, so leaving this null made the fixture less realistic than the host.
        _ = sheet.SetupGet(s => s.Names).Returns(NamesCollection().Object);
        return sheet.Object;
    }

    private static Excel.Worksheet GanttSheet(
        Excel.Names? names = null,
        string sheetName = GanttSheetName,
        string[]? columnNames = null,
        int headerRow = GanttSheetLayout.HeaderRowIndex,
        int firstColumn = GanttSheetLayout.TableOriginColumnIndex,
        Excel.Shapes? shapes = null)
    {
        // The live table's columns: the full schema by default, so a healthy
        // workbook has no extra-column finding. A caller supplies names to model a
        // table carrying a user column or missing a schema one.
        string[] columns =
            columnNames ?? [.. GanttTableSchema.Default.Columns.Select(column => column.Name)];

        Mock<Excel.ListObject> table = new();
        _ = table.SetupGet(t => t.Name).Returns(GanttTableSchema.TableName);
        _ = table.SetupGet(t => t.ListColumns).Returns(ColumnsCollection(columns));
        _ = table.SetupGet(t => t.Range).Returns(RangeAt(headerRow, firstColumn));
        Mock<Excel.ListObjects> tables = new();
        _ = tables.SetupGet(t => t.Count).Returns(1);
        _ = tables.Setup(t => t[It.IsAny<int>()]).Returns(table.Object);
        Mock<Excel.Worksheet> sheet = new();
        _ = sheet.SetupGet(s => s.Name).Returns(sheetName);
        _ = sheet.SetupGet(s => s.ListObjects).Returns(tables.Object);

        // An empty collection rather than a null one: the checker iterates the
        // sheet's defined names, and a null collection makes that read throw
        // instead of reporting "the anchor is missing".
        _ = sheet.SetupGet(s => s.Names).Returns(names ?? NamesCollection().Object);
        // The same for shapes: a healthy chart worksheet may carry none, and the
        // ownership scan walks the collection rather than testing it for null.
        _ = sheet.SetupGet(s => s.Shapes).Returns(shapes ?? ShapesCollection());
        return sheet.Object;
    }

    private static ConfigIntegrityCheckOutcome Check(
        IReadOnlyList<Excel._Worksheet?> sheets,
        Excel.Names? workbookNames = null,
        ConfigReadOutcome? catalogue = null)
    {
        Mock<Excel.Sheets> collection = new();
        _ = collection.SetupGet(s => s.Count).Returns(sheets.Count);
        Mock<Excel.Workbook> workbook = new();
        _ = workbook.SetupGet(w => w.Sheets).Returns(collection.Object);
        _ = workbook.SetupGet(w => w.Names).Returns(workbookNames ?? NamesCollection().Object);

        Mock<Excel.Application> application = new();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);

        Mock<IConfigCatalogueReader> reader = new();
        _ = reader.Setup(r => r.Read())
            .Returns(catalogue ?? ConfigReadOutcome.Refused(ConfigReadRefusalReason.TableMissing));
        Mock<IGanttTableReader> tableReader = new();
        _ = tableReader.Setup(r => r.Read())
            .Returns(GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.TableMissing));

        return new TestableChecker(application.Object, reader.Object, tableReader.Object, sheets).Check();
    }

    [Fact]
    public void A_shape_named_like_an_add_in_shape_without_its_tag_is_reported_and_refused()
    {
        // Finding 1 of R5.17f. ListOwned returns only shapes it can prove owned, so a
        // shape whose tag was stripped is invisible to every reconciliation path and
        // survives every refresh and Fix. The report is the whole of this slice: a
        // shape without a matching tag cannot be proved owned, so the add-in must not
        // delete it, and R9.4 owns the repair matrix.
        const string PrimitiveId = "1:ActivityBar";
        ConfigIntegrityCheckOutcome outcome = Check([
            ConfigSheet(),
            GanttSheet(shapes: ShapesCollection((PrimitiveId, string.Empty))),
        ]);

        ConfigIntegrityFinding finding = Assert.Single(
            outcome.Findings,
            candidate => candidate.Kind == ConfigIntegrityFindingKind.OwnedShapeWithoutValidTag
        );

        Assert.Contains(PrimitiveId, finding.Detail, StringComparison.Ordinal);
        Assert.Equal(ConfigRepairClassification.Refuse, ConfigIntegrityPlan.Classify(finding));
    }

    [Fact]
    public void A_shape_carrying_its_own_ownership_tag_is_not_reported()
    {
        // The healthy negative control: every shape a refresh writes carries the tag
        // recomputed from its own name, and reporting those would report a chart.
        const string PrimitiveId = "1:ActivityBar";
        ConfigIntegrityCheckOutcome outcome = Check([
            ConfigSheet(),
            GanttSheet(
                shapes: ShapesCollection((
                    PrimitiveId,
                    ShapeOwnershipTag.ForPrimitiveId(PrimitiveId)))),
        ]);

        Assert.DoesNotContain(
            outcome.Findings,
            candidate => candidate.Kind == ConfigIntegrityFindingKind.OwnedShapeWithoutValidTag
        );
    }

    [Fact]
    public void A_shape_carrying_another_shapes_tag_is_reported_as_renamed()
    {
        // The "renamed" half of finding 1, and the case a format check alone cannot
        // see: the tag is well formed but belongs to a different identifier, so
        // ownership is unproved for this name even though the prefix is right.
        const string PrimitiveId = "1:ActivityBar";
        ConfigIntegrityCheckOutcome outcome = Check([
            ConfigSheet(),
            GanttSheet(
                shapes: ShapesCollection((
                    PrimitiveId,
                    ShapeOwnershipTag.ForPrimitiveId("1:ActivityLabel")))),
        ]);

        Assert.Single(
            outcome.Findings,
            candidate => candidate.Kind == ConfigIntegrityFindingKind.OwnedShapeWithoutValidTag
        );
    }

    [Fact]
    public void A_user_shape_whose_name_contains_a_colon_is_not_reported()
    {
        // CouldBeOwnedName is a name test, not an ownership proof: a user drawing may
        // legitimately be called "Plan: Q3" and carry its own description. Reporting it
        // would refuse the whole repair plan for ordinary user content, so the scan
        // also requires the tag side to look damaged.
        ConfigIntegrityCheckOutcome outcome = Check([
            ConfigSheet(),
            GanttSheet(shapes: ShapesCollection(("Plan: Q3", "Draft outline"))),
        ]);

        Assert.DoesNotContain(
            outcome.Findings,
            candidate => candidate.Kind == ConfigIntegrityFindingKind.OwnedShapeWithoutValidTag
        );
    }

    [Fact]
    public void Schedule_data_on_the_configuration_sheet_is_reported_and_refused()
    {
        // The finding kind and its Refuse classification predate this producer, so
        // this is the test that makes the arm reachable: a table that is not one of
        // the five catalogues is schedule data (or the user's own) standing on the
        // sheet the add-in reserves for configuration.
        ConfigIntegrityCheckOutcome outcome = Check([
            ConfigSheet(tableNames: GanttTableSchema.TableName),
            GanttSheet(),
        ]);

        ConfigIntegrityFinding finding = Assert.Single(
            outcome.Findings,
            candidate => candidate.Kind == ConfigIntegrityFindingKind.ScheduleDataOnConfig);

        Assert.Equal(GanttTableSchema.TableName, finding.TableName);
        Assert.Equal(ConfigRepairClassification.Refuse, ConfigIntegrityPlan.Classify(finding));
    }

    [Fact]
    public void The_catalogue_tables_on_the_configuration_sheet_are_not_schedule_data()
    {
        // The negative control: a healthy configuration sheet carries exactly the
        // five contract tables, and reporting those would make the finding useless.
        ConfigIntegrityCheckOutcome outcome = Check([
            ConfigSheet(tableNames: [.. GanttCatalogues.ContractTableNames]),
            GanttSheet(),
        ]);

        Assert.DoesNotContain(
            outcome.Findings,
            candidate => candidate.Kind == ConfigIntegrityFindingKind.ScheduleDataOnConfig);
    }

    [Fact]
    public void One_foreign_table_among_the_catalogues_is_still_reported()
    {
        // A copied schedule lands beside the catalogues, not instead of them, so the
        // scan must not stop at the first contract name it recognises.
        string[] names = [.. GanttCatalogues.ContractTableNames, "ScratchPad"];
        ConfigIntegrityCheckOutcome outcome = Check([ConfigSheet(tableNames: names), GanttSheet()]);

        ConfigIntegrityFinding finding = Assert.Single(
            outcome.Findings,
            candidate => candidate.Kind == ConfigIntegrityFindingKind.ScheduleDataOnConfig);

        Assert.Equal("ScratchPad", finding.TableName);
    }

    [Fact]
    public void A_data_table_on_a_worksheet_without_the_gantt_label_is_reported_and_refused()
    {
        // Finding 5: the reader follows the table by name, so the chart renders on
        // a worksheet the add-in never labelled — but recovery and the anchor
        // repair both target the labelled sheet, so no repair path owns this state.
        ConfigIntegrityCheckOutcome outcome = Check([
            ConfigSheet(),
            GanttSheet(sheetName: "Scratch"),
        ]);

        ConfigIntegrityFinding finding = Assert.Single(
            outcome.Findings,
            candidate => candidate.Kind == ConfigIntegrityFindingKind.DataTableOnUnexpectedSheet);

        Assert.Equal(GanttTableSchema.TableName, finding.TableName);
        Assert.Contains("Scratch", finding.Detail, StringComparison.Ordinal);
        Assert.Equal(ConfigRepairClassification.Refuse, ConfigIntegrityPlan.Classify(finding));
        // The sheet-scoped anchor check must not run here: it would report a
        // missing anchor for a table whose real damage is its location.
        Assert.DoesNotContain(
            outcome.Findings,
            candidate => candidate.Kind == ConfigIntegrityFindingKind.PlotAnchorDisagreement);
    }

    [Fact]
    public void A_data_table_on_the_labelled_gantt_worksheet_is_not_mislocated()
    {
        // The negative control: the healthy workbook carries the table on the
        // worksheet the add-in labelled, so the location finding must not fire.
        ConfigIntegrityCheckOutcome outcome = Check([ConfigSheet(), GanttSheet()]);

        Assert.DoesNotContain(
            outcome.Findings,
            candidate => candidate.Kind == ConfigIntegrityFindingKind.DataTableOnUnexpectedSheet);
    }

    [Fact]
    public void A_data_table_on_the_configuration_sheet_is_not_also_mislocated()
    {
        // Schedule data on the configuration worksheet is the stronger statement,
        // and one table must not produce two findings. The same table also used to
        // be the one the anchor check ran against, reporting a disagreement for a
        // chart whose real damage is its location.
        ConfigIntegrityCheckOutcome outcome = Check([
            ConfigSheet(tableNames: GanttTableSchema.TableName),
            GanttSheet(),
        ]);

        Assert.DoesNotContain(
            outcome.Findings,
            candidate => candidate.Kind == ConfigIntegrityFindingKind.DataTableOnUnexpectedSheet);
        Assert.DoesNotContain(
            outcome.Findings,
            candidate => candidate.Kind == ConfigIntegrityFindingKind.PlotAnchorDisagreement);
        _ = Assert.Single(
            outcome.Findings,
            candidate => candidate.Kind == ConfigIntegrityFindingKind.ScheduleDataOnConfig);
    }

    [Fact]
    public void An_extra_user_column_beside_the_schema_is_reported_and_refused()
    {
        // Finding 3: the extra column does not stop a render, but the plot anchor,
        // the title column and the export width are derived from the schema width,
        // so the chart silently moves. The column is the analyst's data: report,
        // never delete.
        string[] columns = [.. GanttTableSchema.Default.Columns.Select(column => column.Name), "Notes"];

        ConfigIntegrityCheckOutcome outcome = Check([ConfigSheet(), GanttSheet(columnNames: columns)]);

        ConfigIntegrityFinding finding = Assert.Single(
            outcome.Findings,
            candidate => candidate.Kind == ConfigIntegrityFindingKind.ExtraTableColumns);

        Assert.Equal(GanttTableSchema.TableName, finding.TableName);
        Assert.Contains("Notes", finding.Detail, StringComparison.Ordinal);
        Assert.Equal(ConfigRepairClassification.Refuse, ConfigIntegrityPlan.Classify(finding));
    }

    [Fact]
    public void The_schema_columns_alone_are_not_extra_columns()
    {
        // The healthy negative control: the fixture table carries exactly the
        // schema, optional columns included.
        ConfigIntegrityCheckOutcome outcome = Check([ConfigSheet(), GanttSheet()]);

        Assert.DoesNotContain(
            outcome.Findings,
            candidate => candidate.Kind == ConfigIntegrityFindingKind.ExtraTableColumns);
    }

    [Fact]
    public void A_renamed_schema_column_is_structure_damage_not_an_extra_column()
    {
        // A rename looks like one missing column plus one extra column. R5.17a owns
        // the rename, so this finding stays silent until every required column is
        // back: one injury, one owner.
        string[] columns =
        [
            .. GanttTableSchema
                .Default.Columns.Where(column => column.Name != "Type")
                .Select(column => column.Name),
            "Kind",
        ];

        ConfigIntegrityCheckOutcome outcome = Check([ConfigSheet(), GanttSheet(columnNames: columns)]);

        Assert.DoesNotContain(
            outcome.Findings,
            candidate => candidate.Kind == ConfigIntegrityFindingKind.ExtraTableColumns);
    }

    [Fact]
    public void A_table_below_the_contract_header_row_is_reported_and_refused()
    {
        // Rows inserted above the table push it down; the title row, the period band
        // and the plot anchor all stay where the layout states them, so the chart
        // draws against the wrong rows.
        int shifted = GanttSheetLayout.HeaderRowIndex + 2;

        ConfigIntegrityCheckOutcome outcome = Check([ConfigSheet(), GanttSheet(headerRow: shifted)]);

        ConfigIntegrityFinding finding = Assert.Single(
            outcome.Findings,
            candidate => candidate.Kind == ConfigIntegrityFindingKind.TableOffContractOrigin);

        Assert.Equal(GanttTableSchema.TableName, finding.TableName);
        Assert.Contains($"row {shifted}", finding.Detail, StringComparison.Ordinal);
        Assert.Equal(ConfigRepairClassification.Refuse, ConfigIntegrityPlan.Classify(finding));
    }

    [Fact]
    public void A_table_moved_off_the_contract_column_is_reported()
    {
        // A sideways move is the same injury on the other axis, so the one finding
        // fires rather than a second kind existing for columns.
        ConfigIntegrityCheckOutcome outcome = Check([
            ConfigSheet(),
            GanttSheet(firstColumn: GanttSheetLayout.TableOriginColumnIndex + 3),
        ]);

        ConfigIntegrityFinding finding = Assert.Single(
            outcome.Findings,
            candidate => candidate.Kind == ConfigIntegrityFindingKind.TableOffContractOrigin);

        Assert.Contains("column 4", finding.Detail, StringComparison.Ordinal);
        Assert.Equal(ConfigRepairClassification.Refuse, ConfigIntegrityPlan.Classify(finding));
    }

    [Fact]
    public void A_table_at_the_contract_origin_is_not_reported()
    {
        // The healthy negative control: the default fixture table starts exactly where
        // the layout states.
        ConfigIntegrityCheckOutcome outcome = Check([ConfigSheet(), GanttSheet()]);

        Assert.DoesNotContain(
            outcome.Findings,
            candidate => candidate.Kind == ConfigIntegrityFindingKind.TableOffContractOrigin);
    }

    [Fact]
    public void A_second_worksheet_carrying_the_configuration_name_is_reported()
    {
        // Two distinct worksheets with the name: the second is a helper sheet the
        // contract forbids, so it must be reported rather than silently used.
        ConfigIntegrityCheckOutcome outcome = Check([ConfigSheet(), ConfigSheet(), GanttSheet(names: null)]);

        Assert.Contains(
            outcome.Findings,
            finding => finding.Kind == ConfigIntegrityFindingKind.SecondHelperSheet);
    }

    [Fact]
    public void The_same_worksheet_resolved_twice_is_not_a_second_helper_sheet()
    {
        // Identity, not name: the scan must compare instances, or resolving one
        // sheet at two indices would report a duplicate that does not exist.
        Excel.Worksheet config = ConfigSheet();
        ConfigIntegrityCheckOutcome outcome = Check([config, config]);

        Assert.DoesNotContain(
            outcome.Findings,
            finding => finding.Kind == ConfigIntegrityFindingKind.SecondHelperSheet);
    }

    [Fact]
    public void A_missing_plot_anchor_defined_name_is_reported_as_a_disagreement()
    {
        Excel.Worksheet gantt = GanttSheet(names: NamesCollection().Object);

        ConfigIntegrityCheckOutcome outcome = Check(
            [ConfigSheet(), gantt],
            workbookNames: NamesCollection().Object);

        Assert.Contains(
            outcome.Findings,
            finding => finding.Kind == ConfigIntegrityFindingKind.PlotAnchorDisagreement);
    }

    [Fact]
    public void An_anchor_pointing_at_another_sheet_is_reported_as_a_disagreement()
    {
        Excel.Worksheet gantt = GanttSheet(
            names: NamesCollection(
                DefinedName(GanttWorkbookContract.PlotAnchorDefinedName, "=SomeOtherSheet!$B$2").Object).Object);

        ConfigIntegrityCheckOutcome outcome = Check(
            [ConfigSheet(), gantt],
            workbookNames: NamesCollection(
                DefinedName(GanttWorkbookContract.TypeOptionsDefinedName, "={1,2}").Object).Object);

        Assert.Contains(
            outcome.Findings,
            finding => finding.Kind == ConfigIntegrityFindingKind.PlotAnchorDisagreement);
    }

    [Fact]
    public void A_matching_anchor_and_a_present_type_options_name_agree()
    {
        string expected = GanttSheetLayout.BuildPlotAnchorRefersTo(GanttSheetName);
        Excel.Worksheet gantt = GanttSheet(
            names: NamesCollection(
                DefinedName(GanttWorkbookContract.PlotAnchorDefinedName, expected).Object).Object);

        ConfigIntegrityCheckOutcome outcome = Check(
            [ConfigSheet(), gantt],
            workbookNames: NamesCollection(
                DefinedName(GanttWorkbookContract.TypeOptionsDefinedName, "={1,2}").Object).Object);

        Assert.DoesNotContain(
            outcome.Findings,
            finding => finding.Kind == ConfigIntegrityFindingKind.PlotAnchorDisagreement);
        Assert.DoesNotContain(
            outcome.Findings,
            finding => finding.Kind == ConfigIntegrityFindingKind.TypeOptionsMissing);
    }

    [Fact]
    public void A_missing_type_options_defined_name_is_reported()
    {
        string expected = GanttSheetLayout.BuildPlotAnchorRefersTo(GanttSheetName);
        Excel.Worksheet gantt = GanttSheet(
            names: NamesCollection(
                DefinedName(GanttWorkbookContract.PlotAnchorDefinedName, expected).Object).Object);

        ConfigIntegrityCheckOutcome outcome = Check(
            [ConfigSheet(), gantt],
            workbookNames: NamesCollection().Object);

        Assert.Contains(
            outcome.Findings,
            finding => finding.Kind == ConfigIntegrityFindingKind.TypeOptionsMissing);
    }

    /// <summary>
    /// Every catalogue refusal maps to its own finding kind. The mapping is the
    /// checker's whole testimony about why the catalogue is unusable, so a default
    /// arm swallowing a new reason would hide it.
    /// </summary>
    public static TheoryData<ConfigReadRefusalReason, ConfigIntegrityFindingKind> RefusalMappings
    {
        get
        {
            TheoryData<ConfigReadRefusalReason, ConfigIntegrityFindingKind> data = [];
            data.Add(ConfigReadRefusalReason.NoActiveWorkbook, ConfigIntegrityFindingKind.Unknown);
            data.Add(ConfigReadRefusalReason.ConfigSheetMissing, ConfigIntegrityFindingKind.ConfigSheetMissing);
            data.Add(ConfigReadRefusalReason.CatalogueHashMismatch, ConfigIntegrityFindingKind.CatalogueHashMismatch);
            data.Add(ConfigReadRefusalReason.TableMissing, ConfigIntegrityFindingKind.CatalogueTableMissing);
            data.Add(ConfigReadRefusalReason.HeaderMismatch, ConfigIntegrityFindingKind.CatalogueTableCorrupt);
            data.Add(ConfigReadRefusalReason.RowCountMismatch, ConfigIntegrityFindingKind.CatalogueTableCorrupt);
            data.Add(ConfigReadRefusalReason.CatalogueMismatch, ConfigIntegrityFindingKind.CatalogueTableCorrupt);
            data.Add(ConfigReadRefusalReason.ValueOutOfRange, ConfigIntegrityFindingKind.CatalogueTableCorrupt);
            data.Add(ConfigReadRefusalReason.BadColourFormat, ConfigIntegrityFindingKind.CatalogueTableCorrupt);
            return data;
        }
    }

    [Theory]
    [MemberData(nameof(RefusalMappings))]
    public void Each_catalogue_refusal_maps_to_its_own_finding_kind(
        ConfigReadRefusalReason refusal,
        ConfigIntegrityFindingKind expected)
    {
        ConfigIntegrityCheckOutcome outcome = Check(
            [ConfigSheet()],
            catalogue: ConfigReadOutcome.Refused(refusal));

        Assert.Contains(
            outcome.Findings,
            finding => finding.Kind == expected && finding.Detail.Contains(refusal.ToString(), StringComparison.Ordinal));
    }
}
