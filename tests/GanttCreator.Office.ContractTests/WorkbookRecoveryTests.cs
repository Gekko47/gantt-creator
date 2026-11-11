using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Pins the read-only recovery decisions. Every case drives <see cref="ExcelWorkbookRecovery.Inspect"/>
/// through the adapter's seams, so a refusal is proved to happen before any host write.
/// </summary>
public class WorkbookRecoveryTests
{
    private sealed class SheetGraph
    {
        public SheetGraph(string name, params string[] tables)
        {
            Name = name;
            TableMocks =
            [
                .. tables.Select(tableName =>
                {
                    Mock<Excel.ListObject> table = new();
                    _ = table.SetupGet(t => t.Name).Returns(tableName);
                    return table;
                }),
            ];

            _ = ListObjects.SetupGet(l => l.Count).Returns(tables.Length);
            _ = Worksheet.SetupGet(w => w.ListObjects).Returns(ListObjects.Object);
            _ = Worksheet.SetupGet(w => w.Names).Returns(new Mock<Excel.Names>().Object);
            _ = Worksheet.SetupGet(w => w.Visible).Returns(() => Visibility);
        }

        public string Name { get; set; }

        public Excel.XlSheetVisibility Visibility { get; set; } =
            Excel.XlSheetVisibility.xlSheetVisible;

        public Mock<Excel.Worksheet> Worksheet { get; } = new();

        public Mock<Excel.ListObjects> ListObjects { get; } = new();

        public IReadOnlyList<Mock<Excel.ListObject>> TableMocks { get; }
    }

    private sealed class FixedInspector(ConfigSheetLocation location) : ExcelConfigSheetInspector
    {
        public override ConfigSheetLocation Locate(Excel.Sheets sheets) => location;
    }

    private sealed class TestableRecovery(
        object? application,
        IGanttTableReader reader,
        IWorksheetProtectionGuard guard,
        ExcelConfigSheetInspector inspector,
        IConfigCatalogueWriter? writer,
        params SheetGraph[] sheets)
        : ExcelWorkbookRecovery(application, reader, guard, null, writer, null, inspector)
    {
        private readonly List<SheetGraph> _sheets = [.. sheets];
        private readonly Dictionary<Excel.ListObject, string> _tableNames = [];
        private readonly Dictionary<(Excel.Worksheet Sheet, int Column), object?> _headers = [];
        private readonly Dictionary<Excel.Worksheet, int> _lastUsedRows = [];
        private readonly Dictionary<Excel.Worksheet, object?[][]> _bodies = [];
        private readonly Dictionary<Excel.Range, Excel.Worksheet> _ranges = [];
        private readonly Mock<Excel.Name> _anchorName = new();
        private int _nameCount;

        public TestableRecovery WithAlignedHeader(SheetGraph sheet)
        {
            IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
            for (int index = 0; index < schema.Count; index++)
            {
                _headers[(sheet.Worksheet.Object, index + 1)] = schema[index].Name;
            }

            return this;
        }

        public TestableRecovery WithBodyRow(SheetGraph sheet)
        {
            _lastUsedRows[sheet.Worksheet.Object] = GanttSheetLayout.HeaderRowIndex + 1;
            object?[] row = new object?[GanttTableSchema.Default.Columns.Count];
            row[0] = "G-1";
            _bodies[sheet.Worksheet.Object] = [row];
            return this;
        }

        public TestableRecovery WithPlotAnchor()
        {
            _nameCount = 1;
            return this;
        }

        public SheetGraph? CreatedSheet { get; private set; }

        public List<Excel.ListObject> CreatedTables { get; } = [];

        public int HeaderWrites { get; private set; }

        internal override int GetSheetCount(Excel.Sheets sheets) => _sheets.Count;

        internal override object GetSheetAt(Excel.Sheets sheets, int index) => _sheets[index - 1].Worksheet.Object;

        internal override string GetSheetName(Excel.Worksheet worksheet) => Graph(worksheet).Name;

        internal override void SetSheetName(Excel.Worksheet sheet, string name) => Graph(sheet).Name = name;

        internal override void SetSheetVisibility(
            Excel.Worksheet sheet,
            Excel.XlSheetVisibility visibility) => Graph(sheet).Visibility = visibility;

        internal override Excel.Worksheet AddSheetAfter(Excel.Sheets sheets, Excel.Worksheet? after)
        {
            SheetGraph created = new("Sheet" + (_sheets.Count + 1));
            _sheets.Add(created);
            CreatedSheet = created;
            return created.Worksheet.Object;
        }

        internal override string GetTableName(Excel.ListObject table) =>
            _tableNames.TryGetValue(table, out string? name) ? name : base.GetTableName(table);

        internal override void SetTableName(Excel.ListObject table, string name) => _tableNames[table] = name;

        public string RecordedTableName(Excel.ListObject table) =>
            _tableNames.TryGetValue(table, out string? name) ? name : string.Empty;

        internal override Excel.Range GetTableRange(
            Excel.Worksheet target,
            int headerRow,
            int bodyRowCount,
            int columnCount) => new Mock<Excel.Range>().Object;

        internal override void SetHeaderValue(Excel.Worksheet target, int columnIndex, string? value) =>
            HeaderWrites++;

        internal override Excel.ListObject AddTable(Excel.Worksheet target, Excel.Range range)
        {
            Mock<Excel.ListObject> table = new();
            Mock<Excel.ListColumns> columns = new();
            List<Mock<Excel.ListColumn>> columnMocks = [];
            for (int schemaIndex = 0; schemaIndex < GanttTableSchema.Default.Columns.Count; schemaIndex++)
            {
                Mock<Excel.ListColumn> column = new();
                Mock<Excel.Range> columnRange = new();
                Mock<Excel.Range> entireColumn = new();
                _ = column.SetupGet(c => c.Range).Returns(columnRange.Object);
                _ = columnRange.SetupGet(r => r.EntireColumn).Returns(entireColumn.Object);
                _ = entireColumn.SetupGet(r => r.Hidden).Returns(false);
                columnMocks.Add(column);
            }

            _ = columns.SetupGet(c => c.Count).Returns(columnMocks.Count);
            for (int index = 1; index <= columnMocks.Count; index++)
            {
                int captured = index;
                _ = columns.Setup(c => c[captured]).Returns(columnMocks[index - 1].Object);
            }

            _ = table.SetupGet(t => t.ListColumns).Returns(columns.Object);
            CreatedTables.Add(table.Object);
            return table.Object;
        }

        internal override int GetTableCount(Excel.ListObjects objects) => Graph(objects).TableMocks.Count;

        internal override Excel.ListObject GetTableAt(Excel.ListObjects objects, int index) =>
            Graph(objects).TableMocks[index - 1].Object;

        internal override object? GetHeaderValue(Excel.Worksheet target, int columnIndex) =>
            _headers.GetValueOrDefault((target, columnIndex));

        internal override int GetUsedRangeLastRow(Excel.Worksheet target) =>
            _lastUsedRows.TryGetValue(target, out int row) ? row : GanttSheetLayout.HeaderRowIndex;

        internal override Excel.Range GetRange(
            Excel.Worksheet target,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn)
        {
            Mock<Excel.Range> range = new();
            _ranges[range.Object] = target;
            return range.Object;
        }

        internal override object? GetRangeValues(Excel.Range range) => BodyMatrix(_bodies[_ranges[range]]);

        internal override int GetNameCount(Excel.Names names) => _nameCount;

        internal override Excel.Name GetNameAt(Excel.Names names, int index) => _anchorName.Object;

        internal override string GetNameText(Excel.Name name) =>
            GanttWorkbookContract.PlotAnchorDefinedName;

        private SheetGraph Graph(Excel.Worksheet worksheet) =>
            _sheets.First(s => ReferenceEquals(s.Worksheet.Object, worksheet));

        private SheetGraph Graph(Excel.ListObjects objects) =>
            _sheets.First(s => ReferenceEquals(s.ListObjects.Object, objects));

        private static Array BodyMatrix(object?[][] rows)
        {
            int columnCount = GanttTableSchema.Default.Columns.Count;
            Array matrix = Array.CreateInstance(
                typeof(object),
                [rows.Length, columnCount],
                [1, 1]);
            for (int row = 1; row <= rows.Length; row++)
            {
                for (int column = 1; column <= columnCount; column++)
                {
                    matrix.SetValue(rows[row - 1][column - 1], row, column);
                }
            }

            return matrix;
        }
    }

    private static (Mock<Excel.Workbook> Workbook, Mock<Excel.Sheets> Sheets) WorkbookGraph()
    {
        Mock<Excel.Workbook> workbook = new();
        Mock<Excel.Sheets> sheets = new();
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);
        return (workbook, sheets);
    }

    private static TestableRecovery Build(
        GanttTableReadOutcome read,
        ConfigSheetLocation config,
        Mock<IConfigCatalogueWriter>? writer = null,
        params SheetGraph[] sheets)
    {
        Mock<IGanttTableReader> reader = new();
        _ = reader.Setup(r => r.Read()).Returns(read);
        _ = reader.Setup(r => r.Read(It.IsAny<IReadOnlyList<int>>())).Returns(read);
        return BuildFrom(reader, config, writer, sheets);
    }

    private static TestableRecovery BuildFrom(
        Mock<IGanttTableReader> reader,
        ConfigSheetLocation config,
        Mock<IConfigCatalogueWriter>? writer,
        params SheetGraph[] sheets)
    {
        (Mock<Excel.Workbook> workbook, _) = WorkbookGraph();
        Mock<Excel.Application> application = new();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = application.SetupProperty(a => a.DisplayAlerts);

        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        _ = guard.Setup(g => g.QueryTarget(It.IsAny<object?>())).Returns(ProtectionGuardOutcome.NotProtected);

        return new TestableRecovery(
            application.Object,
            reader.Object,
            guard.Object,
            new FixedInspector(config),
            writer?.Object,
            sheets);
    }

    private static SheetGraph GanttSheet(string name = "Gantt Data", params string[] tables) =>
        new(name, tables);

    [Fact]
    public void Inspect_without_an_application_refuses_with_no_mutation()
    {
        Mock<IGanttTableReader> reader = new();
        Mock<IWorksheetProtectionGuard> guard = new();

        TestableRecovery recovery = new(
            null,
            reader.Object,
            guard.Object,
            new FixedInspector(new ConfigSheetLocation(ConfigSheetState.Missing, null)),
            null);

        WorkbookRecoveryInspection inspection = recovery.Inspect();

        Assert.Equal(WorkbookRecoveryRefusalReason.NoActiveWorkbook, inspection.Refusal);
        reader.Verify(r => r.Read(), Times.Never);
    }

    [Fact]
    public void Inspect_with_a_readable_table_and_configuration_reports_nothing_to_do()
    {
        TestableRecovery recovery = Build(
            GanttTableReadOutcome.Ok([], configurationPresent: true),
            new ConfigSheetLocation(ConfigSheetState.Present, GanttWorkbookContract.ConfigSheetName),
            sheets: [GanttSheet(tables: GanttTableSchema.TableName)]);

        WorkbookRecoveryInspection inspection = recovery.Inspect();

        Assert.False(inspection.NeedsRecovery);
        Assert.Null(inspection.Refusal);
        Assert.True(inspection.ConfigurationPresent);
    }

    [Theory]
    [InlineData(ConfigSheetState.NameCollision, WorkbookRecoveryRefusalReason.ConfigurationNameCollision)]
    [InlineData(ConfigSheetState.Ambiguous, WorkbookRecoveryRefusalReason.ConfigurationAmbiguous)]
    [InlineData(ConfigSheetState.AmbiguousCandidate, WorkbookRecoveryRefusalReason.ConfigurationCandidateAmbiguous)]
    [InlineData(ConfigSheetState.Unrecognised, WorkbookRecoveryRefusalReason.ConfigurationUnrecognised)]
    public void Inspect_refuses_a_colliding_configuration_before_reading_the_table(
        ConfigSheetState state,
        WorkbookRecoveryRefusalReason expected)
    {
        Mock<IGanttTableReader> reader = new();
        Mock<IWorksheetProtectionGuard> guard = new();
        Mock<Excel.Workbook> workbook = new();

        TestableRecovery recovery = new(
            ApplicationFor(workbook).Object,
            reader.Object,
            guard.Object,
            new FixedInspector(new ConfigSheetLocation(state, null)),
            null);

        WorkbookRecoveryInspection inspection = recovery.Inspect();

        Assert.Equal(expected, inspection.Refusal);
        reader.Verify(r => r.Read(), Times.Never);
    }

    [Fact]
    public void Inspect_asks_for_a_rename_when_the_only_candidate_is_a_renamed_ver_hidden_sheet()
    {
        TestableRecovery recovery = Build(
            GanttTableReadOutcome.Ok([], configurationPresent: false),
            new ConfigSheetLocation(ConfigSheetState.Renamed, "ConfigBackup"),
            sheets: [GanttSheet(tables: GanttTableSchema.TableName)]);

        WorkbookRecoveryInspection inspection = recovery.Inspect();

        Assert.True(inspection.NeedsRecovery);
        Assert.True(inspection.NeedsConfigurationRename);
        Assert.False(inspection.NeedsConfigurationRecovery);
        Assert.True(inspection.ConfigurationPresent);
        Assert.Null(inspection.Refusal);
    }

    [Fact]
    public void Inspect_asks_for_creation_when_no_configuration_candidate_exists()
    {
        TestableRecovery recovery = Build(
            GanttTableReadOutcome.Ok([], configurationPresent: false),
            new ConfigSheetLocation(ConfigSheetState.Missing, null),
            sheets: [GanttSheet(tables: GanttTableSchema.TableName)]);

        WorkbookRecoveryInspection inspection = recovery.Inspect();

        Assert.True(inspection.NeedsConfigurationRecovery);
        Assert.False(inspection.NeedsConfigurationRename);
        Assert.False(inspection.ConfigurationPresent);
    }

    [Fact]
    public void Inspect_offers_to_create_the_gantt_sheet_only_with_confirmation()
    {
        TestableRecovery recovery = Build(
            GanttTableReadOutcome.Refused(
                GanttTableReadRefusalReason.TableMissing,
                configurationPresent: true,
                ganttSheetPresent: false),
            new ConfigSheetLocation(ConfigSheetState.Present, GanttWorkbookContract.ConfigSheetName));

        WorkbookRecoveryInspection inspection = recovery.Inspect();

        Assert.True(inspection.NeedsGanttSheet);
        Assert.True(inspection.RequiresConfirmation);
        Assert.Contains("not restored", inspection.ConfirmationMessage, StringComparison.Ordinal);
        Assert.Null(inspection.Refusal);
    }

    [Fact]
    public void Inspect_refuses_two_standard_labelled_sheets()
    {
        TestableRecovery recovery = Build(
            GanttTableReadOutcome.Refused(
                GanttTableReadRefusalReason.TableMissing,
                configurationPresent: true,
                ganttSheetPresent: true),
            new ConfigSheetLocation(ConfigSheetState.Present, GanttWorkbookContract.ConfigSheetName),
            sheets: [GanttSheet(), GanttSheet()]);

        WorkbookRecoveryInspection inspection = recovery.Inspect();

        Assert.Equal(WorkbookRecoveryRefusalReason.TargetAmbiguous, inspection.Refusal);
    }

    [Fact]
    public void Inspect_plans_a_table_rename_when_the_sheet_proves_the_chart()
    {
        SheetGraph gantt = GanttSheet(tables: "Table1");
        TestableRecovery recovery = Build(
            GanttTableReadOutcome.Refused(
                GanttTableReadRefusalReason.TableMissing,
                configurationPresent: true,
                ganttSheetPresent: true),
            new ConfigSheetLocation(ConfigSheetState.Present, GanttWorkbookContract.ConfigSheetName),
            sheets: gantt);

        _ = recovery.WithAlignedHeader(gantt).WithPlotAnchor();
        WorkbookRecoveryInspection inspection = recovery.Inspect();

        Assert.True(inspection.NeedsTableRename);
        Assert.True(inspection.NeedsRecovery);
        Assert.Equal("Table1", inspection.TargetTableName);
        Assert.True(inspection.TablePresent);
        Assert.Null(inspection.Refusal);
    }

    [Fact]
    public void Inspect_refuses_a_foreign_table_when_the_sheet_is_not_proved_to_be_the_chart()
    {
        SheetGraph gantt = GanttSheet(tables: "Table1");
        TestableRecovery recovery = Build(
            GanttTableReadOutcome.Refused(
                GanttTableReadRefusalReason.TableMissing,
                configurationPresent: true,
                ganttSheetPresent: true),
            new ConfigSheetLocation(ConfigSheetState.Present, GanttWorkbookContract.ConfigSheetName),
            sheets: gantt);

        _ = recovery.WithAlignedHeader(gantt);
        WorkbookRecoveryInspection inspection = recovery.Inspect();

        Assert.Equal(WorkbookRecoveryRefusalReason.TargetHasForeignTable, inspection.Refusal);
        Assert.False(inspection.NeedsTableRename);
    }

    [Fact]
    public void Inspect_confirms_adoption_when_the_labelled_sheet_holds_body_rows()
    {
        SheetGraph gantt = GanttSheet();
        TestableRecovery recovery = Build(
            GanttTableReadOutcome.Refused(
                GanttTableReadRefusalReason.TableMissing,
                configurationPresent: true,
                ganttSheetPresent: true),
            new ConfigSheetLocation(ConfigSheetState.Present, GanttWorkbookContract.ConfigSheetName),
            sheets: gantt);

        _ = recovery.WithAlignedHeader(gantt).WithBodyRow(gantt);
        WorkbookRecoveryInspection inspection = recovery.Inspect();

        Assert.True(inspection.NeedsRecovery);
        Assert.True(inspection.RequiresConfirmation);
        Assert.Equal(1, inspection.BodyRowCount);
        Assert.Contains("adopt 1 existing row", inspection.ConfirmationMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void Recover_refuses_a_declined_confirmation_without_calling_the_catalogue_writer()
    {
        Mock<IConfigCatalogueWriter> writer = new(MockBehavior.Strict);
        TestableRecovery recovery = Build(
            GanttTableReadOutcome.Refused(
                GanttTableReadRefusalReason.TableMissing,
                configurationPresent: true,
                ganttSheetPresent: true),
            new ConfigSheetLocation(ConfigSheetState.Present, GanttWorkbookContract.ConfigSheetName),
            writer,
            GanttSheet());

        WorkbookRecoveryInspection inspection = new(
            TableReadable: false,
            TablePresent: false,
            ConfigurationPresent: true,
            NeedsRecovery: true,
            TargetSheetName: "Gantt Data",
            BodyRowCount: 1,
            HeaderPlan: null,
            StructurePlan: null,
            RequiresConfirmation: true,
            ConfirmationMessage: "adopt?",
            Refusal: null);

        WorkbookRecoveryOutcome outcome = recovery.Recover(inspection, confirmed: false);

        Assert.Equal(WorkbookRecoveryRefusalReason.ConfirmationDeclined, outcome.Refusal);
        Assert.False(outcome.Changed);
        writer.Verify(w => w.Write(), Times.Never);
    }

    [Fact]
    public void Recover_creates_the_configuration_sheet_and_materialises_it()
    {
        Mock<IConfigCatalogueWriter> writer = new(MockBehavior.Strict);
        _ = writer.Setup(w => w.Write()).Returns(ConfigWriteOutcome.Ok());
        Mock<IGanttTableReader> reader = new();
        _ = reader.Setup(r => r.Read())
            .Returns(GanttTableReadOutcome.Ok([], configurationPresent: false));
        SheetGraph gantt = GanttSheet(tables: GanttTableSchema.TableName);
        TestableRecovery recovery = BuildFrom(
            reader,
            new ConfigSheetLocation(ConfigSheetState.Missing, null),
            writer,
            gantt);

        WorkbookRecoveryInspection inspection = recovery.Inspect();
        WorkbookRecoveryOutcome outcome = recovery.Recover(inspection, confirmed: true);

        Assert.True(outcome.Changed);
        Assert.True(outcome.ConfigurationCreated);
        Assert.Null(outcome.Refusal);
        writer.Verify(w => w.Write(), Times.Once);
        Assert.NotNull(recovery.CreatedSheet);
        Assert.Equal(GanttWorkbookContract.ConfigSheetName, recovery.CreatedSheet!.Name);
        Assert.Equal(Excel.XlSheetVisibility.xlSheetVeryHidden, recovery.CreatedSheet.Visibility);
    }

    [Fact]
    public void Recover_rolls_back_a_created_configuration_sheet_when_the_write_fails()
    {
        Mock<IConfigCatalogueWriter> writer = new(MockBehavior.Strict);
        _ = writer.Setup(w => w.Write())
            .Returns(ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.HostRejected));
        Mock<IGanttTableReader> reader = new();
        _ = reader.Setup(r => r.Read())
            .Returns(GanttTableReadOutcome.Ok([], configurationPresent: false));
        TestableRecovery recovery = BuildFrom(
            reader,
            new ConfigSheetLocation(ConfigSheetState.Missing, null),
            writer,
            GanttSheet(tables: GanttTableSchema.TableName));

        WorkbookRecoveryInspection inspection = recovery.Inspect();
        WorkbookRecoveryOutcome outcome = recovery.Recover(inspection, confirmed: true);

        Assert.Equal(WorkbookRecoveryRefusalReason.ConfigurationWriteFailed, outcome.Refusal);
        Assert.False(outcome.Changed);
        recovery.CreatedSheet!.Worksheet.Verify(w => w.Delete(), Times.Once);
    }

    [Fact]
    public void Recover_renames_the_configuration_sheet_and_repairs_its_catalogues()
    {
        Mock<IConfigCatalogueWriter> writer = new(MockBehavior.Strict);
        _ = writer.Setup(w => w.Write()).Returns(ConfigWriteOutcome.Ok());
        Mock<IGanttTableReader> reader = new();
        _ = reader.Setup(r => r.Read())
            .Returns(GanttTableReadOutcome.Ok([], configurationPresent: false));
        SheetGraph backup = new("ConfigBackup");
        TestableRecovery recovery = BuildFrom(
            reader,
            new ConfigSheetLocation(ConfigSheetState.Renamed, "ConfigBackup"),
            writer,
            GanttSheet(tables: GanttTableSchema.TableName),
            backup);

        WorkbookRecoveryInspection inspection = recovery.Inspect();
        WorkbookRecoveryOutcome outcome = recovery.Recover(inspection, confirmed: true);

        Assert.True(outcome.ConfigurationRenamed);
        Assert.False(outcome.ConfigurationCreated);
        Assert.Null(outcome.Refusal);
        Assert.Equal(GanttWorkbookContract.ConfigSheetName, backup.Name);
        Assert.Equal(Excel.XlSheetVisibility.xlSheetVeryHidden, backup.Visibility);
        writer.Verify(w => w.Write(), Times.Once);
    }

    [Fact]
    public void Recover_restores_the_configuration_name_when_the_write_fails()
    {
        Mock<IConfigCatalogueWriter> writer = new(MockBehavior.Strict);
        _ = writer.Setup(w => w.Write())
            .Returns(ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.HostRejected));
        Mock<IGanttTableReader> reader = new();
        _ = reader.Setup(r => r.Read())
            .Returns(GanttTableReadOutcome.Ok([], configurationPresent: false));
        SheetGraph backup = new("ConfigBackup");
        TestableRecovery recovery = BuildFrom(
            reader,
            new ConfigSheetLocation(ConfigSheetState.Renamed, "ConfigBackup"),
            writer,
            GanttSheet(tables: GanttTableSchema.TableName),
            backup);

        WorkbookRecoveryInspection inspection = recovery.Inspect();
        WorkbookRecoveryOutcome outcome = recovery.Recover(inspection, confirmed: true);

        Assert.Equal(WorkbookRecoveryRefusalReason.ConfigurationWriteFailed, outcome.Refusal);
        Assert.Equal("ConfigBackup", backup.Name);
    }

    [Fact]
    public void Recover_renames_a_proven_table_and_verifies_the_read()
    {
        SheetGraph gantt = GanttSheet(tables: "Table1");
        Mock<IGanttTableReader> reader = new();
        _ = reader.SetupSequence(r => r.Read())
            .Returns(
                GanttTableReadOutcome.Refused(
                    GanttTableReadRefusalReason.TableMissing,
                    configurationPresent: true,
                    ganttSheetPresent: true))
            .Returns(GanttTableReadOutcome.Ok([], configurationPresent: true));
        TestableRecovery recovery = BuildFrom(
            reader,
            new ConfigSheetLocation(ConfigSheetState.Present, GanttWorkbookContract.ConfigSheetName),
            null,
            gantt);
        _ = recovery.WithAlignedHeader(gantt).WithPlotAnchor();

        WorkbookRecoveryInspection inspection = recovery.Inspect();
        WorkbookRecoveryOutcome outcome = recovery.Recover(inspection, confirmed: true);

        Assert.True(outcome.TableRenamed);
        Assert.True(outcome.Changed);
        Assert.Null(outcome.Refusal);
        Assert.Equal(
            GanttTableSchema.TableName,
            recovery.RecordedTableName(gantt.TableMocks[0].Object));
        reader.Verify(r => r.Read(), Times.Exactly(2));
    }

    [Fact]
    public void Recover_restores_the_table_name_when_the_verification_fails()
    {
        SheetGraph gantt = GanttSheet(tables: "Table1");
        Mock<IGanttTableReader> reader = new();
        _ = reader.SetupSequence(r => r.Read())
            .Returns(
                GanttTableReadOutcome.Refused(
                    GanttTableReadRefusalReason.TableMissing,
                    configurationPresent: true,
                    ganttSheetPresent: true))
            .Returns(
                GanttTableReadOutcome.Refused(
                    GanttTableReadRefusalReason.TableUnreadable,
                    configurationPresent: true,
                    tablePresent: true));
        TestableRecovery recovery = BuildFrom(
            reader,
            new ConfigSheetLocation(ConfigSheetState.Present, GanttWorkbookContract.ConfigSheetName),
            null,
            gantt);
        _ = recovery.WithAlignedHeader(gantt).WithPlotAnchor();

        WorkbookRecoveryInspection inspection = recovery.Inspect();
        WorkbookRecoveryOutcome outcome = recovery.Recover(inspection, confirmed: true);

        Assert.Equal(WorkbookRecoveryRefusalReason.VerificationFailed, outcome.Refusal);
        Assert.Equal("Table1", recovery.RecordedTableName(gantt.TableMocks[0].Object));
    }

    [Fact]
    public void Recover_creates_the_gantt_sheet_with_the_contract_header_and_table()
    {
        Mock<IGanttTableReader> reader = new();
        _ = reader.SetupSequence(r => r.Read())
            .Returns(
                GanttTableReadOutcome.Refused(
                    GanttTableReadRefusalReason.TableMissing,
                    configurationPresent: true,
                    ganttSheetPresent: false))
            .Returns(GanttTableReadOutcome.Ok([], configurationPresent: true));
        TestableRecovery recovery = BuildFrom(
            reader,
            new ConfigSheetLocation(ConfigSheetState.Present, GanttWorkbookContract.ConfigSheetName),
            null);

        WorkbookRecoveryInspection inspection = recovery.Inspect();
        WorkbookRecoveryOutcome outcome = recovery.Recover(inspection, confirmed: true);

        Assert.True(outcome.GanttSheetCreated);
        Assert.True(outcome.TableRebuilt);
        Assert.Null(outcome.Refusal);
        Assert.NotNull(recovery.CreatedSheet);
        Assert.Equal(GanttWorkbookContract.GanttSheetLabel, recovery.CreatedSheet!.Name);
        Assert.Equal(GanttTableSchema.Default.Columns.Count, recovery.HeaderWrites);
        Assert.Single(recovery.CreatedTables);
    }

    [Fact]
    public void Recover_deletes_the_created_gantt_sheet_when_the_verification_fails()
    {
        Mock<IGanttTableReader> reader = new();
        _ = reader.SetupSequence(r => r.Read())
            .Returns(
                GanttTableReadOutcome.Refused(
                    GanttTableReadRefusalReason.TableMissing,
                    configurationPresent: true,
                    ganttSheetPresent: false))
            .Returns(
                GanttTableReadOutcome.Refused(
                    GanttTableReadRefusalReason.TableUnreadable,
                    configurationPresent: true,
                    tablePresent: true));
        TestableRecovery recovery = BuildFrom(
            reader,
            new ConfigSheetLocation(ConfigSheetState.Present, GanttWorkbookContract.ConfigSheetName),
            null);

        WorkbookRecoveryInspection inspection = recovery.Inspect();
        WorkbookRecoveryOutcome outcome = recovery.Recover(inspection, confirmed: true);

        Assert.Equal(WorkbookRecoveryRefusalReason.VerificationFailed, outcome.Refusal);
        recovery.CreatedSheet!.Worksheet.Verify(w => w.Delete(), Times.Once);
    }

    private static Mock<Excel.Application> ApplicationFor(Mock<Excel.Workbook> workbook)
    {
        Mock<Excel.Application> application = new();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(new Mock<Excel.Sheets>().Object);
        return application;
    }
}
