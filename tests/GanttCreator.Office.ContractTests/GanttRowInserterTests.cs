using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>Contract tests for the R2.8 row insertion adapter.</summary>
public class GanttRowInserterTests
{
    private sealed class TestableInserter : ExcelGanttRowInserter
    {
        private readonly Func<Excel.Sheets, int, Excel.Worksheet> _sheetAt;
        private readonly Func<Excel.ListObjects, int, Excel.ListObject> _tableAt;
        private readonly Func<Excel.ListColumns, int, Excel.ListColumn> _columnAt;
        private readonly Func<Excel.ListObject, Excel.ListRows> _rowsAt;
        private readonly Func<Excel.ListRows, Excel.ListRow> _addRow;
        private readonly Func<Excel.ListRow, int> _rowIndex;
        private readonly Func<Excel.ListRow, Excel.Range> _rowRange;

        public TestableInserter(
            Excel.Application application,
            IWorksheetProtectionGuard guard,
            Func<Excel.Sheets, int, Excel.Worksheet> sheetAt,
            Func<Excel.ListObjects, int, Excel.ListObject> tableAt,
            Func<Excel.ListColumns, int, Excel.ListColumn> columnAt,
            Func<Excel.ListObject, Excel.ListRows> rowsAt,
            Func<Excel.ListRows, Excel.ListRow> addRow,
            Func<Excel.ListRow, int> rowIndex,
            Func<Excel.ListRow, Excel.Range> rowRange)
            : base(application, guard)
        {
            _sheetAt = sheetAt;
            _tableAt = tableAt;
            _columnAt = columnAt;
            _rowsAt = rowsAt;
            _addRow = addRow;
            _rowIndex = rowIndex;
            _rowRange = rowRange;
        }

        internal override Excel.Worksheet GetSheetAt(Excel.Sheets sheets, int index) => _sheetAt(sheets, index);
        internal override Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index) => _tableAt(listObjects, index);
        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index) => _columnAt(columns, index);
        internal override Excel.ListRows GetListRows(Excel.ListObject table) => _rowsAt(table);
        internal override Excel.ListRow AddRow(Excel.ListRows rows) => _addRow(rows);
        internal override int GetRowIndex(Excel.ListRow row) => _rowIndex(row);
        internal override Excel.Range GetRowRange(Excel.ListRow row) => _rowRange(row);
    }

    private sealed class Graph
    {
        public Mock<Excel.Application> Application { get; } = new();
        public Mock<Excel.Workbook> Workbook { get; } = new();
        public Mock<Excel.Sheets> Sheets { get; } = new();
        public Mock<Excel.Worksheet> Worksheet { get; } = new();
        public Mock<Excel.ListObjects> ListObjects { get; } = new();
        public Mock<Excel.ListObject> Table { get; } = new();
        public Mock<Excel.ListColumns> ListColumns { get; } = new();
        public Mock<Excel.ListColumn>[] Columns { get; } = Enumerable.Range(0, 14).Select(_ => new Mock<Excel.ListColumn>()).ToArray();
        public Mock<Excel.ListRows> ListRows { get; } = new();
        public Mock<Excel.ListRow> NewRow { get; } = new();
        public Mock<Excel.Range> RowRange { get; } = new();
        public object? WrittenValue { get; private set; }
        public int AddRowCalls { get; private set; }

        public Graph()
        {
            _ = Application.SetupGet(a => a.ActiveWorkbook).Returns(Workbook.Object);
            _ = Workbook.SetupGet(w => w.Sheets).Returns(Sheets.Object);
            _ = Sheets.SetupGet(s => s.Count).Returns(1);
            _ = Worksheet.SetupGet(w => w.ListObjects).Returns(ListObjects.Object);
            _ = ListObjects.SetupGet(l => l.Count).Returns(1);
            _ = Table.SetupGet(t => t.Name).Returns(GanttTableSchema.TableName);
            _ = Table.SetupGet(t => t.ListColumns).Returns(ListColumns.Object);
            _ = ListColumns.SetupGet(c => c.Count).Returns(Columns.Length);
            for (var index = 0; index < GanttTableSchema.Default.Columns.Count; index++)
            {
                Columns[index].SetupGet(c => c.Name).Returns(GanttTableSchema.Default.Columns[index].Name);
            }
            _ = Table.SetupGet(t => t.ListRows).Returns(ListRows.Object);
            _ = ListRows.SetupGet(r => r.Count).Returns(1);
            _ = NewRow.SetupGet(r => r.Index).Returns(2);
            _ = NewRow.SetupGet(r => r.Range).Returns(RowRange.Object);
            _ = RowRange.SetupSet(r => r.Value2 = It.IsAny<object>())
                .Callback<object>(value => WrittenValue = value);
            _ = ListRows.Setup(r => r.Add(Type.Missing)).Callback(() => AddRowCalls++).Returns(NewRow.Object);
            _ = ListRows.Setup(r => r.Add(It.IsAny<object>())).Callback(() => AddRowCalls++).Returns(NewRow.Object);
        }

        public TestableInserter Build(IWorksheetProtectionGuard guard) => new(
            Application.Object,
            guard,
            (_, _) => Worksheet.Object,
            (listObjects, _) =>
            {
                Assert.Same(ListObjects.Object, listObjects);
                return Table.Object;
            },
            (_, index) => Columns[index - 1].Object,
            table =>
            {
                Assert.Same(Table.Object, table);
                return ListRows.Object;
            },
            rows =>
            {
                Assert.Same(ListRows.Object, rows);
                AddRowCalls++;
                return NewRow.Object;
            },
            row =>
            {
                Assert.Same(NewRow.Object, row);
                return 2;
            },
            row =>
            {
                Assert.Same(NewRow.Object, row);
                return RowRange.Object;
            });
    }

    [Fact]
    public void Insert_appends_one_bulk_row_with_exact_scaffold_values()
    {
        var graph = new Graph();
        var guard = new Mock<IWorksheetProtectionGuard>();
        guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        GanttRowId id = GanttRowId.Parse("G-0123456789abcdef0123456789abcdef");

        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            () => id);

        Assert.True(outcome.Succeeded);
        Assert.Equal(2, outcome.BodyIndex);
        Assert.Equal(1, graph.AddRowCalls);
        guard.Verify(g => g.Query(), Times.Once);
        var matrix = Assert.IsType<object[,]>(graph.WrittenValue);
        Assert.Equal(1, matrix.GetLength(0));
        Assert.Equal(14, matrix.GetLength(1));
        Assert.Equal(id.Value, matrix[0, 0]);
        Assert.Equal("As-Planned Activity", matrix[0, 3]);
        Assert.Equal("AsPlannedActivity", matrix[0, 8]);
        Assert.Equal(string.Empty, matrix[0, 13]);
    }

    [Fact]
    public void Insert_maps_columns_by_header_name_not_position()
    {
        var graph = new Graph();
        var guard = new Mock<IWorksheetProtectionGuard>();
        guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var names = GanttTableSchema.Default.Columns.Select((column, index) => (column, index)).ToArray();
        for (var index = 0; index < names.Length; index++)
        {
            graph.Columns[index].SetupGet(c => c.Name).Returns(names[index].column.Name);
        }
        // Reverse the table's physical order; the logical scaffold still lands
        // in the correct header-named cells.
        for (var index = 0; index < names.Length; index++)
        {
            var schemaIndex = names.Length - index - 1;
            graph.Columns[index].SetupGet(c => c.Name).Returns(names[schemaIndex].column.Name);
        }

        GanttRowId id = GanttRowId.Parse("G-0123456789abcdef0123456789abcdef");
        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(GanttEntityType.Delineator, () => id);

        Assert.True(outcome.Succeeded);
        var matrix = Assert.IsType<object[,]>(graph.WrittenValue);
        // Physical column 0 is logical SortOrder, physical column 13 is Id.
        Assert.Equal(string.Empty, matrix[0, 0]);
        Assert.Equal(id.Value, matrix[0, 13]);
        Assert.Equal("Delineator", matrix[0, 10]);
    }

    [Theory]
    [InlineData(ProtectionGuardOutcome.SheetProtected, GanttRowInsertRefusalReason.TargetProtected)]
    [InlineData(ProtectionGuardOutcome.WorkbookStructureProtected, GanttRowInsertRefusalReason.TargetProtected)]
    [InlineData(ProtectionGuardOutcome.NoActiveWorkbook, GanttRowInsertRefusalReason.NoActiveWorkbook)]
    public void Insert_refuses_before_any_workbook_mutation(
        ProtectionGuardOutcome guardOutcome,
        GanttRowInsertRefusalReason expected)
    {
        var graph = new Graph();
        var guard = new Mock<IWorksheetProtectionGuard>();
        guard.Setup(g => g.Query()).Returns(guardOutcome);

        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            GanttRowId.New);

        Assert.False(outcome.Succeeded);
        Assert.Equal(expected, outcome.Refusal);
        Assert.Equal(0, graph.AddRowCalls);
        Assert.Null(graph.WrittenValue);
        guard.Verify(g => g.Query(), Times.Once);
        graph.Workbook.VerifyGet(w => w.Sheets, Times.Never);
    }

    [Fact]
    public void Insert_returns_table_missing_when_no_table_is_found()
    {
        var graph = new Graph();
        graph.ListObjects.SetupGet(l => l.Count).Returns(0);
        var guard = new Mock<IWorksheetProtectionGuard>();
        guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);

        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            GanttRowId.New);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttRowInsertRefusalReason.TableMissing, outcome.Refusal);
        Assert.Equal(0, graph.AddRowCalls);
        Assert.Null(graph.WrittenValue);
    }
}
