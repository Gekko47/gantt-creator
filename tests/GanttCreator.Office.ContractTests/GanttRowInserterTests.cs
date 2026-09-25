using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>Contract tests for the R2.8 row insertion adapter.</summary>
public class GanttRowInserterTests
{
    private sealed class StubTypeOptionsMaterialiser : ITypeOptionsMaterialiser
    {
        public TypeOptionsMaterialiseOutcome Materialise() => TypeOptionsMaterialiseOutcome.Ok();
    }

    private sealed class TestableInserter(
        Excel.Application application,
        IWorksheetProtectionGuard guard,
        Func<Excel.Sheets, int, Excel.Worksheet> sheetAt,
        Func<Excel.ListObjects, int, Excel.ListObject> tableAt,
        Func<Excel.ListColumns, int, Excel.ListColumn> columnAt,
        Func<Excel.ListObject, Excel.ListRows> rowsAt,
        Func<Excel.ListObject, Excel.Range> tableRangeAt,
        Func<Excel.Range, Excel.Range> rangeRowsAt,
        Func<Excel.Range, int> rangeRowCountAt,
        Func<Excel.Range, int, Excel.Range> rangeAt,
        Func<Excel.Range, object?> rangeValueAt,
        Action<Excel.Range> clearRange,
        Func<Excel.ListRows, Excel.ListRow> addRow,
        Func<Excel.ListRows, int, Excel.ListRow> addRowAtPosition,
        Func<Excel.Application, Excel.Range?> activeCell,
        Func<Excel.ListObject, bool> tableActive,
        Func<Excel.Range, int> rangeRow,
        Func<Excel.ListRow, int> rowIndex,
        Func<Excel.ListRow, Excel.Range> rowRange,
        ITypeOptionsMaterialiser? typeOptionsMaterialiser = null) : ExcelGanttRowInserter(application, guard, typeOptionsMaterialiser)
    {
        private readonly Func<Excel.Sheets, int, Excel.Worksheet> _sheetAt = sheetAt;
        private readonly Func<Excel.ListObjects, int, Excel.ListObject> _tableAt = tableAt;
        private readonly Func<Excel.ListColumns, int, Excel.ListColumn> _columnAt = columnAt;
        private readonly Func<Excel.ListObject, Excel.ListRows> _rowsAt = rowsAt;
        private readonly Func<Excel.ListObject, Excel.Range> _tableRangeAt = tableRangeAt;
        private readonly Func<Excel.Range, Excel.Range> _rangeRowsAt = rangeRowsAt;
        private readonly Func<Excel.Range, int> _rangeRowCountAt = rangeRowCountAt;
        private readonly Func<Excel.Range, int, Excel.Range> _rangeAt = rangeAt;
        private readonly Func<Excel.Range, object?> _rangeValueAt = rangeValueAt;
        private readonly Action<Excel.Range> _clearRange = clearRange;
        private readonly Func<Excel.ListRows, Excel.ListRow> _addRow = addRow;
        private readonly Func<Excel.ListRows, int, Excel.ListRow> _addRowAtPosition = addRowAtPosition;
        private readonly Func<Excel.Application, Excel.Range?> _activeCell = activeCell;
        private readonly Func<Excel.ListObject, bool> _tableActive = tableActive;
        private readonly Func<Excel.Range, int> _rangeRow = rangeRow;
        private readonly Func<Excel.ListRow, int> _rowIndex = rowIndex;
        private readonly Func<Excel.ListRow, Excel.Range> _rowRange = rowRange;

        internal override Excel.Worksheet GetSheetAt(Excel.Sheets sheets, int index) => _sheetAt(sheets, index);
        internal override Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index) => _tableAt(listObjects, index);
        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index) => _columnAt(columns, index);
        internal override Excel.ListRows GetListRows(Excel.ListObject table) => _rowsAt(table);
        internal override Excel.Range GetTableRange(Excel.ListObject table) => _tableRangeAt(table);
        internal override Excel.Range GetRangeRows(Excel.Range range) => _rangeRowsAt(range);
        internal override int GetRangeRowCount(Excel.Range rows) => _rangeRowCountAt(rows);
        internal override Excel.Range GetRangeAt(Excel.Range rows, int index) => _rangeAt(rows, index);
        internal override object? GetRangeValue2(Excel.Range range) => _rangeValueAt(range);
        internal override void ClearRange(Excel.Range range) => _clearRange(range);
        internal override Excel.ListRow AddRow(Excel.ListRows rows) => _addRow(rows);
        internal override Excel.ListRow AddRowAtPosition(Excel.ListRows rows, int position) => _addRowAtPosition(rows, position);
        internal override Excel.Range? GetActiveCell(Excel.Application application) => _activeCell(application);
        internal override bool GetTableActive(Excel.ListObject table) => _tableActive(table);
        internal override int GetRangeRow(Excel.Range range) => _rangeRow(range);
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
        public Mock<Excel.ListColumn>[] Columns { get; } = [.. Enumerable.Range(0, 14).Select(_ => new Mock<Excel.ListColumn>())];
        public Mock<Excel.ListRows> ListRows { get; } = new();
        public Mock<Excel.ListRow> NewRow { get; } = new();
        public Mock<Excel.ListRow> PositionedRow { get; } = new();
        public Mock<Excel.Range> RowRange { get; } = new();
        public Mock<Excel.Range> TableRange { get; } = new();
        public Mock<Excel.Range> TableRows { get; } = new();
        public Mock<Excel.Range> InitialBlankRow { get; } = new();
        public Mock<Excel.Range> ActiveCell { get; } = new();
        public object? WrittenValue { get; private set; }
        public int AddRowCalls { get; private set; }
        public int AddRowAtPositionCalls { get; private set; }
        public int LastInsertionPosition { get; private set; }
        public int ClearRangeCalls { get; private set; }

        public void RecordWrittenValue(object value) => WrittenValue = value;

        public void RecordAddRow() => AddRowCalls++;

        public void RecordAddRowAtPosition(int position)
        {
            AddRowAtPositionCalls++;
            LastInsertionPosition = position;
        }

        public void RecordClearRange() => ClearRangeCalls++;

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
                _ = Columns[index].SetupGet(c => c.Name).Returns(GanttTableSchema.Default.Columns[index].Name);
            }
            _ = Table.SetupGet(t => t.ListRows).Returns(ListRows.Object);
            _ = Table.SetupGet(t => t.Range).Returns(TableRange.Object);
            _ = Table.SetupGet(t => t.Active).Returns(false);
            _ = TableRange.SetupGet(r => r.Row).Returns(1);
            _ = ActiveCell.SetupGet(r => r.Row).Returns(1);
            _ = Application.SetupGet(a => a.ActiveCell).Returns(ActiveCell.Object);
            _ = TableRange.SetupGet(r => r.Rows).Returns(TableRows.Object);
            _ = TableRows.SetupGet(r => r.Count).Returns(2);
            _ = TableRows.SetupGet(r => r[2]).Returns(InitialBlankRow.Object);
            _ = InitialBlankRow.SetupGet(r => r.Value2)
                .Returns(Array.CreateInstance(typeof(object), [1, 14], [1, 1]));
            _ = ListRows.SetupGet(r => r.Count).Returns(1);
            _ = NewRow.SetupGet(r => r.Index).Returns(2);
            _ = NewRow.SetupGet(r => r.Range).Returns(RowRange.Object);
            _ = PositionedRow.SetupGet(r => r.Index).Returns(2);
            _ = PositionedRow.SetupGet(r => r.Range).Returns(RowRange.Object);
            _ = RowRange.SetupSet(r => r.Value2 = It.IsAny<object>())
                .Callback<object>(RecordWrittenValue);
            _ = InitialBlankRow.SetupSet(r => r.Value2 = It.IsAny<object>())
                .Callback<object>(RecordWrittenValue);
            _ = InitialBlankRow.Setup(r => r.ClearContents()).Callback(RecordClearRange);
            _ = ListRows.Setup(r => r.Add(Type.Missing)).Callback(RecordAddRow).Returns(NewRow.Object);
            _ = ListRows.Setup(r => r.Add(It.IsAny<object>()))
                .Callback<object>(position => RecordAddRowAtPosition(
                    Convert.ToInt32(position, System.Globalization.CultureInfo.InvariantCulture)))
                .Returns(PositionedRow.Object);
        }

        public TestableInserter Build(
            IWorksheetProtectionGuard guard,
            ITypeOptionsMaterialiser? typeOptionsMaterialiser = null) => new(
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
            table =>
            {
                Assert.Same(Table.Object, table);
                return TableRange.Object;
            },
            range =>
            {
                Assert.Same(TableRange.Object, range);
                return TableRows.Object;
            },
            rows =>
            {
                Assert.Same(TableRows.Object, rows);
                return 2;
            },
            (rows, index) =>
            {
                Assert.Same(TableRows.Object, rows);
                Assert.Equal(2, index);
                return InitialBlankRow.Object;
            },
            range => range == InitialBlankRow.Object ? range.Value2 : null,
            range => ClearRangeCalls++,
            rows =>
            {
                Assert.Same(ListRows.Object, rows);
                AddRowCalls++;
                return NewRow.Object;
            },
            (rows, position) =>
            {
                Assert.Same(ListRows.Object, rows);
                AddRowAtPositionCalls++;
                LastInsertionPosition = position;
                return PositionedRow.Object;
            },
            _ => ActiveCell.Object,
            table => table.Active,
            range => range.Row,
            row =>
            {
                Assert.True(ReferenceEquals(NewRow.Object, row) || ReferenceEquals(PositionedRow.Object, row));
                return 2;
            },
            row =>
            {
                Assert.True(ReferenceEquals(NewRow.Object, row) || ReferenceEquals(PositionedRow.Object, row));
                return RowRange.Object;
            },
            typeOptionsMaterialiser ?? new StubTypeOptionsMaterialiser());
    }

    [Fact]
    public void Insert_appends_one_bulk_row_with_exact_scaffold_values()
    {
        var graph = new Graph();
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var id = GanttRowId.Parse("G-0123456789abcdef0123456789abcdef");

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
    public void Insert_reuses_the_initial_blank_row_without_appending_a_second_row()
    {
        var graph = new Graph();
        _ = graph.ListRows.SetupGet(r => r.Count).Returns(0);
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var id = GanttRowId.Parse("G-0123456789abcdef0123456789abcdef");

        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            () => id);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.BodyIndex);
        Assert.Equal(0, graph.AddRowCalls);
        var matrix = Assert.IsType<object[,]>(graph.WrittenValue);
        Assert.Equal(id.Value, matrix[0, 0]);
        Assert.Equal("As-Planned Activity", matrix[0, 3]);
        Assert.Equal(0, graph.ClearRangeCalls);
    }

    [Fact]
    public void Insert_clears_the_reused_blank_row_when_type_options_refuses()
    {
        var graph = new Graph();
        _ = graph.ListRows.SetupGet(r => r.Count).Returns(0);
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var materialiser = new Mock<ITypeOptionsMaterialiser>();
        _ = materialiser.Setup(m => m.Materialise())
            .Returns(TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.CatalogueHashMismatch));

        GanttRowInsertOutcome outcome = graph.Build(guard.Object, materialiser.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            GanttRowId.New);

        Assert.Equal(GanttRowInsertOutcome.Refused(GanttRowInsertRefusalReason.TypeOptionsUnavailable), outcome);
        Assert.Equal(0, graph.AddRowCalls);
        Assert.Equal(1, graph.ClearRangeCalls);
        graph.NewRow.Verify(r => r.Delete(), Times.Never);
    }

    [Fact]
    public void Insert_uses_active_row_position_for_a_middle_table_row()
    {
        var graph = new Graph();
        _ = graph.Table.SetupGet(t => t.Active).Returns(true);
        _ = graph.ActiveCell.SetupGet(r => r.Row).Returns(2);
        _ = graph.ListRows.SetupGet(r => r.Count).Returns(2);
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);

        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            GanttRowId.New);

        Assert.True(outcome.Succeeded);
        Assert.Equal(2, outcome.BodyIndex);
        Assert.Equal(1, graph.AddRowAtPositionCalls);
        Assert.Equal(2, graph.LastInsertionPosition);
        Assert.Equal(0, graph.AddRowCalls);
    }

    [Fact]
    public void Insert_uses_append_for_the_last_active_table_row()
    {
        var graph = new Graph();
        _ = graph.Table.SetupGet(t => t.Active).Returns(true);
        _ = graph.ActiveCell.SetupGet(r => r.Row).Returns(2);
        _ = graph.ListRows.SetupGet(r => r.Count).Returns(1);
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);

        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            GanttRowId.New);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, graph.AddRowCalls);
        Assert.Equal(0, graph.AddRowAtPositionCalls);
    }

    [Fact]
    public void Insert_uses_append_when_the_active_cell_is_outside_the_table()
    {
        var graph = new Graph();
        _ = graph.Table.SetupGet(t => t.Active).Returns(false);
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);

        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            GanttRowId.New);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, graph.AddRowCalls);
        Assert.Equal(0, graph.AddRowAtPositionCalls);
    }

    [Fact]
    public void Insert_uses_position_one_when_the_active_cell_is_in_the_header()
    {
        var graph = new Graph();
        _ = graph.Table.SetupGet(t => t.Active).Returns(true);
        _ = graph.ActiveCell.SetupGet(r => r.Row).Returns(1);
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);

        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            GanttRowId.New);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, graph.AddRowAtPositionCalls);
        Assert.Equal(1, graph.LastInsertionPosition);
    }

    [Fact]
    public void Insert_maps_columns_by_header_name_not_position()
    {
        var graph = new Graph();
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        IReadOnlyList<GanttTableColumn> columns = GanttTableSchema.Default.Columns;
        for (var index = 0; index < columns.Count; index++)
        {
            _ = graph.Columns[index].SetupGet(c => c.Name).Returns(columns[index].Name);
        }
        // Reverse the table's physical order; the logical scaffold still lands
        // in the correct header-named cells.
        for (var index = 0; index < columns.Count; index++)
        {
            var schemaIndex = columns.Count - index - 1;
            _ = graph.Columns[index].SetupGet(c => c.Name).Returns(columns[schemaIndex].Name);
        }

        var id = GanttRowId.Parse("G-0123456789abcdef0123456789abcdef");
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
        _ = guard.Setup(g => g.Query()).Returns(guardOutcome);

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
        _ = graph.ListObjects.SetupGet(l => l.Count).Returns(0);
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);

        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            GanttRowId.New);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttRowInsertRefusalReason.TableMissing, outcome.Refusal);
        Assert.Equal(0, graph.AddRowCalls);
        Assert.Null(graph.WrittenValue);
    }

    [Fact]
    public void Insert_deletes_the_new_row_when_type_options_refuses()
    {
        var graph = new Graph();
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var materialiser = new Mock<ITypeOptionsMaterialiser>();
        _ = materialiser.Setup(m => m.Materialise())
            .Returns(TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.CatalogueHashMismatch));

        GanttRowInsertOutcome outcome = graph.Build(guard.Object, materialiser.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            () => GanttRowId.Parse("G-0123456789abcdef0123456789abcdef"));

        Assert.Equal(GanttRowInsertOutcome.Refused(GanttRowInsertRefusalReason.TypeOptionsUnavailable), outcome);
        graph.NewRow.Verify(r => r.Delete(), Times.Once);
    }
}
