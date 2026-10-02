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

        public TypeOptionsMaterialiseOutcome EnsureCurrent() => TypeOptionsMaterialiseOutcome.Ok();
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
        /// <summary>
        /// Sized from the schema rather than a literal, so adding a column cannot
        /// leave the fake short of a mock and fail with an IndexOutOfRange that
        /// names neither the column nor the contract.
        /// </summary>
        public Mock<Excel.ListColumn>[] Columns { get; } =
            [.. Enumerable.Range(0, GanttTableSchema.Default.Columns.Count).Select(_ => new Mock<Excel.ListColumn>())];
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
            row =>
            {
                Assert.Same(NewRow.Object, row);
                return 2;
            },
            row =>
            {
                Assert.Same(NewRow.Object, row);
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
        Assert.Equal(GanttTableSchema.Default.Columns.Count, matrix.GetLength(1));

        // Unreversed headers here, so a logical column's physical position IS its
        // schema index -- looked up rather than written as a literal. The previous
        // pinned `matrix[0, 8]` for StyleKey silently became SiblingOrder's cell
        // when R4.7A inserted that column, which is how the production
        // GanttRowDefaults defect was found.
        int Of(string columnName)
        {
            for (var i = 0; i < GanttTableSchema.Default.Columns.Count; i++)
            {
                if (GanttTableSchema.Default.Columns[i].Name == columnName)
                {
                    return i;
                }
            }

            throw new InvalidOperationException($"Column '{columnName}' is not in the schema.");
        }

        Assert.Equal(id.Value, matrix[0, Of("Id")]);
        Assert.Equal("As-Planned Activity", matrix[0, Of("Type")]);
        Assert.Equal("AsPlannedActivity", matrix[0, Of("StyleKey")]);

        // SiblingOrder is engine-maintained and has no scaffold default. The
        // writer coalesces a null cell value to string.Empty, so "blank" is an
        // empty string here, not null.
        Assert.Equal(string.Empty, matrix[0, Of("SiblingOrder")]);
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
        _ = materialiser.Setup(m => m.EnsureCurrent())
            .Returns(TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.CatalogueHashMismatch));

        GanttRowInsertOutcome outcome = graph.Build(guard.Object, materialiser.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            GanttRowId.New);

        Assert.Equal(GanttRowInsertOutcome.Refused(GanttRowInsertRefusalReason.TypeOptionsUnavailable), outcome);
        Assert.Equal(0, graph.AddRowCalls);
        Assert.Equal(1, graph.ClearRangeCalls);
        graph.NewRow.Verify(r => r.Delete(), Times.Never);
    }

    /// <summary>
    /// ADR-0035 D3: EVERY insert appends at the end of the body, whatever the
    /// active cell was. Nothing below the insertion point moves.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This test replaces four tests that asserted the OPPOSITE -- insert at the
    /// active row, append for the last row, append when the active cell was outside
    /// the table, and insert at position one for the header. Each described a branch
    /// of <c>GetInsertionPosition</c>, which no longer exists. They were not deleted
    /// silently: the single property they collectively implied is the one worth
    /// keeping, and it is stronger than any of them -- the position is not merely
    /// different, it is <em>absent</em>.
    /// </para>
    /// <para>
    /// The active cell is varied across every case the old code branched on. If any
    /// of them still produced a positioned insert, this theory would catch it.
    /// </para>
    /// </remarks>
    [Theory]
    [InlineData(true, 1)]  // header row selected
    [InlineData(true, 2)]  // first body row
    [InlineData(true, 5)]  // middle body row -- the case that used to shift rows down
    [InlineData(true, 9)]  // last body row
    [InlineData(false, 4)] // active cell outside the table
    public void Every_insert_appends_and_never_uses_the_active_cell(bool tableActive, int activeRow)
    {
        var graph = new Graph();
        _ = graph.Table.SetupGet(t => t.Active).Returns(tableActive);
        _ = graph.ActiveCell.SetupGet(r => r.Row).Returns(activeRow);
        _ = graph.ListRows.SetupGet(r => r.Count).Returns(3);
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);

        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            GanttRowId.New);

        Assert.True(outcome.Succeeded);

        // The ONLY mutation is the positional-free append, for every active cell.
        Assert.Equal(1, graph.AddRowCalls);
        Assert.Equal(0, graph.AddRowAtPositionCalls);

        // The row really was written, so "appends" is not satisfied by a no-op.
        Assert.NotNull(graph.WrittenValue);
    }

    /// <summary>
    /// No positional insert path survives anywhere in the adapter (ADR-0035 D3).
    /// </summary>
    /// <remarks>
    /// <para>
    /// The behavioural theory above cannot prove this on its own, and the reason is
    /// structural rather than incidental: <c>ListRows</c> exposes a single
    /// <c>Add(object)</c> method, so a positional insert and an append are the SAME
    /// method call distinguished only by the argument -- <c>Add(Type.Missing)</c>
    /// versus <c>Add(position)</c>. A Moq <c>Verify(r =&gt; r.Add(It.IsAny&lt;object&gt;()))</c>
    /// therefore matches <em>both</em> and cannot separate them.
    /// </para>
    /// <para>
    /// Asserting on the source is the honest way to pin this, and the repository
    /// already does it for the layout authority and the protection-guard patterns,
    /// so it follows an established convention rather than inventing one.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_adapter_contains_no_positional_insert_path()
    {
        string source = File.ReadAllText(LocateInserterSource());

        // The method that resolved an insertion position is gone, not merely unused.
        Assert.DoesNotContain("GetInsertionPosition", source, StringComparison.Ordinal);
        Assert.DoesNotContain("AddRowAtPosition", source, StringComparison.Ordinal);

        // The only Add call is the positional-free append.
        Assert.Contains("rows.Add(Type.Missing)", source, StringComparison.Ordinal);
        Assert.DoesNotContain("rows.Add(position)", source, StringComparison.Ordinal);

        // The active cell is never read, so it cannot influence the insert.
        Assert.DoesNotContain("ActiveCell", source, StringComparison.Ordinal);
        Assert.DoesNotContain("table.Active", source, StringComparison.Ordinal);
    }

    private static string LocateInserterSource()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            string candidate = Path.Combine(
                directory.FullName,
                "src",
                "GanttCreator.Office",
                "ExcelGanttRowInserter.cs");
            if (File.Exists(candidate))
            {
                return candidate;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException(
            "Could not locate src/GanttCreator.Office/ExcelGanttRowInserter.cs from " + AppContext.BaseDirectory);
    }

    [Fact]
    public void Insert_maps_columns_by_header_name_not_position()
    {
        var graph = new Graph();
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        IReadOnlyList<GanttTableColumn> columns = GanttTableSchema.Default.Columns;

        // Reverse the table's physical order so the logical scaffold cannot be
        // found by position alone.
        for (var index = 0; index < columns.Count; index++)
        {
            var schemaIndex = columns.Count - index - 1;
            _ = graph.Columns[index].SetupGet(c => c.Name).Returns(columns[schemaIndex].Name);
        }

        // Ask the fake which physical index now carries each header, rather than
        // recomputing it. GetColumnAt is 1-based and resolves to Columns[index - 1],
        // so a second arithmetic copy of that rule is a second place to be wrong
        // by one, and the failure reads as a production defect when it is a
        // mistake in the expectation.
        int PhysicalOf(string logicalName) => Array.FindIndex(graph.Columns, c => c.Object.Name == logicalName);

        var id = GanttRowId.Parse("G-0123456789abcdef0123456789abcdef");
        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(GanttEntityType.Delineator, () => id);

        Assert.True(outcome.Succeeded);
        var matrix = Assert.IsType<object[,]>(graph.WrittenValue);

        // The physical column for a logical column is the REVERSED index, because
        // the loop above gave physical index i the name of logical
        // (Count - 1 - i). Deriving both from the schema means adding a column
        // moves them together; a pinned literal would have asserted the wrong
        // cell and still passed, which is the failure mode this test exists to
        // rule out.
        // The scaffold writes Id, StyleKey and Type and leaves the rest null --
        // SortOrder included, which has no default. The original assertion of
        // string.Empty at a pinned physical 0 was only true by accident: it read
        // whatever column the reversal happened to put there. What this test is
        // for is that the named values land in the header-named cells.
        Assert.Equal(id.Value, matrix[0, PhysicalOf("Id")]);
        Assert.Equal("Delineator", matrix[0, PhysicalOf("Type")]);
        Assert.Equal("DefaultDelineator", matrix[0, PhysicalOf("StyleKey")]);
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
        _ = materialiser.Setup(m => m.EnsureCurrent())
            .Returns(TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.CatalogueHashMismatch));

        GanttRowInsertOutcome outcome = graph.Build(guard.Object, materialiser.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            () => GanttRowId.Parse("G-0123456789abcdef0123456789abcdef"));

        Assert.Equal(GanttRowInsertOutcome.Refused(GanttRowInsertRefusalReason.TypeOptionsUnavailable), outcome);
        graph.NewRow.Verify(r => r.Delete(), Times.Once);
    }
}
