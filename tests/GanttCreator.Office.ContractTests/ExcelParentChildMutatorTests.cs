using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for the R5.3 parent/child mutation adapter.
/// </summary>
/// <remarks>
/// The tests drive the adapter through the same fake-COM-graph
/// seam discipline as the R2.8 inserter tests: every COM
/// parameterised property is an <c>internal virtual</c> seam
/// the testable subclass overrides, so the full guard-first
/// pipeline runs without a host. The Core planner runs for
/// real -- these tests pin the ADAPTER's application of the
/// plan, not the planner's arithmetic (that is the Core
/// tests' job).
/// </remarks>
public class ExcelParentChildMutatorTests
{
    /// <summary>The one-based table column index of a schema column.</summary>
    /// <param name="columnName">The schema column name.</param>
    /// <returns>The one-based table column index.</returns>
    private static int TableIndexOf(string columnName)
    {
        for (var index = 0; index < GanttTableSchema.Default.Columns.Count; index++)
        {
            if (GanttTableSchema.Default.Columns[index].Name == columnName)
            {
                return index + 1;
            }
        }

        throw new InvalidOperationException($"'{columnName}' is not a schema column.");
    }

    /// <summary>
    /// Builds one raw body row carrying the fields the hierarchy
    /// planners read, with the remaining schema columns blank.
    /// </summary>
    /// <param name="id">The row's stable identifier.</param>
    /// <param name="type">The row's type display name.</param>
    /// <param name="parentId">The row's parent, or <see langword="null"/>.</param>
    /// <param name="siblingOrder">The row's <c>SiblingOrder</c>, or <see langword="null"/>.</param>
    /// <param name="description">The row's description, which distinguishes content in the move.</param>
    /// <returns>The raw row, one value per table column.</returns>
    private static object?[] BodyRow(
        string id,
        string type,
        string? parentId = null,
        int? siblingOrder = null,
        string description = "Row")
    {
        IReadOnlyList<GanttTableColumn> columns = GanttTableSchema.Default.Columns;
        var row = new object?[columns.Count];
        for (var index = 0; index < columns.Count; index++)
        {
            row[index] = columns[index].Name switch
            {
                "Id" => id,
                "Type" => type,
                "Description" => description,
                "ParentId" => parentId,
                "SiblingOrder" => siblingOrder,
                "Visible" => true,
                _ => null,
            };
        }

        return row;
    }

    /// <summary>
    /// Builds the rectangular <c>Value2</c> payload of the table
    /// body from raw rows, with Excel's one-based lower bounds.
    /// A jagged array cannot express it, so CA1814 is suppressed
    /// here for the same reason it is suppressed in the adapter.
    /// </summary>
    /// <param name="rows">The raw body rows, in body order.</param>
    /// <returns>The body matrix.</returns>
    [System.Diagnostics.CodeAnalysis.SuppressMessage(
        "Performance",
        "CA1814:MultiDimArrays",
        Justification = "Excel's Value2 is a SAFEARRAY and is shaped as a multidimensional array; a jagged array cannot express it.")]
    private static object?[,] Matrix(params object?[][] rows)
    {
        var matrix = new object?[rows.Length, GanttTableSchema.Default.Columns.Count];
        for (var row = 0; row < rows.Length; row++)
        {
            for (var column = 0; column < rows[row].Length; column++)
            {
                matrix[row, column] = rows[row][column];
            }
        }

        return matrix;
    }

    /// <summary>Produces the exception a host throws for a refused write (CA2201).</summary>
    /// <returns>The COM exception for a refused write.</returns>
    private static System.Runtime.InteropServices.COMException RefusedComException()
    {
        var exception = (System.Runtime.InteropServices.COMException?)
            System.Runtime.InteropServices.Marshal.GetExceptionForHR(unchecked((int)0x800A03EC));

        return exception ?? throw new InvalidOperationException(
            "The runtime did not produce a COMException for HRESULT 0x800A03EC.");
    }

    private static string TypeOf(GanttEntityType type) => EntityTypeCatalog.Format(type);

    // Deterministic, parseable row identifiers: "G-" plus 32 hex
    // characters, one per row role in the fixtures below.
    private const string ParentId = "G-0000000000000000000000000000000a";
    private const string ChildOneId = "G-0000000000000000000000000000000b";
    private const string ChildTwoId = "G-0000000000000000000000000000000c";
    private const string ChildThreeId = "G-0000000000000000000000000000000d";
    private const string GroupOneId = "G-0000000000000000000000000000000e";
    private const string GroupTwoId = "G-0000000000000000000000000000000f";
    private const string TargetId = "G-00000000000000000000000000000010";

    /// <summary>
    /// The fake COM graph. Every member the adapter touches outside
    /// a seam is set up here; everything else flows through the
    /// testable subclass's overridden seams.
    /// </summary>
    private sealed class Graph
    {
        public Mock<Excel.Application> Application { get; } = new();
        public Mock<Excel.Workbook> Workbook { get; } = new();
        public Mock<Excel.Sheets> Sheets { get; } = new();
        public Mock<Excel.Worksheet> Worksheet { get; } = new();
        public Mock<Excel.ListObjects> ListObjects { get; } = new();
        public Mock<Excel.ListObject> Table { get; } = new();
        public Mock<Excel.ListColumns> ListColumns { get; } = new();

        /// <summary>Sized from the schema so a schema change cannot leave the fake short.</summary>
        public Mock<Excel.ListColumn>[] Columns { get; } =
            [.. Enumerable.Range(0, GanttTableSchema.Default.Columns.Count)
                .Select(_ => new Mock<Excel.ListColumn>())];

        public Mock<Excel.ListRows> ListRows { get; } = new();
        public Mock<Excel.Range> TableRange { get; } = new();
        public Mock<Excel.Range> DataBodyRange { get; } = new();
        public Mock<Excel.Range> BodyRows { get; } = new();
        public Mock<Excel.Range> ActiveCell { get; } = new();

        private readonly Dictionary<int, Mock<Excel.Range>> _rowRanges = new();
        private readonly Dictionary<(int Row, int Column), Mock<Excel.Range>> _cellRanges = new();
        private readonly Dictionary<Excel.Range, int> _rowIndexes = new();
        private readonly Dictionary<Excel.Range, (int Row, int Column)> _cellIndexes = new();

        /// <summary>The body <c>Value2</c> payload the read returns.</summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage(
            "Performance",
            "CA1814:MultiDimArrays",
            Justification = "Excel's Value2 is a SAFEARRAY and is shaped as a multidimensional array; a jagged array cannot express it.")]
        public object?[,]? BodyMatrix { get; set; }

        /// <summary>The number of body rows. Tests set it from their matrix.</summary>
        public int RowCount { get; set; }

        /// <summary>The worksheet row of the active cell. Table header is at worksheet row 1.</summary>
        public int ActiveRow { get; set; }

        /// <summary>Whether the table owns the selection. Defaults to <see langword="true"/>.</summary>
        public bool TableIsActive { get; set; } = true;

        /// <summary>Whether <c>tblGanttData</c> is present. Defaults to <see langword="true"/>.</summary>
        public bool TablePresent { get; set; } = true;

        /// <summary>Whether the host refuses every write with a COM error.</summary>
        public bool ThrowOnWrite { get; set; }

        /// <summary>Full-width row writes, as (one-based body row, payload).</summary>
        public List<(int Row, object? Payload)> WrittenRows { get; } = [];

        /// <summary>Single-cell writes, as (one-based body row, one-based table column, value).</summary>
        public List<(int Row, int Column, object? Value)> WrittenCells { get; } = [];

        public Graph()
        {
            _ = Application.SetupGet(application => application.ActiveWorkbook)
                .Returns(Workbook.Object);
            _ = Workbook.SetupGet(workbook => workbook.Sheets).Returns(Sheets.Object);
            _ = Sheets.SetupGet(sheets => sheets.Count).Returns(1);
            _ = Worksheet.SetupGet(worksheet => worksheet.ListObjects).Returns(ListObjects.Object);
            _ = ListObjects.SetupGet(listObjects => listObjects.Count)
                .Returns(() => TablePresent ? 1 : 0);
            _ = Table.SetupGet(table => table.Name).Returns(GanttTableSchema.TableName);
            _ = Table.SetupGet(table => table.ListColumns).Returns(ListColumns.Object);
            _ = ListColumns.SetupGet(columns => columns.Count)
                .Returns(GanttTableSchema.Default.Columns.Count);
            for (var index = 0; index < Columns.Length; index++)
            {
                _ = Columns[index].SetupGet(column => column.Name)
                    .Returns(GanttTableSchema.Default.Columns[index].Name);
            }

            _ = Table.SetupGet(table => table.Range).Returns(TableRange.Object);
            _ = TableRange.SetupGet(range => range.Row).Returns(1);
            _ = Table.SetupGet(table => table.DataBodyRange).Returns(DataBodyRange.Object);
            _ = ActiveCell.SetupGet(range => range.Row).Returns(() => ActiveRow);
        }

        /// <summary>
        /// The range of a one-based body row, created on demand
        /// and reverse-mapped so a write names its row.
        /// </summary>
        /// <param name="row">The one-based body row.</param>
        /// <returns>The row's range.</returns>
        public Excel.Range RowRange(int row)
        {
            if (!_rowRanges.TryGetValue(row, out Mock<Excel.Range>? range))
            {
                range = new Mock<Excel.Range>();
                _rowRanges[row] = range;
                _rowIndexes[range.Object] = row;
            }

            return range.Object;
        }

        /// <summary>
        /// The range of one cell, created on demand and
        /// reverse-mapped so a write names its row and column.
        /// </summary>
        /// <param name="row">The one-based body row.</param>
        /// <param name="column">The one-based table column.</param>
        /// <returns>The cell's range.</returns>
        public Excel.Range CellRange(int row, int column)
        {
            if (!_cellRanges.TryGetValue((row, column), out Mock<Excel.Range>? range))
            {
                range = new Mock<Excel.Range>();
                _cellRanges[(row, column)] = range;
                _cellIndexes[range.Object] = (row, column);
            }

            return range.Object;
        }

        /// <summary>Records a write through whichever range kind the adapter addressed.</summary>
        /// <param name="range">The written range.</param>
        /// <param name="payload">The written payload.</param>
        public void RecordWrite(Excel.Range range, object? payload)
        {
            if (ThrowOnWrite)
            {
                throw RefusedComException();
            }

            if (_rowIndexes.TryGetValue(range, out int row))
            {
                WrittenRows.Add((row, payload));
                return;
            }

            if (_cellIndexes.TryGetValue(range, out (int Row, int Column) cell))
            {
                WrittenCells.Add((cell.Row, cell.Column, payload));
                return;
            }

            throw new InvalidOperationException("A write reached an unmapped range.");
        }

        /// <summary>
        /// The values of a written full-width row payload, keyed
        /// by schema column name.
        /// </summary>
        /// <param name="payload">The written <c>object[,]</c> payload.</param>
        /// <returns>Schema column name to value.</returns>
        public static Dictionary<string, object?> ValuesOf(object? payload)
        {
            var matrix = Assert.IsType<object[,]>(payload);
            var values = new Dictionary<string, object?>();
            for (var column = 0; column < GanttTableSchema.Default.Columns.Count; column++)
            {
                values[GanttTableSchema.Default.Columns[column].Name] = matrix[0, column];
            }

            return values;
        }

        /// <summary>Builds the testable adapter over this graph.</summary>
        /// <param name="guard">The protection guard.</param>
        /// <param name="craftedPromotePlan">
        /// A crafted promote plan, when the test drives the apply
        /// path directly instead of the real planner.
        /// </param>
        /// <param name="craftedDemotePlan">
        /// A crafted demote plan, when the test drives the apply
        /// path directly instead of the real planner.
        /// </param>
        /// <returns>The testable adapter.</returns>
        public TestableMutator Build(
            IWorksheetProtectionGuard guard,
            PromotePlan? craftedPromotePlan = null,
            DemotePlan? craftedDemotePlan = null) => new(
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
                return RowCount;
            },
            table =>
            {
                Assert.Same(Table.Object, table);
                return TableRange.Object;
            },
            table =>
            {
                Assert.Same(Table.Object, table);
                return DataBodyRange.Object;
            },
            range =>
            {
                Assert.Same(DataBodyRange.Object, range);
                return BodyRows.Object;
            },
            (rows, index) =>
            {
                Assert.Same(BodyRows.Object, rows);
                return RowRange(index);
            },
            (row, columnIndex) =>
            {
                // A cell write must name a row the adapter
                // already addressed through the rows seam.
                if (!_rowIndexes.TryGetValue(row, out int writtenRow))
                {
                    throw new InvalidOperationException(
                        "A cell write reached an unmapped row range.");
                }

                return CellRange(writtenRow, columnIndex);
            },
            range => range == DataBodyRange.Object ? BodyMatrix : null,
            RecordWrite,
            _ => ActiveCell.Object,
            table => TableIsActive,
            range => range.Row,
            craftedPromotePlan,
            craftedDemotePlan);
    }

    /// <summary>
    /// The adapter under test with every COM parameterised
    /// property routed to the fake graph, and the two plan
    /// seams overridable so a test can drive the apply path
    /// with a crafted plan.
    /// </summary>
    private sealed class TestableMutator(
        Excel.Application? application,
        IWorksheetProtectionGuard? guard,
        Func<Excel.Sheets, int, object> sheetAt,
        Func<Excel.ListObjects, int, Excel.ListObject> tableAt,
        Func<Excel.ListColumns, int, Excel.ListColumn> columnAt,
        Func<Excel.ListObject, Excel.ListRows> listRowsAt,
        Func<Excel.ListRows, int> listRowCountAt,
        Func<Excel.ListObject, Excel.Range> tableRangeAt,
        Func<Excel.ListObject, Excel.Range?> dataBodyRangeAt,
        Func<Excel.Range, Excel.Range> rangeRowsAt,
        Func<Excel.Range, int, Excel.Range> rangeAt,
        Func<Excel.Range, int, Excel.Range> cellAt,
        Func<Excel.Range, object?> rangeValue2At,
        Action<Excel.Range, object> setRangeValues,
        Func<Excel.Application, Excel.Range?> activeCellAt,
        Func<Excel.ListObject, bool> tableActiveAt,
        Func<Excel.Range, int> rangeRowAt,
        PromotePlan? craftedPromotePlan = null,
        DemotePlan? craftedDemotePlan = null)
        : ExcelParentChildMutator(application, guard)
    {
        private readonly Func<Excel.Sheets, int, object> _sheetAt = sheetAt;
        private readonly Func<Excel.ListObjects, int, Excel.ListObject> _tableAt = tableAt;
        private readonly Func<Excel.ListColumns, int, Excel.ListColumn> _columnAt = columnAt;
        private readonly Func<Excel.ListObject, Excel.ListRows> _listRowsAt = listRowsAt;
        private readonly Func<Excel.ListRows, int> _listRowCountAt = listRowCountAt;
        private readonly Func<Excel.ListObject, Excel.Range> _tableRangeAt = tableRangeAt;
        private readonly Func<Excel.ListObject, Excel.Range?> _dataBodyRangeAt = dataBodyRangeAt;
        private readonly Func<Excel.Range, Excel.Range> _rangeRowsAt = rangeRowsAt;
        private readonly Func<Excel.Range, int, Excel.Range> _rangeAt = rangeAt;
        private readonly Func<Excel.Range, int, Excel.Range> _cellAt = cellAt;
        private readonly Func<Excel.Range, object?> _rangeValue2At = rangeValue2At;
        private readonly Action<Excel.Range, object> _setRangeValues = setRangeValues;
        private readonly Func<Excel.Application, Excel.Range?> _activeCellAt = activeCellAt;
        private readonly Func<Excel.ListObject, bool> _tableActiveAt = tableActiveAt;
        private readonly Func<Excel.Range, int> _rangeRowAt = rangeRowAt;
        private readonly PromotePlan? _craftedPromotePlan = craftedPromotePlan;
        private readonly DemotePlan? _craftedDemotePlan = craftedDemotePlan;

        internal override object GetSheetAt(Excel.Sheets sheets, int index) =>
            _sheetAt(sheets, index);

        internal override Excel.ListObject GetTableAt(
            Excel.ListObjects listObjects,
            int index) => _tableAt(listObjects, index);

        internal override Excel.ListColumn GetColumnAt(
            Excel.ListColumns columns,
            int index) => _columnAt(columns, index);

        internal override Excel.ListRows GetListRows(Excel.ListObject table) =>
            _listRowsAt(table);

        internal override int GetListRowCount(Excel.ListRows rows) =>
            _listRowCountAt(rows);

        internal override Excel.Range GetTableRange(Excel.ListObject table) =>
            _tableRangeAt(table);

        internal override Excel.Range? GetDataBodyRange(Excel.ListObject table) =>
            _dataBodyRangeAt(table);

        internal override Excel.Range GetRangeRows(Excel.Range range) =>
            _rangeRowsAt(range);

        internal override Excel.Range GetRangeAt(Excel.Range rows, int index) =>
            _rangeAt(rows, index);

        internal override Excel.Range GetCellAt(Excel.Range row, int columnIndex) =>
            _cellAt(row, columnIndex);

        internal override object? GetRangeValue2(Excel.Range range) =>
            _rangeValue2At(range);

        internal override void SetRangeValues(Excel.Range range, object payload) =>
            _setRangeValues(range, payload);

        internal override Excel.Range? GetActiveCell(Excel.Application application) =>
            _activeCellAt(application);

        internal override bool GetTableActive(Excel.ListObject table) =>
            _tableActiveAt(table);

        internal override int GetRangeRow(Excel.Range range) =>
            _rangeRowAt(range);

        internal override PromotePlan CreatePromotePlan(
            IReadOnlyList<GanttEvent> rows,
            GanttRowId selectedId) =>
            _craftedPromotePlan ?? base.CreatePromotePlan(rows, selectedId);

        internal override DemotePlan CreateDemotePlan(
            IReadOnlyList<GanttEvent> rows,
            GanttRowId selectedId) =>
            _craftedDemotePlan ?? base.CreateDemotePlan(rows, selectedId);
    }

    /// <summary>Seeds a graph with a body matrix and an active-cell row.</summary>
    /// <param name="matrix">The body matrix.</param>
    /// <param name="activeRow">The active cell's worksheet row.</param>
    /// <returns>The seeded graph.</returns>
    [System.Diagnostics.CodeAnalysis.SuppressMessage(
        "Performance",
        "CA1814:MultiDimArrays",
        Justification = "Excel's Value2 is a SAFEARRAY and is shaped as a multidimensional array; a jagged array cannot express it.")]
    private static Graph SeededGraph(object?[,] matrix, int activeRow) => new()
    {
        BodyMatrix = matrix,
        RowCount = matrix.GetLength(0),
        ActiveRow = activeRow,
    };

    /// <summary>A three-row fixture: a parent and two of its children.</summary>
    /// <returns>The body matrix.</returns>
    [System.Diagnostics.CodeAnalysis.SuppressMessage(
        "Performance",
        "CA1814:MultiDimArrays",
        Justification = "Excel's Value2 is a SAFEARRAY and is shaped as a multidimensional array; a jagged array cannot express it.")]
    private static object?[,] ThreeRowRun() => Matrix(
        BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)),
        BodyRow(
            ChildOneId,
            TypeOf(GanttEntityType.AsPlannedActivity),
            ParentId,
            0,
            "First child"),
        BodyRow(
            ChildTwoId,
            TypeOf(GanttEntityType.AsPlannedActivity),
            ParentId,
            1,
            "Second child"));

    /// <summary>A guard that reports no protection anywhere.</summary>
    /// <returns>The guard.</returns>
    private static IWorksheetProtectionGuard NotProtected()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        _ = guard.Setup(g => g.QueryTarget(It.IsAny<object?>()))
            .Returns(ProtectionGuardOutcome.NotProtected);
        return guard.Object;
    }

    [Fact]
    public void Promote_refuses_when_the_guard_reports_no_active_workbook()
    {
        Graph graph = new();
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NoActiveWorkbook);

        ParentChildMutationOutcome outcome = graph.Build(guard.Object).Promote();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ParentChildMutationRefusalReason.NoActiveWorkbook, outcome.Refusal);
        Assert.Null(outcome.PlanRefusal);
        Assert.Null(outcome.SelectedRowNumber);
        Assert.Equal(0, outcome.CellsWritten);
        Assert.Empty(graph.WrittenRows);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void Demote_refuses_when_the_active_sheet_is_protected()
    {
        Graph graph = new();
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.SheetProtected);

        ParentChildMutationOutcome outcome = graph.Build(guard.Object).Demote();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ParentChildMutationRefusalReason.TargetProtected, outcome.Refusal);
        // The workbook is never touched: the guard runs first (ADR-0008 D4).
        graph.Application.VerifyGet(application => application.ActiveWorkbook, Times.Never);
        Assert.Empty(graph.WrittenRows);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void Mutation_refuses_when_the_target_worksheet_is_protected()
    {
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 3);
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        _ = guard.Setup(g => g.QueryTarget(It.IsAny<object?>()))
            .Returns(ProtectionGuardOutcome.SheetProtected);

        ParentChildMutationOutcome promote = graph.Build(guard.Object).Promote();
        ParentChildMutationOutcome demote = graph.Build(guard.Object).Demote();

        Assert.Equal(ParentChildMutationRefusalReason.TargetProtected, promote.Refusal);
        Assert.Equal(ParentChildMutationRefusalReason.TargetProtected, demote.Refusal);
        // The authoritative target check ran for the resolved worksheet.
        guard.Verify(g => g.QueryTarget(graph.Worksheet.Object), Times.Exactly(2));
        Assert.Empty(graph.WrittenRows);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void Mutation_refuses_when_the_table_is_missing()
    {
        Graph graph = new() { TablePresent = false };

        ParentChildMutationOutcome promote = graph.Build(NotProtected()).Promote();
        ParentChildMutationOutcome demote = graph.Build(NotProtected()).Demote();

        Assert.Equal(ParentChildMutationRefusalReason.TableMissing, promote.Refusal);
        Assert.Equal(ParentChildMutationRefusalReason.TableMissing, demote.Refusal);
        Assert.Empty(graph.WrittenRows);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void Mutation_refuses_when_the_selection_is_the_header_row()
    {
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 1);

        ParentChildMutationOutcome outcome = graph.Build(NotProtected()).Promote();

        Assert.Equal(ParentChildMutationRefusalReason.SelectionOutsideTable, outcome.Refusal);
        Assert.Null(outcome.SelectedRowNumber);
        Assert.Empty(graph.WrittenRows);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void Mutation_refuses_when_the_selection_is_below_the_body()
    {
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 5);

        ParentChildMutationOutcome outcome = graph.Build(NotProtected()).Demote();

        Assert.Equal(ParentChildMutationRefusalReason.SelectionOutsideTable, outcome.Refusal);
        Assert.Empty(graph.WrittenRows);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void Mutation_refuses_when_the_table_does_not_own_the_selection()
    {
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 2);
        graph.TableIsActive = false;

        ParentChildMutationOutcome outcome = graph.Build(NotProtected()).Promote();

        Assert.Equal(ParentChildMutationRefusalReason.SelectionOutsideTable, outcome.Refusal);
        Assert.Empty(graph.WrittenRows);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void Mutation_refuses_when_the_row_count_falls_short_of_the_selection()
    {
        // Two body rows but a stale row count of three: the
        // selection names a row the bulk read never saw.
        Graph graph = SeededGraph(
            Matrix(
                BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)),
                BodyRow(
                    ChildOneId,
                    TypeOf(GanttEntityType.AsPlannedActivity),
                    ParentId)),
            activeRow: 4);
        // The stale count names a row the bulk read never saw.
        graph.RowCount = 3;

        ParentChildMutationOutcome outcome = graph.Build(NotProtected()).Promote();

        Assert.Equal(ParentChildMutationRefusalReason.SelectionOutsideTable, outcome.Refusal);
        Assert.Empty(graph.WrittenRows);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void Mutation_refuses_when_the_selected_row_has_an_unparseable_id()
    {
        Graph graph = SeededGraph(
            Matrix(
                BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)),
                BodyRow(
                    "not-an-id",
                    TypeOf(GanttEntityType.AsPlannedActivity),
                    ParentId)),
            activeRow: 3);

        ParentChildMutationOutcome outcome = graph.Build(NotProtected()).Promote();

        Assert.Equal(ParentChildMutationRefusalReason.SelectionUnknown, outcome.Refusal);
        Assert.Empty(graph.WrittenRows);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void Mutation_refuses_when_any_other_row_has_an_unparseable_id()
    {
        Graph graph = SeededGraph(
            Matrix(
                BodyRow("not-an-id", TypeOf(GanttEntityType.AsPlannedActivity)),
                BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity))),
            activeRow: 3);

        ParentChildMutationOutcome outcome = graph.Build(NotProtected()).Demote();

        Assert.Equal(ParentChildMutationRefusalReason.TableInvalid, outcome.Refusal);
        Assert.Empty(graph.WrittenRows);
        Assert.Empty(graph.WrittenCells);
    }

    /// <summary>
    /// A parent with two children whose original SiblingOrder
    /// values are distinctive, so a test can tell a kept value
    /// from a resequenced one.
    /// </summary>
    /// <returns>The body matrix.</returns>
    [System.Diagnostics.CodeAnalysis.SuppressMessage(
        "Performance",
        "CA1814:MultiDimArrays",
        Justification = "Excel's Value2 is a SAFEARRAY and is shaped as a multidimensional array; a jagged array cannot express it.")]
    private static object?[,] TwoChildRun() => Matrix(
        BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)),
        BodyRow(
            ChildOneId,
            TypeOf(GanttEntityType.AsPlannedActivity),
            ParentId,
            7,
            "First child"),
        BodyRow(
            ChildTwoId,
            TypeOf(GanttEntityType.AsPlannedActivity),
            ParentId,
            1,
            "Second child"));

    [Fact]
    public void Promote_moves_the_selected_child_to_the_run_bottom()
    {
        Graph graph = SeededGraph(TwoChildRun(), activeRow: 3);

        ParentChildMutationOutcome outcome = graph.Build(NotProtected()).Promote();

        Assert.True(outcome.Succeeded);
        Assert.Null(outcome.Refusal);
        Assert.Null(outcome.PlanRefusal);
        Assert.Equal(2, outcome.SelectedRowNumber);
        // The two-row window moves full width: 2 rows x 16 columns.
        Assert.Equal(2 * GanttTableSchema.Default.Columns.Count, outcome.CellsWritten);
        Assert.Empty(graph.WrittenCells);
        Assert.Equal(2, graph.WrittenRows.Count);

        // Row 2 now holds the second child's content, resequenced to 0.
        Assert.Equal(2, graph.WrittenRows[0].Row);
        Dictionary<string, object?> movedUp = Graph.ValuesOf(graph.WrittenRows[0].Payload);
        Assert.Equal(ChildTwoId, movedUp["Id"]);
        Assert.Equal("Second child", movedUp["Description"]);
        Assert.Equal(ParentId, movedUp["ParentId"]);
        Assert.Equal(0, movedUp["SiblingOrder"]);

        // Row 3 now holds the promoted child with its ParentId
        // cleared. Its own SiblingOrder keeps its pre-move value:
        // the plan emits no write for the promoted row, and the
        // adapter invents none (R5.3 D4).
        Assert.Equal(3, graph.WrittenRows[1].Row);
        Dictionary<string, object?> promoted = Graph.ValuesOf(graph.WrittenRows[1].Payload);
        Assert.Equal(ChildOneId, promoted["Id"]);
        Assert.Equal("First child", promoted["Description"]);
        Assert.Null(promoted["ParentId"]);
        Assert.Equal(7, promoted["SiblingOrder"]);
    }

    [Fact]
    public void Promote_of_a_middle_child_resequences_the_run_member_above_the_window()
    {
        // parent, first child, second child, third child;
        // the selection is the second child.
        Graph graph = SeededGraph(
            Matrix(
                BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)),
                BodyRow(
                    ChildOneId,
                    TypeOf(GanttEntityType.AsPlannedActivity),
                    ParentId,
                    0,
                    "First child"),
                BodyRow(
                    ChildTwoId,
                    TypeOf(GanttEntityType.AsPlannedActivity),
                    ParentId,
                    1,
                    "Second child"),
                BodyRow(
                    ChildThreeId,
                    TypeOf(GanttEntityType.AsPlannedActivity),
                    ParentId,
                    2,
                    "Third child")),
            activeRow: 4);

        ParentChildMutationOutcome outcome = graph.Build(NotProtected()).Promote();

        Assert.True(outcome.Succeeded);
        Assert.Equal(3, outcome.SelectedRowNumber);
        // The two-row window moves full width, plus one
        // SiblingOrder cell above the window.
        Assert.Equal(2 * GanttTableSchema.Default.Columns.Count + 1, outcome.CellsWritten);
        Assert.Equal(2, graph.WrittenRows.Count);
        Assert.Single(graph.WrittenCells);

        // Row 3 now holds the third child, resequenced to 1.
        Assert.Equal(3, graph.WrittenRows[0].Row);
        Dictionary<string, object?> shiftedUp = Graph.ValuesOf(graph.WrittenRows[0].Payload);
        Assert.Equal(ChildThreeId, shiftedUp["Id"]);
        Assert.Equal(1, shiftedUp["SiblingOrder"]);

        // Row 4 now holds the promoted second child, parent cleared.
        Assert.Equal(4, graph.WrittenRows[1].Row);
        Dictionary<string, object?> promoted = Graph.ValuesOf(graph.WrittenRows[1].Payload);
        Assert.Equal(ChildTwoId, promoted["Id"]);
        Assert.Null(promoted["ParentId"]);

        // The first child sits above the moved window: only
        // its SiblingOrder cell is written.
        Assert.Equal(2, graph.WrittenCells[0].Row);
        Assert.Equal(TableIndexOf("SiblingOrder"), graph.WrittenCells[0].Column);
        Assert.Equal(0, graph.WrittenCells[0].Value);
    }

    [Fact]
    public void Promote_of_the_last_child_writes_the_row_and_the_run_above()
    {
        // The selection is the last child: it is already at
        // the run bottom, so the moved window is one row.
        Graph graph = SeededGraph(TwoChildRun(), activeRow: 4);

        ParentChildMutationOutcome outcome = graph.Build(NotProtected()).Promote();

        Assert.True(outcome.Succeeded);
        Assert.Equal(3, outcome.SelectedRowNumber);
        // One full-width row plus the first child's
        // resequenced SiblingOrder cell above the window.
        Assert.Equal(GanttTableSchema.Default.Columns.Count + 1, outcome.CellsWritten);
        Assert.Single(graph.WrittenRows);
        Assert.Single(graph.WrittenCells);

        Assert.Equal(3, graph.WrittenRows[0].Row);
        Dictionary<string, object?> promoted = Graph.ValuesOf(graph.WrittenRows[0].Payload);
        Assert.Equal(ChildTwoId, promoted["Id"]);
        Assert.Null(promoted["ParentId"]);

        Assert.Equal(2, graph.WrittenCells[0].Row);
        Assert.Equal(TableIndexOf("SiblingOrder"), graph.WrittenCells[0].Column);
        Assert.Equal(0, graph.WrittenCells[0].Value);
    }

    [Fact]
    public void Demote_assigns_the_nearest_eligible_parent_and_writes_only_the_planned_cells()
    {
        // Two top-level rows: the second is the selection.
        Graph graph = SeededGraph(
            Matrix(
                BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)),
                BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity))),
            activeRow: 3);

        ParentChildMutationOutcome outcome = graph.Build(NotProtected()).Demote();

        Assert.True(outcome.Succeeded);
        Assert.Equal(2, outcome.SelectedRowNumber);
        // ParentId + the demoted row's tail assignment +
        // the resequencing of the top-level run it leaves.
        Assert.Equal(3, outcome.CellsWritten);
        Assert.Empty(graph.WrittenRows);
        Assert.Equal(3, graph.WrittenCells.Count);

        var expected = new List<(int Row, int Column, object? Value)>
        {
            (2, TableIndexOf("ParentId"), ParentId),
            (2, TableIndexOf("SiblingOrder"), 0),
            (1, TableIndexOf("SiblingOrder"), 0),
        };
        Assert.Equal(
            expected.OrderBy(cell => (cell.Row, cell.Column)),
            graph.WrittenCells.OrderBy(cell => (cell.Row, cell.Column)));
    }

    [Fact]
    public void Demoting_adjacent_to_a_full_group_is_refused_as_parent_unavailable()
    {
        // Group one (top-level, no children), group two
        // (top-level, at the seven-child capacity), group
        // two's seven children, then the target (top-level).
        // The target extends group two's run, and no parent
        // further up can be adjacent, so the capacity rule
        // refuses instead of skipping to a landing the
        // outline cannot represent (R5.3 D3).
        var rows = new List<object?[]>
        {
            BodyRow(GroupOneId, TypeOf(GanttEntityType.AsPlannedActivity), null, 0, "Group one"),
            BodyRow(GroupTwoId, TypeOf(GanttEntityType.AsPlannedActivity), null, 1, "Group two"),
        };
        for (var child = 0; child < 7; child++)
        {
            rows.Add(BodyRow(
                $"G-{1000 + child:x32}",
                TypeOf(GanttEntityType.AsPlannedActivity),
                GroupTwoId,
                child,
                $"Group two child {child}"));
        }

        rows.Add(BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity), null, 2, "Target"));
        Graph graph = SeededGraph(Matrix([.. rows]), activeRow: 11);

        ParentChildMutationOutcome outcome = graph.Build(NotProtected()).Demote();

        Assert.False(outcome.Succeeded);
        Assert.Null(outcome.Refusal);
        Assert.Equal(ParentChildRefusalReason.ParentUnavailable, outcome.PlanRefusal);
        Assert.Equal(10, outcome.SelectedRowNumber);
        Assert.Equal(0, outcome.CellsWritten);
        Assert.Empty(graph.WrittenRows);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void Promote_of_a_top_level_row_is_refused_as_not_a_child()
    {
        // The selection is the parent row, which has no parent.
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 2);

        ParentChildMutationOutcome outcome = graph.Build(NotProtected()).Promote();

        Assert.False(outcome.Succeeded);
        Assert.Null(outcome.Refusal);
        Assert.Equal(ParentChildRefusalReason.NotAChild, outcome.PlanRefusal);
        Assert.Equal(1, outcome.SelectedRowNumber);
        Assert.Equal(0, outcome.CellsWritten);
        Assert.Empty(graph.WrittenRows);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void Promote_of_a_critical_interval_is_refused()
    {
        // A CriticalInterval is child-only by contract, so
        // it cannot be promoted out of its run.
        Graph graph = SeededGraph(
            Matrix(
                BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)),
                BodyRow(
                    ChildOneId,
                    TypeOf(GanttEntityType.CriticalInterval),
                    ParentId)),
            activeRow: 3);

        ParentChildMutationOutcome outcome = graph.Build(NotProtected()).Promote();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ParentChildRefusalReason.PromotionForbiddenForType, outcome.PlanRefusal);
        Assert.Equal(2, outcome.SelectedRowNumber);
        Assert.Empty(graph.WrittenRows);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void Promote_of_a_row_that_owns_children_is_refused()
    {
        // The selected row has a child of its own; promoting
        // it would orphan that child.
        Graph graph = SeededGraph(
            Matrix(
                BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)),
                BodyRow(
                    ChildOneId,
                    TypeOf(GanttEntityType.AsPlannedActivity),
                    ParentId),
                BodyRow(
                    ChildTwoId,
                    TypeOf(GanttEntityType.AsPlannedActivity),
                    ChildOneId)),
            activeRow: 3);

        ParentChildMutationOutcome outcome = graph.Build(NotProtected()).Promote();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ParentChildRefusalReason.HasChildren, outcome.PlanRefusal);
        Assert.Equal(2, outcome.SelectedRowNumber);
        Assert.Empty(graph.WrittenRows);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void Demote_with_no_eligible_parent_above_is_refused()
    {
        // The selection is the first row: nothing sits above it.
        Graph graph = SeededGraph(
            Matrix(BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity))),
            activeRow: 2);

        ParentChildMutationOutcome outcome = graph.Build(NotProtected()).Demote();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ParentChildRefusalReason.NoEligibleParentAbove, outcome.PlanRefusal);
        Assert.Equal(1, outcome.SelectedRowNumber);
        Assert.Empty(graph.WrittenRows);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void Demote_of_a_row_that_is_already_a_child_is_refused()
    {
        Graph graph = SeededGraph(
            Matrix(
                BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)),
                BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity), ParentId)),
            activeRow: 3);

        ParentChildMutationOutcome outcome = graph.Build(NotProtected()).Demote();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ParentChildRefusalReason.AlreadyAChild, outcome.PlanRefusal);
        Assert.Equal(2, outcome.SelectedRowNumber);
        Assert.Empty(graph.WrittenRows);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void A_promote_plan_that_moves_upward_is_refused_rather_than_rotated()
    {
        // A crafted plan whose move target precedes the
        // selection: the plan and the table disagree about
        // row positions, so the adapter refuses instead of
        // rotating the wrong window.
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 3);
        var plan = new PromotePlan(
            GanttRowId.Parse(ChildOneId),
            GanttRowId.Parse(ParentId),
            MoveToBodyIndex: 1,
            SiblingOrderWrites: [],
            Refusal: null);
        IParentChildMutator mutator = graph.Build(
            NotProtected(),
            craftedPromotePlan: plan);

        ParentChildMutationOutcome outcome = mutator.Promote();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ParentChildMutationRefusalReason.TableInvalid, outcome.Refusal);
        Assert.Empty(graph.WrittenRows);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void A_host_write_refusal_is_reported_as_failed()
    {
        Graph promoteGraph = SeededGraph(ThreeRowRun(), activeRow: 3);
        promoteGraph.ThrowOnWrite = true;
        Graph demoteGraph = SeededGraph(
            Matrix(
                BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)),
                BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity))),
            activeRow: 3);
        demoteGraph.ThrowOnWrite = true;

        ParentChildMutationOutcome promote = promoteGraph.Build(NotProtected()).Promote();
        ParentChildMutationOutcome demote = demoteGraph.Build(NotProtected()).Demote();

        Assert.Equal(ParentChildMutationRefusalReason.WriteFailed, promote.Refusal);
        Assert.Equal(ParentChildMutationRefusalReason.WriteFailed, demote.Refusal);
        Assert.Equal(0, promote.CellsWritten);
        Assert.Equal(0, demote.CellsWritten);
    }
}