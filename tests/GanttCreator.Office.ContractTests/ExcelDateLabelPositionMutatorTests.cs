using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for the R5.14 date-label mutation adapter.
/// </summary>
/// <remarks>
/// Mirrors <c>ExcelLabelPositionMutatorTests</c> (R5.5) but the
/// capability is type-only (R5.14 D4): the adapter reads no style
/// registry, so there is no catalogue reader and no style-related
/// refusal. The tests pin the write set — EXACTLY ONE
/// <c>DateLabelPosition</c> cell on the selected row — and every
/// typed refusal before the write, plus the null-clears-to-Auto
/// path (R5.14 D3).
/// </remarks>
public class ExcelDateLabelPositionMutatorTests
{
    private const string RowOneId = "G-0000000000000000000000000000000a";
    private const string RowTwoId = "G-0000000000000000000000000000000b";
    private const string RowThreeId = "G-0000000000000000000000000000000c";

    private sealed class Graph
    {
        private readonly Dictionary<int, Mock<Excel.ListRow>> _listRows = new();
        private readonly Dictionary<Excel.ListRow, int> _listRowIndexes = new();
        private readonly Dictionary<int, Mock<Excel.Range>> _rowRanges = new();
        private readonly Dictionary<Excel.Range, int> _rowRangeIndexes = new();
        private readonly Dictionary<(int Row, int Column), Mock<Excel.Range>> _cellRanges = new();
        private readonly Dictionary<Excel.Range, (int Row, int Column)> _cellIndexes = new();

        public Mock<Excel.Application> Application { get; } = new();
        public Mock<Excel.Workbook> Workbook { get; } = new();
        public Mock<Excel.Sheets> Sheets { get; } = new();
        public Mock<Excel.Worksheet> Worksheet { get; } = new();
        public Mock<Excel.ListObjects> ListObjects { get; } = new();
        public Mock<Excel.ListObject> Table { get; } = new();
        public Mock<Excel.ListColumns> ListColumns { get; } = new();
        public Mock<Excel.ListColumn>[] Columns { get; }
        public Mock<Excel.ListRows> ListRows { get; } = new();
        public Mock<Excel.Range> TableRange { get; } = new();
        public Mock<Excel.Range> DataBodyRange { get; } = new();
        public Mock<Excel.Range> ActiveCell { get; } = new();

        [System.Diagnostics.CodeAnalysis.SuppressMessage(
            "Performance",
            "CA1814:MultiDimArrays",
            Justification = "Excel Value2 is a SAFEARRAY shaped as a multidimensional array; a jagged array cannot express it.")]
        public object?[,]? BodyMatrix { get; set; }
        public int RowCount { get; set; }
        public int ActiveRow { get; set; }
        public bool TableIsActive { get; set; } = true;
        public bool TablePresent { get; set; } = true;
        public bool ThrowOnWrite { get; set; }

        public List<(int Row, int Column, object? Value)> WrittenCells { get; } = [];

        public Graph(IReadOnlyList<string>? columnNames = null)
        {
            IReadOnlyList<string> names = columnNames
                ?? GanttTableSchema.Default.Columns.Select(static c => c.Name).ToArray();
            Columns = Enumerable.Range(0, names.Count)
                .Select(_ => new Mock<Excel.ListColumn>()).ToArray();
            Application.SetupGet(a => a.ActiveWorkbook).Returns(Workbook.Object);
            Workbook.SetupGet(w => w.Sheets).Returns(Sheets.Object);
            Sheets.SetupGet(s => s.Count).Returns(1);
            Worksheet.SetupGet(w => w.ListObjects).Returns(ListObjects.Object);
            ListObjects.SetupGet(lo => lo.Count).Returns(() => TablePresent ? 1 : 0);
            Table.SetupGet(t => t.Name).Returns("tblGanttData");
            Table.SetupGet(t => t.ListColumns).Returns(ListColumns.Object);
            ListColumns.SetupGet(c => c.Count).Returns(names.Count);
            for (var i = 0; i < Columns.Length; i++)
            {
                Columns[i].SetupGet(c => c.Name).Returns(names[i]);
            }

            Table.SetupGet(t => t.Range).Returns(TableRange.Object);
            TableRange.SetupGet(r => r.Row).Returns(1);
            Table.SetupGet(t => t.DataBodyRange).Returns(DataBodyRange.Object);
            DataBodyRange.SetupGet(r => r.Row).Returns(2);
            ActiveCell.SetupGet(r => r.Row).Returns(() => ActiveRow);
        }

        public Excel.ListRow ListRowAt(int index)
        {
            if (!_listRows.TryGetValue(index, out Mock<Excel.ListRow>? value))
            {
                value = new Mock<Excel.ListRow>();
                _listRows[index] = value;
                _listRowIndexes[value.Object] = index;
            }

            return value.Object;
        }

        public Excel.Range RowRangeOfListRow(Excel.ListRow row) => RowRange(_listRowIndexes[row]);

        public Excel.Range CellOfRowRange(Excel.Range rowRange, int column) =>
            CellRange(_rowRangeIndexes[rowRange], column);

        public Excel.Range RowRange(int row)
        {
            if (!_rowRanges.TryGetValue(row, out Mock<Excel.Range>? value))
            {
                value = new Mock<Excel.Range>();
                _rowRanges[row] = value;
                _rowRangeIndexes[value.Object] = row;
            }

            return value.Object;
        }

        public Excel.Range CellRange(int row, int column)
        {
            if (!_cellRanges.TryGetValue((row, column), out Mock<Excel.Range>? value))
            {
                value = new Mock<Excel.Range>();
                _cellRanges[(row, column)] = value;
                _cellIndexes[value.Object] = (row, column);
            }

            return value.Object;
        }

        public void RecordWrite(Excel.Range range, object? payload)
        {
            if (ThrowOnWrite)
            {
                throw RefusedComException();
            }

            if (_cellIndexes.TryGetValue(range, out (int Row, int Column) cell))
            {
                WrittenCells.Add((cell.Row, cell.Column, payload));
                return;
            }

            throw new InvalidOperationException("A write reached an unmapped range.");
        }

        public TestableDateLabelMutator Build(IWorksheetProtectionGuard guard)
        {
            var mutator = new TestableDateLabelMutator(
                Application.Object,
                guard,
                this);
            return mutator;
        }
    }

    private sealed class TestableDateLabelMutator(
        Excel.Application? application,
        IWorksheetProtectionGuard guard,
        Graph graph) : ExcelDateLabelPositionMutator(application, guard)
    {
        private readonly Graph _graph = graph;

        internal override int GetSheetCount(Excel.Sheets sheets) => 1;

        internal override object GetSheetAt(Excel.Sheets sheets, int index) => _graph.Worksheet.Object;

        internal override Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index) =>
            _graph.Table.Object;

        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index) =>
            _graph.Columns[index - 1].Object;

        internal override Excel.ListRows GetListRows(Excel.ListObject table) => _graph.ListRows.Object;

        internal override int GetListRowCount(Excel.ListRows rows) => _graph.RowCount;

        internal override Excel.Range GetTableRange(Excel.ListObject table) => _graph.TableRange.Object;

        internal override Excel.Range? GetDataBodyRange(Excel.ListObject table) => _graph.DataBodyRange.Object;

        internal override object? GetRangeValue2(Excel.Range range) =>
            ReferenceEquals(range, _graph.DataBodyRange.Object) ? _graph.BodyMatrix : null;

        internal override Excel.Range? GetActiveCell(Excel.Application application) =>
            _graph.ActiveCell.Object;

        internal override bool GetTableActive(Excel.ListObject table) => _graph.TableIsActive;

        internal override Excel.ListRow GetListRowAt(Excel.ListRows rows, int index) =>
            _graph.ListRowAt(index);

        internal override Excel.Range GetListRowRange(Excel.ListRow row) =>
            _graph.RowRangeOfListRow(row);

        internal override Excel.Range GetCellAt(Excel.Range row, int index) =>
            _graph.CellOfRowRange(row, index);

        internal override void SetRangeValues(Excel.Range range, object? value) =>
            _graph.RecordWrite(range, value);
    }

    private static int TableIndexOf(string columnName)
    {
        for (var i = 0; i < GanttTableSchema.Default.Columns.Count; i++)
        {
            if (string.Equals(GanttTableSchema.Default.Columns[i].Name, columnName, StringComparison.Ordinal))
            {
                return i + 1;
            }
        }

        throw new InvalidOperationException($"'{columnName}' is not a schema column.");
    }

    private static object?[] DateLabelRow(string id, string type, string? dateLabelPosition = null)
    {
        IReadOnlyList<GanttTableColumn> columns = GanttTableSchema.Default.Columns;
        var row = new object?[columns.Count];
        for (var index = 0; index < columns.Count; index++)
        {
            row[index] = columns[index].Name switch
            {
                "Id" => id,
                "Type" => type,
                "DateLabelPosition" => dateLabelPosition,
                "Visible" => true,
                _ => null,
            };
        }

        return row;
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage(
        "Performance",
        "CA1814:MultiDimArrays",
        Justification = "Excel Value2 is a SAFEARRAY shaped as a multidimensional array; a jagged array cannot express it.")]
    private static object?[,] Matrix(params object?[][] rows)
    {
        var array = new object?[rows.Length, GanttTableSchema.Default.Columns.Count];
        for (var i = 0; i < rows.Length; i++)
        {
            for (var j = 0; j < rows[i].Length; j++)
            {
                array[i, j] = rows[i][j];
            }
        }

        return array;
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage(
        "Performance",
        "CA1814:MultiDimArrays",
        Justification = "Excel Value2 is a SAFEARRAY shaped as a multidimensional array; a jagged array cannot express it.")]
    private static object?[,] LegacyMatrix(IReadOnlyList<string> legacyColumns, params object?[][] rows)
    {
        var array = new object?[rows.Length, legacyColumns.Count];
        for (var i = 0; i < rows.Length; i++)
        {
            for (var j = 0; j < rows[i].Length && j < legacyColumns.Count; j++)
            {
                array[i, j] = rows[i][j];
            }
        }

        return array;
    }

    private static object?[] LegacyRow(IReadOnlyList<string> legacyColumns, string id, string type)
    {
        var row = new object?[legacyColumns.Count];
        for (var index = 0; index < legacyColumns.Count; index++)
        {
            row[index] = legacyColumns[index] switch
            {
                "Id" => id,
                "Type" => type,
                "Visible" => true,
                _ => null,
            };
        }

        return row;
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage(
        "Performance",
        "CA1814:MultiDimArrays",
        Justification = "Excel Value2 is a SAFEARRAY shaped as a multidimensional array; a jagged array cannot express it.")]
    private static Graph SeededGraph(object?[,] matrix, int activeRow)
    {
        return new Graph
        {
            BodyMatrix = matrix,
            RowCount = matrix.GetLength(0),
            ActiveRow = activeRow,
        };
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage(
        "Performance",
        "CA1814:MultiDimArrays",
        Justification = "Excel Value2 is a SAFEARRAY shaped as a multidimensional array; a jagged array cannot express it.")]
    private static object?[,] ThreeRowRun() =>
        Matrix(
            DateLabelRow(RowOneId, EntityTypeCatalog.Format(GanttEntityType.AsPlannedActivity)),
            DateLabelRow(RowTwoId, EntityTypeCatalog.Format(GanttEntityType.AsBuiltMilestone)),
            DateLabelRow(RowThreeId, EntityTypeCatalog.Format(GanttEntityType.Delineator)));

    private static IWorksheetProtectionGuard NotProtected()
    {
        var mock = new Mock<IWorksheetProtectionGuard>();
        mock.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        mock.Setup(g => g.QueryTarget(It.IsAny<object>())).Returns(ProtectionGuardOutcome.NotProtected);
        return mock.Object;
    }

    private static System.Runtime.InteropServices.COMException RefusedComException()
    {
        var exception = (System.Runtime.InteropServices.COMException?)
            System.Runtime.InteropServices.Marshal.GetExceptionForHR(unchecked((int)0x800A03EC));
        return exception ?? throw new InvalidOperationException(
            "The runtime did not produce a COMException for HRESULT 0x800A03EC.");
    }

    [Fact]
    public void SetDateLabelPosition_writes_exactly_the_date_label_cell_of_the_selected_row()
    {
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 2);

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPosition(GanttDateLabelPosition.Finish);

        Assert.True(outcome.Succeeded);
        (int Row, int Column, object? Value) write = Assert.Single(graph.WrittenCells);
        Assert.Equal(1, write.Row);
        Assert.Equal(TableIndexOf("DateLabelPosition"), write.Column);
        Assert.Equal(nameof(GanttDateLabelPosition.Finish), write.Value);
    }

    [Fact]
    public void SetDateLabelPosition_clears_the_cell_for_a_null_choice()
    {
        // R5.14 D3: a null choice (the Auto ribbon item) writes a
        // blank cell, which resolves to Auto at render time.
        var matrix = Matrix(
            DateLabelRow(RowOneId, EntityTypeCatalog.Format(GanttEntityType.AsPlannedActivity), "Finish"));
        Graph graph = SeededGraph(matrix, activeRow: 2);

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPosition(position: null);

        Assert.True(outcome.Succeeded);
        (int Row, int Column, object? Value) write = Assert.Single(graph.WrittenCells);
        Assert.Equal(1, write.Row);
        Assert.Equal(TableIndexOf("DateLabelPosition"), write.Column);
        Assert.Null(write.Value);
    }

    [Fact]
    public void SetDateLabelPosition_writes_a_start_choice_for_a_start_only_type()
    {
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 2);

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPosition(GanttDateLabelPosition.Start);

        Assert.True(outcome.Succeeded);
        (int, int, object? Value) write = Assert.Single(graph.WrittenCells);
        Assert.Equal(nameof(GanttDateLabelPosition.Start), write.Value);
    }

    [Fact]
    public void SetDateLabelPosition_writes_a_none_choice_for_a_span_type()
    {
        // R5.14 D5 (second owner ruling 2026-10-14): None names
        // no date labels; a span admits it, so exactly the
        // selected row's cell is written with the enum name.
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 2);

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPosition(GanttDateLabelPosition.None);

        Assert.True(outcome.Succeeded);
        (int Row, int Column, object? Value) write = Assert.Single(graph.WrittenCells);
        Assert.Equal(1, write.Row);
        Assert.Equal(TableIndexOf("DateLabelPosition"), write.Column);
        Assert.Equal(nameof(GanttDateLabelPosition.None), write.Value);
    }

    [Fact]
    public void SetDateLabelPosition_writes_a_none_choice_for_a_start_only_type()
    {
        // A milestone is StartOnly and still admits None (no
        // date labels), so the choice writes its own row's cell.
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 3);

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPosition(GanttDateLabelPosition.None);

        Assert.True(outcome.Succeeded);
        (int, int, object? Value) write = Assert.Single(graph.WrittenCells);
        Assert.Equal(nameof(GanttDateLabelPosition.None), write.Value);
    }

    [Fact]
    public void SetDateLabelPosition_refuses_a_none_choice_for_a_dateless_type()
    {
        // None is a date-label choice like any other, so the
        // type-only capability set (D4) governs it: a Splitter
        // names no choice at all, so the control would be
        // disabled and the write is refused before any cell.
        var matrix = Matrix(
            DateLabelRow(RowOneId, EntityTypeCatalog.Format(GanttEntityType.Splitter)));
        Graph graph = SeededGraph(matrix, activeRow: 2);

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPosition(GanttDateLabelPosition.None);

        Assert.False(outcome.Succeeded);
        Assert.Equal(DateLabelMutationRefusalReason.DateLabelNotAllowedForType, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetDateLabelPosition_writes_a_finish_choice_for_custom_activity_type_only()
    {
        // R5.14 D4: the capability is type-only. Custom Activity is a
        // StartFinish span, so Finish is allowed WITHOUT any style
        // registry read — there is no style-related refusal here.
        var matrix = Matrix(
            DateLabelRow(RowOneId, EntityTypeCatalog.Format(GanttEntityType.CustomActivity)));
        Graph graph = SeededGraph(matrix, activeRow: 2);

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPosition(GanttDateLabelPosition.Finish);

        Assert.True(outcome.Succeeded);
        (int, int, object? Value) write = Assert.Single(graph.WrittenCells);
        Assert.Equal(nameof(GanttDateLabelPosition.Finish), write.Value);
    }

    [Fact]
    public void SetDateLabelPosition_refuses_a_finish_choice_for_a_start_only_type()
    {
        // A milestone is StartOnly: it reads Start as its single date,
        // so Finish is not a permitted date-label choice (R5.14 D4).
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 3);

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPosition(GanttDateLabelPosition.Finish);

        Assert.False(outcome.Succeeded);
        Assert.Equal(DateLabelMutationRefusalReason.DateLabelNotAllowedForType, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetDateLabelPosition_refuses_a_choice_for_a_dateless_type()
    {
        // A Splitter is dateless (EntityDateMode.None): it names no
        // date-label choice at all, so the control would be disabled.
        var matrix = Matrix(
            DateLabelRow(RowOneId, EntityTypeCatalog.Format(GanttEntityType.Splitter)));
        Graph graph = SeededGraph(matrix, activeRow: 2);

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPosition(GanttDateLabelPosition.Start);

        Assert.False(outcome.Succeeded);
        Assert.Equal(DateLabelMutationRefusalReason.DateLabelNotAllowedForType, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetDateLabelPosition_refuses_an_unknown_type_before_any_write()
    {
        var matrix = Matrix(DateLabelRow(RowOneId, "Widget"));
        Graph graph = SeededGraph(matrix, activeRow: 2);

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPosition(GanttDateLabelPosition.Finish);

        Assert.False(outcome.Succeeded);
        Assert.Equal(DateLabelMutationRefusalReason.TypeUnknown, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetDateLabelPosition_refuses_a_malformed_id_before_any_write()
    {
        var matrix = Matrix(DateLabelRow(string.Empty, EntityTypeCatalog.Format(GanttEntityType.AsPlannedActivity)));
        Graph graph = SeededGraph(matrix, activeRow: 2);

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPosition(GanttDateLabelPosition.Finish);

        Assert.False(outcome.Succeeded);
        Assert.Equal(DateLabelMutationRefusalReason.SelectionUnknown, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Theory]
    [InlineData(1)] // the header row
    [InlineData(5)] // below the body
    public void SetDateLabelPosition_refuses_a_selection_outside_the_table_body(int activeRow)
    {
        Graph graph = SeededGraph(ThreeRowRun(), activeRow);

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPosition(GanttDateLabelPosition.Finish);

        Assert.False(outcome.Succeeded);
        Assert.Equal(DateLabelMutationRefusalReason.SelectionOutsideTable, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetDateLabelPosition_runs_the_protection_guard_before_any_workbook_access()
    {
        // Positive guard-first test: the application is NULL, so reaching
        // the workbook would report NoActiveWorkbook. Getting TargetProtected
        // back proves the guard ran before any workbook/application access.
        var mock = new Mock<IWorksheetProtectionGuard>();
        mock.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.SheetProtected);
        var graph = new Graph { BodyMatrix = ThreeRowRun(), RowCount = 3, ActiveRow = 2 };
        var mutator = new TestableDateLabelMutator(application: null, mock.Object, graph);

        DateLabelMutationOutcome outcome = mutator.SetDateLabelPosition(GanttDateLabelPosition.Finish);

        Assert.False(outcome.Succeeded);
        Assert.Equal(DateLabelMutationRefusalReason.TargetProtected, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetDateLabelPosition_refuses_when_the_table_is_missing()
    {
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 2);
        graph.TablePresent = false;

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPosition(GanttDateLabelPosition.Finish);

        Assert.False(outcome.Succeeded);
        Assert.Equal(DateLabelMutationRefusalReason.TableMissing, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetDateLabelPosition_reports_a_host_write_failure_without_claiming_success()
    {
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 2);
        graph.ThrowOnWrite = true;

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPosition(GanttDateLabelPosition.Finish);

        Assert.False(outcome.Succeeded);
        Assert.Equal(DateLabelMutationRefusalReason.WriteFailed, outcome.Refusal);
        Assert.Equal(0, outcome.CellsWritten);
    }

    [Fact]
    public void SetDateLabelPosition_writes_only_the_selected_parent_cell_when_children_follow_it()
    {
        // Owner ruling 2026-10-14: a collapsed group's parent edits ITSELF
        // only. The children's cells (rows 2 and 3) are never written, and
        // no hidden-flag or any other column is touched.
        var matrix = Matrix(
            DateLabelRow(RowOneId, EntityTypeCatalog.Format(GanttEntityType.AsPlannedActivity), dateLabelPosition: "Start"),
            DateLabelRow(RowTwoId, EntityTypeCatalog.Format(GanttEntityType.AsPlannedActivity), dateLabelPosition: "Start"),
            DateLabelRow(RowThreeId, EntityTypeCatalog.Format(GanttEntityType.AsPlannedActivity), dateLabelPosition: "Start"));
        Graph graph = SeededGraph(matrix, activeRow: 2);

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPosition(GanttDateLabelPosition.Finish);

        Assert.True(outcome.Succeeded);
        (int Row, int _, object? Value) write = Assert.Single(graph.WrittenCells);
        Assert.Equal(1, write.Row);
        Assert.Equal(nameof(GanttDateLabelPosition.Finish), write.Value);
    }

    [Fact]
    public void SetDateLabelPosition_refuses_when_the_table_predates_the_column()
    {
        // R5.14 defect regression: a workbook saved before R5.14 has
        // no DateLabelPosition column. The column is optional so the
        // table still reads, but a write has no cell to write: the
        // mutator refuses with the typed TableMissing refusal (whose
        // message tells the user to run Initialise, which materialises
        // the column) instead of aiming the write at a non-existent
        // column index.
        IReadOnlyList<string> legacyColumns = GanttTableSchema.Default.Columns
            .Select(static c => c.Name)
            .Where(static name => name != "DateLabelPosition")
            .ToArray();
        var graph = new Graph(legacyColumns)
        {
            BodyMatrix = LegacyMatrix(legacyColumns,
                LegacyRow(legacyColumns, RowOneId, EntityTypeCatalog.Format(GanttEntityType.AsPlannedActivity))),
            RowCount = 1,
            ActiveRow = 2,
        };

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPosition(GanttDateLabelPosition.Start);

        Assert.Equal(DateLabelMutationOutcome.Refused(DateLabelMutationRefusalReason.TableMissing, 1), outcome);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetDateLabelPlacement_writes_exactly_the_date_label_cell_of_the_delineator_row()
    {
        // Owner ruling 2026-10-14: a delineator's date label is
        // positionable. The corner name lands in the SAME
        // DateLabelPosition cell — the vocabulary is
        // type-dependent, not schema-dependent — and only the
        // selected row's cell is written.
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 4);

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPlacement(GanttLabelPosition.TopLeft);

        Assert.True(outcome.Succeeded);
        (int Row, int Column, object? Value) write = Assert.Single(graph.WrittenCells);
        Assert.Equal(3, write.Row);
        Assert.Equal(TableIndexOf("DateLabelPosition"), write.Column);
        Assert.Equal(nameof(GanttLabelPosition.TopLeft), write.Value);
    }

    [Fact]
    public void SetDateLabelPlacement_clears_the_cell_for_a_null_placement()
    {
        // The Auto item of the placement dropdown: a null
        // placement writes a blank cell, which resolves to
        // Auto (the first free corner) at render time.
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 4);

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPlacement(placement: null);

        Assert.True(outcome.Succeeded);
        (int Row, int Column, object? Value) write = Assert.Single(graph.WrittenCells);
        Assert.Equal(3, write.Row);
        Assert.Equal(TableIndexOf("DateLabelPosition"), write.Column);
        Assert.Null(write.Value);
    }

    [Fact]
    public void SetDateLabelPlacement_refuses_a_placement_for_a_span_row()
    {
        // The mirror of the which-date refusal: a span's date
        // label selects WHICH date shows, not where it sits,
        // so a corner write is refused before any cell.
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 2);

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPlacement(GanttLabelPosition.TopLeft);

        Assert.False(outcome.Succeeded);
        Assert.Equal(DateLabelMutationRefusalReason.DateLabelNotAllowedForType, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetDateLabelPlacement_refuses_a_corner_outside_the_delineator_set()
    {
        // Inside is a description-label position but not one
        // of the six corners the positionable date-label set
        // names, so the type-only capability set refuses it.
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 4);

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPlacement(GanttLabelPosition.Inside);

        Assert.False(outcome.Succeeded);
        Assert.Equal(DateLabelMutationRefusalReason.DateLabelNotAllowedForType, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetDateLabelPosition_refuses_every_which_date_choice_for_a_delineator()
    {
        // The vocabulary split from the which-date side: a
        // delineator's which-date set is empty (its Start is
        // its single date), so every which-date choice is
        // refused and the dropdown offers the corners instead.
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 4);

        DateLabelMutationOutcome outcome = graph.Build(NotProtected())
            .SetDateLabelPosition(GanttDateLabelPosition.Finish);

        Assert.False(outcome.Succeeded);
        Assert.Equal(DateLabelMutationRefusalReason.DateLabelNotAllowedForType, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetDateLabelPlacement_runs_the_protection_guard_before_any_workbook_access()
    {
        // Positive guard-first test for the placement entry
        // point (ADR-0008 D4): the application is NULL, so
        // reaching the workbook would report NoActiveWorkbook.
        // Getting TargetProtected back proves the guard ran
        // before any workbook/application access.
        var mock = new Mock<IWorksheetProtectionGuard>();
        mock.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.SheetProtected);
        var graph = new Graph { BodyMatrix = ThreeRowRun(), RowCount = 3, ActiveRow = 4 };
        var mutator = new TestableDateLabelMutator(application: null, mock.Object, graph);

        DateLabelMutationOutcome outcome = mutator.SetDateLabelPlacement(GanttLabelPosition.TopLeft);

        Assert.False(outcome.Succeeded);
        Assert.Equal(DateLabelMutationRefusalReason.TargetProtected, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }
}

