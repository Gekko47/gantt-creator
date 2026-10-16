using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for the R5.5 label-position mutation adapter.
/// </summary>
/// <remarks>
/// The tests drive the adapter through the same fake-COM-graph seam
/// discipline as the R5.3 parent/child mutator tests: every COM
/// parameterised property is an <c>internal virtual</c> seam the testable
/// subclass overrides, so the full guard-first pipeline runs without a
/// host. They pin the adapter's write set — EXACTLY ONE LabelPosition cell
/// on the selected row — and every typed refusal before the write.
/// </remarks>
public class ExcelLabelPositionMutatorTests
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
        public Mock<Excel.ListColumn>[] Columns { get; } =
            Enumerable.Range(0, GanttTableSchema.Default.Columns.Count)
                .Select(_ => new Mock<Excel.ListColumn>()).ToArray();
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

        public ConfigReadOutcome Catalogue { get; set; } =
            ConfigReadOutcome.Ok("wb", new Dictionary<string, string>(), GanttStyleRegistry.Empty, new Dictionary<string, double>());

        public Graph()
        {
            Application.SetupGet(a => a.ActiveWorkbook).Returns(Workbook.Object);
            Workbook.SetupGet(w => w.Sheets).Returns(Sheets.Object);
            Sheets.SetupGet(s => s.Count).Returns(1);
            Worksheet.SetupGet(w => w.ListObjects).Returns(ListObjects.Object);
            ListObjects.SetupGet(lo => lo.Count).Returns(() => TablePresent ? 1 : 0);
            Table.SetupGet(t => t.Name).Returns("tblGanttData");
            Table.SetupGet(t => t.ListColumns).Returns(ListColumns.Object);
            ListColumns.SetupGet(c => c.Count).Returns(GanttTableSchema.Default.Columns.Count);
            for (var i = 0; i < Columns.Length; i++)
            {
                Columns[i].SetupGet(c => c.Name).Returns(GanttTableSchema.Default.Columns[i].Name);
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

        /// <summary>Resolves a ListRow proxy back to its one-based row index.</summary>
        public Excel.Range RowRangeOfListRow(Excel.ListRow row) => RowRange(_listRowIndexes[row]);

        /// <summary>Resolves a row-range proxy plus a table column to its cell proxy.</summary>
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

        public TestableLabelMutator Build(
            IWorksheetProtectionGuard guard,
            IConfigCatalogueReader? reader = null)
        {
            var mutator = new TestableLabelMutator(
                Application.Object,
                guard,
                reader ?? new StubCatalogueReader(() => Catalogue),
                this);
            return mutator;
        }
    }

    private sealed class StubCatalogueReader(Func<ConfigReadOutcome> read) : IConfigCatalogueReader
    {
        private readonly Func<ConfigReadOutcome> _read = read;

        public ConfigReadOutcome Read() => _read();
    }

    private sealed class TestableLabelMutator(
        Excel.Application? application,
        IWorksheetProtectionGuard guard,
        IConfigCatalogueReader reader,
        Graph graph) : ExcelLabelPositionMutator(application, guard, reader)
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

        internal override ConfigReadOutcome ReadCatalogue() => _graph.Catalogue;
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

    private static object?[] LabelRow(string id, string type, string? styleKey = null, string? labelPosition = null)
    {
        IReadOnlyList<GanttTableColumn> columns = GanttTableSchema.Default.Columns;
        var row = new object?[columns.Count];
        for (var index = 0; index < columns.Count; index++)
        {
            row[index] = columns[index].Name switch
            {
                "Id" => id,
                "Type" => type,
                "StyleKey" => styleKey,
                "LabelPosition" => labelPosition,
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
            LabelRow(RowOneId, EntityTypeCatalog.Format(GanttEntityType.AsPlannedActivity)),
            LabelRow(RowTwoId, EntityTypeCatalog.Format(GanttEntityType.AsPlannedMilestone)),
            LabelRow(RowThreeId, EntityTypeCatalog.Format(GanttEntityType.Delineator)));

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
    public void SetLabelPosition_writes_exactly_the_label_cell_of_the_selected_row()
    {
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 2);

        LabelPositionMutationOutcome outcome = graph.Build(NotProtected())
            .SetLabelPosition(GanttLabelPosition.Right);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.CellsWritten);
        Assert.Equal(1, outcome.SelectedRowNumber);
        (int Row, int Column, object? Value) write = Assert.Single(graph.WrittenCells);
        Assert.Equal(1, write.Row);
        Assert.Equal(TableIndexOf("LabelPosition"), write.Column);
        Assert.Equal(nameof(GanttLabelPosition.Right), write.Value);
    }

    [Fact]
    public void SetLabelPosition_default_clears_the_cell_to_blank()
    {
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 3);

        LabelPositionMutationOutcome outcome = graph.Build(NotProtected())
            .SetLabelPosition(null);

        Assert.True(outcome.Succeeded);
        Assert.Equal(2, outcome.SelectedRowNumber);
        (int Row, int Column, object? Value) write = Assert.Single(graph.WrittenCells);
        Assert.Equal(2, write.Row);
        Assert.Equal(TableIndexOf("LabelPosition"), write.Column);
        Assert.Null(write.Value);
    }

    [Theory]
    [InlineData(GanttEntityType.Spacer, GanttLabelPosition.Left)]
    [InlineData(GanttEntityType.CriticalInterval, GanttLabelPosition.Right)]
    [InlineData(GanttEntityType.AsPlannedMilestone, GanttLabelPosition.Inside)]
    [InlineData(GanttEntityType.Delineator, GanttLabelPosition.Left)]
    public void SetLabelPosition_refuses_a_position_not_allowed_for_the_type_before_any_write(
        GanttEntityType type, GanttLabelPosition position)
    {
        // Positive tests: a bad input MUST reach the refusal path (AGENTS.md).
        var matrix = Matrix(LabelRow(RowOneId, EntityTypeCatalog.Format(type)));
        Graph graph = SeededGraph(matrix, activeRow: 2);

        LabelPositionMutationOutcome outcome = graph.Build(NotProtected())
            .SetLabelPosition(position);

        Assert.False(outcome.Succeeded);
        Assert.Equal(LabelPositionMutationRefusalReason.PositionNotAllowedForType, outcome.Refusal);
        Assert.Equal(0, outcome.CellsWritten);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetLabelPosition_refuses_an_unknown_type_before_any_write()
    {
        var matrix = Matrix(LabelRow(RowOneId, "Widget"));
        Graph graph = SeededGraph(matrix, activeRow: 2);

        LabelPositionMutationOutcome outcome = graph.Build(NotProtected())
            .SetLabelPosition(GanttLabelPosition.Left);

        Assert.False(outcome.Succeeded);
        Assert.Equal(LabelPositionMutationRefusalReason.TypeUnknown, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetLabelPosition_refuses_a_malformed_id_before_any_write()
    {
        var matrix = Matrix(LabelRow(string.Empty, EntityTypeCatalog.Format(GanttEntityType.AsPlannedActivity)));
        Graph graph = SeededGraph(matrix, activeRow: 2);

        LabelPositionMutationOutcome outcome = graph.Build(NotProtected())
            .SetLabelPosition(GanttLabelPosition.Left);

        Assert.False(outcome.Succeeded);
        Assert.Equal(LabelPositionMutationRefusalReason.SelectionUnknown, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Theory]
    [InlineData(1)] // the header row
    [InlineData(5)] // below the body
    public void SetLabelPosition_refuses_a_selection_outside_the_table_body(int activeRow)
    {
        Graph graph = SeededGraph(ThreeRowRun(), activeRow);

        LabelPositionMutationOutcome outcome = graph.Build(NotProtected())
            .SetLabelPosition(GanttLabelPosition.Left);

        Assert.False(outcome.Succeeded);
        Assert.Equal(LabelPositionMutationRefusalReason.SelectionOutsideTable, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetLabelPosition_runs_the_protection_guard_before_any_workbook_access()
    {
        // Positive guard-first test: the application is NULL, so reaching the
        // workbook would report NoActiveWorkbook. Getting TargetProtected back
        // proves the guard ran before any workbook/application access.
        var mock = new Mock<IWorksheetProtectionGuard>();
        mock.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.SheetProtected);
        var graph = new Graph { BodyMatrix = ThreeRowRun(), RowCount = 3, ActiveRow = 2 };
        var mutator = new TestableLabelMutator(
            application: null,
            mock.Object,
            new StubCatalogueReader(() => graph.Catalogue),
            graph);

        LabelPositionMutationOutcome outcome = mutator.SetLabelPosition(GanttLabelPosition.Left);

        Assert.False(outcome.Succeeded);
        Assert.Equal(LabelPositionMutationRefusalReason.TargetProtected, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetLabelPosition_refuses_when_the_table_is_missing()
    {
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 2);
        graph.TablePresent = false;

        LabelPositionMutationOutcome outcome = graph.Build(NotProtected())
            .SetLabelPosition(GanttLabelPosition.Left);

        Assert.False(outcome.Succeeded);
        Assert.Equal(LabelPositionMutationRefusalReason.TableMissing, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetLabelPosition_reports_a_host_write_failure_without_claiming_success()
    {
        Graph graph = SeededGraph(ThreeRowRun(), activeRow: 2);
        graph.ThrowOnWrite = true;

        LabelPositionMutationOutcome outcome = graph.Build(NotProtected())
            .SetLabelPosition(GanttLabelPosition.Left);

        Assert.False(outcome.Succeeded);
        Assert.Equal(LabelPositionMutationRefusalReason.WriteFailed, outcome.Refusal);
        Assert.Equal(0, outcome.CellsWritten);
    }

    [Fact]
    public void SetLabelPosition_writes_an_allowed_position_for_custom_activity_from_its_style()
    {
        var registry = new GanttStyleRegistry(
        [
            new GanttStyleDefinition(
                "MyStyle",
                new HashSet<GanttLabelPosition>(
                    [GanttLabelPosition.Left, GanttLabelPosition.Right, GanttLabelPosition.None]),
                EntityColourCapability.StyleDefined),
        ]);
        var matrix = Matrix(
            LabelRow(RowOneId, EntityTypeCatalog.Format(GanttEntityType.CustomActivity), "MyStyle"));
        Graph graph = SeededGraph(matrix, activeRow: 2);
        graph.Catalogue = ConfigReadOutcome.Ok(
            "wb", new Dictionary<string, string>(), registry, new Dictionary<string, double>());

        LabelPositionMutationOutcome outcome = graph.Build(NotProtected())
            .SetLabelPosition(GanttLabelPosition.Left);

        Assert.True(outcome.Succeeded);
        (int Row, int Column, object? Value) write = Assert.Single(graph.WrittenCells);
        Assert.Equal(nameof(GanttLabelPosition.Left), write.Value);
    }

    [Fact]
    public void SetLabelPosition_refuses_an_unknown_style_key_for_custom_activity()
    {
        var matrix = Matrix(
            LabelRow(RowOneId, EntityTypeCatalog.Format(GanttEntityType.CustomActivity), "Missing"));
        Graph graph = SeededGraph(matrix, activeRow: 2);
        graph.Catalogue = ConfigReadOutcome.Ok(
            "wb", new Dictionary<string, string>(), GanttStyleRegistry.Empty, new Dictionary<string, double>());

        LabelPositionMutationOutcome outcome = graph.Build(NotProtected())
            .SetLabelPosition(GanttLabelPosition.Left);

        Assert.False(outcome.Succeeded);
        Assert.Equal(LabelPositionMutationRefusalReason.StyleUnknown, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetLabelPosition_refuses_custom_activity_when_the_catalogue_read_fails()
    {
        var matrix = Matrix(
            LabelRow(RowOneId, EntityTypeCatalog.Format(GanttEntityType.CustomActivity), "MyStyle"));
        Graph graph = SeededGraph(matrix, activeRow: 2);
        graph.Catalogue = ConfigReadOutcome.Refused(ConfigReadRefusalReason.ConfigSheetMissing);

        LabelPositionMutationOutcome outcome = graph.Build(NotProtected())
            .SetLabelPosition(GanttLabelPosition.Left);

        Assert.False(outcome.Succeeded);
        Assert.Equal(LabelPositionMutationRefusalReason.TableInvalid, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetLabelPosition_writes_only_the_selected_parent_cell_when_children_follow_it()
    {
        // Owner ruling 2026-10-14: a collapsed group's parent edits ITSELF
        // only. The children's cells (rows 2 and 3) are never written, and no
        // hidden-flag or any other column is touched.
        var matrix = Matrix(
            LabelRow(RowOneId, EntityTypeCatalog.Format(GanttEntityType.AsPlannedActivity), labelPosition: "None"),
            LabelRow(RowTwoId, EntityTypeCatalog.Format(GanttEntityType.AsPlannedActivity), labelPosition: "Left"),
            LabelRow(RowThreeId, EntityTypeCatalog.Format(GanttEntityType.AsPlannedActivity), labelPosition: "Right"));
        Graph graph = SeededGraph(matrix, activeRow: 2);

        LabelPositionMutationOutcome outcome = graph.Build(NotProtected())
            .SetLabelPosition(GanttLabelPosition.Right);

        Assert.True(outcome.Succeeded);
        (int Row, int Column, object? Value) write = Assert.Single(graph.WrittenCells);
        Assert.Equal(1, write.Row);
        Assert.Equal(nameof(GanttLabelPosition.Right), write.Value);
    }
}
