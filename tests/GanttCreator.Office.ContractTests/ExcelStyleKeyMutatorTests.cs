using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for the R5.6a style-key mutation adapter.
/// </summary>
/// <remarks>
/// The tests drive the adapter through the same fake-COM-graph seam
/// discipline as the R5.5 label-position mutator tests: every COM
/// parameterised property is an <c>internal virtual</c> seam the testable
/// subclass overrides, so the full guard-first pipeline runs without a
/// host. They pin the adapter's write set — EXACTLY ONE StyleKey cell on
/// the selected style-key-type row — and every typed refusal before the
/// write, including the new validator's positive tests (non-style type,
/// blank key, unknown key).
/// </remarks>
public class ExcelStyleKeyMutatorTests
{
    private const string RowOneId = "G-0000000000000000000000000000000a";
    private const string RowTwoId = "G-0000000000000000000000000000000b";

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

        public TestableStyleKeyMutator Build(
            IWorksheetProtectionGuard guard,
            IConfigCatalogueReader? reader = null)
        {
            var mutator = new TestableStyleKeyMutator(
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

    private sealed class TestableStyleKeyMutator(
        Excel.Application? application,
        IWorksheetProtectionGuard guard,
        IConfigCatalogueReader reader,
        Graph graph) : ExcelStyleKeyMutator(application, guard, reader)
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

    private static object?[] Row(string id, string type, string? styleKey = null)
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

    private static IWorksheetProtectionGuard NotProtected()
    {
        var mock = new Mock<IWorksheetProtectionGuard>();
        mock.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        mock.Setup(g => g.QueryTarget(It.IsAny<object>())).Returns(ProtectionGuardOutcome.NotProtected);
        return mock.Object;
    }

    private static IWorksheetProtectionGuard Protected()
    {
        var mock = new Mock<IWorksheetProtectionGuard>();
        mock.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.SheetProtected);
        return mock.Object;
    }

    private static ConfigReadOutcome RegistryWith(string key)
    {
        var registry = new GanttStyleRegistry(
        [
            new GanttStyleDefinition(
                key,
                new HashSet<GanttLabelPosition>([GanttLabelPosition.Auto]),
                EntityColourCapability.StyleDefined),
        ]);
        return ConfigReadOutcome.Ok("wb", new Dictionary<string, string>(), registry, new Dictionary<string, double>());
    }

    private static System.Runtime.InteropServices.COMException RefusedComException()
    {
        var exception = (System.Runtime.InteropServices.COMException?)
            System.Runtime.InteropServices.Marshal.GetExceptionForHR(unchecked((int)0x800A03EC));
        return exception ?? throw new InvalidOperationException(
            "The runtime did not produce a COMException for HRESULT 0x800A03EC.");
    }

    [Fact]
    public void SetStyleKey_refuses_when_the_worksheet_is_protected()
    {
        var matrix = Matrix(Row(RowOneId, EntityTypeCatalog.Format(GanttEntityType.CustomActivity)));
        Graph graph = SeededGraph(matrix, activeRow: 2);

        StyleKeyMutationOutcome outcome = graph.Build(Protected())
            .SetStyleKey("MyStyle");

        Assert.False(outcome.Succeeded);
        Assert.Equal(StyleKeyMutationRefusalReason.TargetProtected, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetStyleKey_refuses_without_an_active_workbook()
    {
        var matrix = Matrix(Row(RowOneId, EntityTypeCatalog.Format(GanttEntityType.CustomActivity)));
        Graph graph = SeededGraph(matrix, activeRow: 2);
        graph.Application.SetupGet(a => a.ActiveWorkbook).Returns((Excel.Workbook)null!);

        StyleKeyMutationOutcome outcome = graph.Build(NotProtected())
            .SetStyleKey("MyStyle");

        Assert.False(outcome.Succeeded);
        Assert.Equal(StyleKeyMutationRefusalReason.NoActiveWorkbook, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetStyleKey_refuses_when_the_table_is_missing()
    {
        var matrix = Matrix(Row(RowOneId, EntityTypeCatalog.Format(GanttEntityType.CustomActivity)));
        Graph graph = SeededGraph(matrix, activeRow: 2);
        graph.TablePresent = false;

        StyleKeyMutationOutcome outcome = graph.Build(NotProtected())
            .SetStyleKey("MyStyle");

        Assert.False(outcome.Succeeded);
        Assert.Equal(StyleKeyMutationRefusalReason.TableMissing, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetStyleKey_refuses_a_selection_outside_the_table_body()
    {
        var matrix = Matrix(Row(RowOneId, EntityTypeCatalog.Format(GanttEntityType.CustomActivity)));
        Graph graph = SeededGraph(matrix, activeRow: 9);

        StyleKeyMutationOutcome outcome = graph.Build(NotProtected())
            .SetStyleKey("MyStyle");

        Assert.False(outcome.Succeeded);
        Assert.Equal(StyleKeyMutationRefusalReason.SelectionOutsideTable, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetStyleKey_refuses_a_malformed_row_id()
    {
        var matrix = Matrix(Row("not-a-row-id", EntityTypeCatalog.Format(GanttEntityType.CustomActivity)));
        Graph graph = SeededGraph(matrix, activeRow: 2);

        StyleKeyMutationOutcome outcome = graph.Build(NotProtected())
            .SetStyleKey("MyStyle");

        Assert.False(outcome.Succeeded);
        Assert.Equal(StyleKeyMutationRefusalReason.SelectionUnknown, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetStyleKey_refuses_an_unknown_type()
    {
        var matrix = Matrix(Row(RowOneId, "Not A Type"));
        Graph graph = SeededGraph(matrix, activeRow: 2);

        StyleKeyMutationOutcome outcome = graph.Build(NotProtected())
            .SetStyleKey("MyStyle");

        Assert.False(outcome.Succeeded);
        Assert.Equal(StyleKeyMutationRefusalReason.TypeUnknown, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetStyleKey_refuses_a_non_style_type_before_any_write()
    {
        // Positive test for the TypeNotCustomStyled validator: a plain
        // activity carries no StyleKey, so the write is refused even though
        // the requested key exists in the registry.
        var matrix = Matrix(Row(RowOneId, EntityTypeCatalog.Format(GanttEntityType.AsPlannedActivity)));
        Graph graph = SeededGraph(matrix, activeRow: 2);
        graph.Catalogue = RegistryWith("MyStyle");

        StyleKeyMutationOutcome outcome = graph.Build(NotProtected())
            .SetStyleKey("MyStyle");

        Assert.False(outcome.Succeeded);
        Assert.Equal(StyleKeyMutationRefusalReason.TypeNotCustomStyled, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetStyleKey_clears_a_blank_key_on_a_style_key_row()
    {
        // The dropdown's <c>None</c> item clears the cell: a blank key on a
        // style-key-type row writes empty and succeeds rather than refusing.
        var matrix = Matrix(Row(RowOneId, EntityTypeCatalog.Format(GanttEntityType.CustomActivity)));
        Graph graph = SeededGraph(matrix, activeRow: 2);
        graph.Catalogue = RegistryWith("MyStyle");

        StyleKeyMutationOutcome outcome = graph.Build(NotProtected())
            .SetStyleKey("   ");

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.CellsWritten);
        (int RowIndex, int ColumnIndex, object? Value) write = Assert.Single(graph.WrittenCells);
        Assert.Equal(1, write.RowIndex);
        Assert.Equal("StyleKey", GanttTableSchema.Default.Columns[write.ColumnIndex - 1].Name);
        Assert.Equal(string.Empty, write.Value);
    }

    [Fact]
    public void SetStyleKey_refuses_an_unknown_style_key()
    {
        // Positive test for the StyleUnknown validator: the key is not in
        // the configuration registry, so nothing is written.
        var matrix = Matrix(Row(RowOneId, EntityTypeCatalog.Format(GanttEntityType.CustomActivity)));
        Graph graph = SeededGraph(matrix, activeRow: 2);
        graph.Catalogue = RegistryWith("MyStyle");

        StyleKeyMutationOutcome outcome = graph.Build(NotProtected())
            .SetStyleKey("Missing");

        Assert.False(outcome.Succeeded);
        Assert.Equal(StyleKeyMutationRefusalReason.StyleUnknown, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetStyleKey_refuses_when_the_catalogue_read_fails()
    {
        var matrix = Matrix(Row(RowOneId, EntityTypeCatalog.Format(GanttEntityType.CustomActivity)));
        Graph graph = SeededGraph(matrix, activeRow: 2);
        graph.Catalogue = ConfigReadOutcome.Refused(ConfigReadRefusalReason.ConfigSheetMissing);

        StyleKeyMutationOutcome outcome = graph.Build(NotProtected())
            .SetStyleKey("MyStyle");

        Assert.False(outcome.Succeeded);
        Assert.Equal(StyleKeyMutationRefusalReason.TableInvalid, outcome.Refusal);
        Assert.Empty(graph.WrittenCells);
    }

    [Fact]
    public void SetStyleKey_writes_exactly_the_style_key_cell_of_the_selected_custom_activity()
    {
        var matrix = Matrix(Row(RowOneId, EntityTypeCatalog.Format(GanttEntityType.CustomActivity)));
        Graph graph = SeededGraph(matrix, activeRow: 2);
        graph.Catalogue = RegistryWith("MyStyle");

        StyleKeyMutationOutcome outcome = graph.Build(NotProtected())
            .SetStyleKey("MyStyle");

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.CellsWritten);
        (int Row, int Column, object? Value) write = Assert.Single(graph.WrittenCells);
        Assert.Equal(1, write.Row);
        Assert.Equal("StyleKey", GanttTableSchema.Default.Columns[write.Column - 1].Name);
        Assert.Equal("MyStyle", write.Value);
    }

    [Fact]
    public void SetStyleKey_writes_the_style_key_cell_of_a_custom_milestone()
    {
        // R5.6a: Custom Milestone is a style-key type, so the same single-cell
        // write applies to it exactly as to a Custom Activity.
        var matrix = Matrix(Row(RowOneId, EntityTypeCatalog.Format(GanttEntityType.CustomMilestone)));
        Graph graph = SeededGraph(matrix, activeRow: 2);
        graph.Catalogue = RegistryWith("MyStyle");

        StyleKeyMutationOutcome outcome = graph.Build(NotProtected())
            .SetStyleKey("MyStyle");

        Assert.True(outcome.Succeeded);
        (int RowIndex, int ColumnIndex, object? Payload) write = Assert.Single(graph.WrittenCells);
        Assert.Equal("StyleKey", GanttTableSchema.Default.Columns[write.ColumnIndex - 1].Name);
        Assert.Equal("MyStyle", write.Payload);
    }

    [Fact]
    public void SetStyleKey_trims_the_key_before_writing()
    {
        var matrix = Matrix(Row(RowOneId, EntityTypeCatalog.Format(GanttEntityType.CustomActivity)));
        Graph graph = SeededGraph(matrix, activeRow: 2);
        graph.Catalogue = RegistryWith("MyStyle");

        StyleKeyMutationOutcome outcome = graph.Build(NotProtected())
            .SetStyleKey("  MyStyle  ");

        Assert.True(outcome.Succeeded);
        (int RowIndex, int ColumnIndex, object? Payload) write = Assert.Single(graph.WrittenCells);
        Assert.Equal("MyStyle", write.Payload);
    }

    [Fact]
    public void SetStyleKey_writes_only_the_selected_row_when_a_second_row_exists()
    {
        var matrix = Matrix(
            Row(RowOneId, EntityTypeCatalog.Format(GanttEntityType.CustomActivity), "First"),
            Row(RowTwoId, EntityTypeCatalog.Format(GanttEntityType.CustomActivity), "Second"));
        Graph graph = SeededGraph(matrix, activeRow: 3);
        graph.Catalogue = RegistryWith("MyStyle");

        StyleKeyMutationOutcome outcome = graph.Build(NotProtected())
            .SetStyleKey("MyStyle");

        Assert.True(outcome.Succeeded);
        (int RowIndex, int ColumnIndex, object? Payload) write = Assert.Single(graph.WrittenCells);
        Assert.Equal(2, write.RowIndex);
    }

    [Fact]
    public void SetStyleKey_reports_a_host_write_failure_without_claiming_success()
    {
        var matrix = Matrix(Row(RowOneId, EntityTypeCatalog.Format(GanttEntityType.CustomActivity)));
        Graph graph = SeededGraph(matrix, activeRow: 2);
        graph.Catalogue = RegistryWith("MyStyle");
        graph.ThrowOnWrite = true;

        StyleKeyMutationOutcome outcome = graph.Build(NotProtected())
            .SetStyleKey("MyStyle");

        Assert.False(outcome.Succeeded);
        Assert.Equal(StyleKeyMutationRefusalReason.WriteFailed, outcome.Refusal);
    }
}
