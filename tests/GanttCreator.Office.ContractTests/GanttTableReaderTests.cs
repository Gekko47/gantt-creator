using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using Moq;

namespace GanttCreator.Office.ContractTests;

public class GanttTableReaderTests
{
    private sealed class ColumnGraph
    {
        public Mock<Excel.ListColumn> Column { get; } = new();

        public ColumnGraph(string name)
        {
            _ = Column.SetupGet(c => c.Name).Returns(name);
        }
    }

    private sealed class TableGraph
    {
        public Mock<Excel.ListObject> Table { get; } = new();
        public Mock<Excel.ListColumns> Columns { get; } = new();
        public Mock<Excel.Range> Body { get; } = new();
        public List<ColumnGraph> ColumnGraphs { get; } = new();

        public TableGraph(string name, IReadOnlyList<string> headers)
        {
            _ = Table.SetupGet(t => t.Name).Returns(name);
            _ = Table.SetupGet(t => t.ListColumns).Returns(Columns.Object);
            _ = Table.SetupGet(t => t.DataBodyRange).Returns(Body.Object);
            _ = Columns.SetupGet(c => c.Count).Returns(headers.Count);
            foreach (var header in headers)
            {
                ColumnGraphs.Add(new ColumnGraph(header));
            }
        }

        public Excel.ListColumn ColumnAt(int index) => ColumnGraphs[index - 1].Column.Object;

        public void WithNoBody()
        {
            _ = Table.SetupGet(t => t.DataBodyRange).Returns((Excel.Range)null!);
        }
    }

    private sealed class SheetGraph
    {
        public Mock<Excel.Worksheet> Worksheet { get; } = new();
        public Mock<Excel.ListObjects> ListObjects { get; } = new();
        public List<TableGraph> Tables { get; } = new();

        public SheetGraph(params TableGraph[] tables)
        {
            Tables.AddRange(tables);
            _ = Worksheet.SetupGet(w => w.ListObjects).Returns(ListObjects.Object);
            _ = ListObjects.SetupGet(l => l.Count).Returns(tables.Length);
        }
    }

    private sealed class TestableReader : ExcelGanttTableReader
    {
        public TestableReader(
            object? application,
            Func<Excel.Sheets, int, object>? sheetAt = null,
            Func<Excel.ListObjects, int, Excel.ListObject>? tableAt = null,
            Func<Excel.ListColumns, int, Excel.ListColumn>? columnAt = null,
            Func<Excel.Workbook, bool>? date1904 = null,
            Func<Excel.Range, object?>? bodyValues = null,
            IExcelDateSystemConverter? dateSystemConverter = null)
            : base(application, dateSystemConverter)
        {
            SheetAt = sheetAt ?? ((_, _) => throw new InvalidOperationException("should not be called"));
            TableAt = tableAt ?? ((_, _) => throw new InvalidOperationException("should not be called"));
            ColumnAt = columnAt ?? ((_, _) => throw new InvalidOperationException("should not be called"));
            Date1904 = date1904 ?? (_ => throw new InvalidOperationException("should not be called"));
            BodyValue = bodyValues ?? (_ => throw new InvalidOperationException("should not be called"));
        }

        public Func<Excel.Sheets, int, object> SheetAt { get; }
        public Func<Excel.ListObjects, int, Excel.ListObject> TableAt { get; }
        public Func<Excel.ListColumns, int, Excel.ListColumn> ColumnAt { get; }
        public Func<Excel.Workbook, bool> Date1904 { get; }
        public Func<Excel.Range, object?> BodyValue { get; }

        internal override bool GetDate1904(Excel.Workbook workbook) => Date1904(workbook);
        internal override object GetSheetAt(Excel.Sheets sheets, int index) => SheetAt(sheets, index);
        internal override Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index) => TableAt(listObjects, index);
        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index) => ColumnAt(columns, index);
        internal override object? GetBodyValues(Excel.Range body) => BodyValue(body);
    }

    private static List<string> SchemaHeaders() =>
        GanttTableSchema.Default.Columns.Select(c => c.Name).ToList();

    private static Array BodyMatrix(params object?[][] rows)
    {
        int rowCount = rows.Length;
        int colCount = GanttTableSchema.Default.Columns.Count;
        var matrix = Array.CreateInstance(typeof(object), [rowCount, colCount], [1, 1]);
        for (int row = 1; row <= rowCount; row++)
        {
            for (int column = 1; column <= colCount; column++)
            {
                matrix.SetValue(rows[row - 1][column - 1], row, column);
            }
        }

        return matrix;
    }

    private static object?[] FullRow(
        object? id = null,
        object? stackIndex = null,
        object? type = null,
        object? start = null,
        object? finish = null,
        object? visible = null) =>
        [
            id, null, stackIndex, type, null, start, finish,
            null, null, null, null, null, visible, null,
        ];

    private static Mock<Excel.Sheets> CreateSheetsMock(IReadOnlyList<SheetGraph> sheets)
    {
        var sheetsMock = new Mock<Excel.Sheets>();
        _ = sheetsMock.SetupGet(s => s.Count).Returns(sheets.Count);
        for (int index = 0; index < sheets.Count; index++)
        {
            var captured = sheets[index];
            _ = sheetsMock.Setup(s => s[index + 1]).Returns(captured.Worksheet.Object);
        }

        return sheetsMock;
    }

    private static TestableReader Build(
        Mock<Excel.Application> application,
        Mock<Excel.Workbook> workbook,
        IReadOnlyList<SheetGraph> sheets,
        Func<Excel.Range, object?> bodyValues)
    {
        var sheetsMock = CreateSheetsMock(sheets);
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheetsMock.Object);

        var table = sheets.SelectMany(s => s.Tables).ToArray();
        if (table.Length == 0)
        {
            return new TestableReader(application.Object, bodyValues: bodyValues);
        }

        var firstTable = table[0];
        var firstSheet = sheets[0];

        return new TestableReader(
            application.Object,
            (sheets, index) => firstSheet.Worksheet.Object,
            (listObjects, index) => firstTable.Table.Object,
            (columns, index) => firstTable.ColumnAt(index),
            (wb) => false,
            bodyValues);
    }

    [Fact]
    public void Read_refuses_when_the_application_is_null()
    {
        var reader = new TestableReader(null);

        GanttTableReadOutcome outcome = reader.Read();

        Assert.Equal(GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.NoActiveWorkbook), outcome);
        Assert.Empty(outcome.Rows);
    }

    [Fact]
    public void Read_refuses_when_the_active_workbook_is_null()
    {
        var application = new Mock<Excel.Application>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns((Excel.Workbook)null!);

        var reader = new TestableReader(application.Object);

        GanttTableReadOutcome outcome = reader.Read();

        Assert.Equal(GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.NoActiveWorkbook), outcome);
        Assert.Empty(outcome.Rows);
    }

    [Fact]
    public void Read_refuses_when_the_workbook_uses_the_1904_date_system()
    {
        var table = new TableGraph(GanttTableSchema.TableName, SchemaHeaders());
        var sheet = new SheetGraph(table);
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();

        var reader = new TestableReader(
            application.Object,
            (sheets, index) => sheet.Worksheet.Object,
            (listObjects, index) => table.Table.Object,
            (columns, index) => table.ColumnAt(index),
            (wb) => true,
            (body) => throw new InvalidOperationException("should not be called"));

        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(CreateSheetsMock([sheet]).Object);

        GanttTableReadOutcome outcome = reader.Read();

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttTableReadRefusalReason.DateSystemUnsupported, outcome.Refusal);
        Assert.Empty(outcome.Rows);
    }

    [Fact]
    public void Read_asks_the_date_converter_before_accessing_the_table()
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var converter = new Mock<IExcelDateSystemConverter>();
        _ = converter.Setup(c => c.IsSupported(ExcelDateSystemKind.Windows1900)).Returns(false);
        var sheetAtCalled = false;

        var reader = new TestableReader(
            application.Object,
            (_, _) =>
            {
                sheetAtCalled = true;
                throw new InvalidOperationException("table lookup must not run");
            },
            date1904: _ => false,
            bodyValues: _ => throw new InvalidOperationException("body must not be read"),
            dateSystemConverter: converter.Object);
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);

        GanttTableReadOutcome outcome = reader.Read();

        Assert.Equal(GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.DateSystemUnsupported), outcome);
        converter.Verify(c => c.IsSupported(ExcelDateSystemKind.Windows1900), Times.Once);
        Assert.False(sheetAtCalled);
    }

    [Fact]
    public void Read_refuses_when_the_table_is_missing()
    {
        var table = new TableGraph("SomeOtherTable", SchemaHeaders());
        var sheet = new SheetGraph(table);

        var reader = Build(new Mock<Excel.Application>(), new Mock<Excel.Workbook>(), [sheet], _ => throw new InvalidOperationException("should not be called"));

        GanttTableReadOutcome outcome = reader.Read();

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttTableReadRefusalReason.TableMissing, outcome.Refusal);
        Assert.Empty(outcome.Rows);
    }

    [Fact]
    public void Read_refuses_when_a_required_header_is_missing()
    {
        var headers = SchemaHeaders();
        headers.Remove("Type");
        var table = new TableGraph(GanttTableSchema.TableName, headers);
        var sheet = new SheetGraph(table);

        var reader = Build(new Mock<Excel.Application>(), new Mock<Excel.Workbook>(), [sheet], _ => throw new InvalidOperationException("should not be called"));

        GanttTableReadOutcome outcome = reader.Read();

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttTableReadRefusalReason.TableMissing, outcome.Refusal);
        Assert.Empty(outcome.Rows);
    }

    [Fact]
    public void Read_returns_an_empty_list_when_the_table_has_no_body()
    {
        var table = new TableGraph(GanttTableSchema.TableName, SchemaHeaders());
        table.WithNoBody();
        var sheet = new SheetGraph(table);

        var reader = Build(new Mock<Excel.Application>(), new Mock<Excel.Workbook>(), [sheet], _ => throw new InvalidOperationException("should not be called"));

        GanttTableReadOutcome outcome = reader.Read();

        Assert.True(outcome.Succeeded);
        Assert.Null(outcome.Refusal);
        Assert.Empty(outcome.Rows);
    }

    [Fact]
    public void Read_returns_rows_in_body_order_with_converted_values()
    {
        var table = new TableGraph(GanttTableSchema.TableName, SchemaHeaders());
        var sheet = new SheetGraph(table);
        var matrix = BodyMatrix(
            FullRow(id: "  G-1  ", stackIndex: 1.0, type: "As-Planned Activity", start: 44927.0, finish: 44931.0, visible: true),
            FullRow(id: "G-2", stackIndex: "2", type: "Delineator", start: 44932.0, visible: "TRUE"));

        var reader = Build(new Mock<Excel.Application>(), new Mock<Excel.Workbook>(), [sheet], _ => matrix);

        var outcome = reader.Read();

        Assert.True(outcome.Succeeded);
        Assert.Null(outcome.Refusal);
        Assert.Equal(2, outcome.Rows.Count);
        Assert.Equal(1, outcome.Rows[0].RowNumber);
        Assert.Equal("G-1", outcome.Rows[0].Id);
        Assert.Equal(1, outcome.Rows[0].StackIndex);
        Assert.Equal("As-Planned Activity", outcome.Rows[0].TypeText);
        Assert.Equal(new DateOnly(2023, 1, 1), outcome.Rows[0].Start);
        Assert.Equal(new DateOnly(2023, 1, 5), outcome.Rows[0].Finish);
        Assert.True(outcome.Rows[0].Visible);
        Assert.Equal(2, outcome.Rows[1].RowNumber);
        Assert.Equal("G-2", outcome.Rows[1].Id);
        Assert.Equal("Delineator", outcome.Rows[1].TypeText);
        Assert.Equal(2, outcome.Rows[1].StackIndex);
    }

    [Fact]
    public void Read_maps_columns_by_header_name_not_position()
    {
        var reordered = SchemaHeaders();
        (reordered[0], reordered[5]) = (reordered[5], reordered[0]);
        var table = new TableGraph(GanttTableSchema.TableName, reordered);
        var sheet = new SheetGraph(table);
        var row = FullRow();
        row[0] = 44927.0;
        row[5] = "G-reordered";
        var matrix = BodyMatrix(row);

        var reader = Build(new Mock<Excel.Application>(), new Mock<Excel.Workbook>(), [sheet], _ => matrix);

        var outcome = reader.Read();

        Assert.True(outcome.Succeeded);
        Assert.Equal("G-reordered", outcome.Rows[0].Id);
        Assert.Equal(new DateOnly(2023, 1, 1), outcome.Rows[0].Start);
    }

    [Fact]
    public void Read_preserves_rows_with_blanks_and_error_cells_as_null()
    {
        var table = new TableGraph(GanttTableSchema.TableName, SchemaHeaders());
        var sheet = new SheetGraph(table);
        var matrix = BodyMatrix(FullRow(id: "G-blank", start: -2146826281, visible: "   "));

        var reader = Build(new Mock<Excel.Application>(), new Mock<Excel.Workbook>(), [sheet], _ => matrix);

        var outcome = reader.Read();

        Assert.True(outcome.Succeeded);
        var row = Assert.Single(outcome.Rows);
        Assert.Equal("G-blank", row.Id);
        Assert.Null(row.Start);
        Assert.Null(row.StackIndex);
        Assert.Null(row.Visible);
    }
    [Fact]
    public void Read_maps_each_excel_error_cell_to_null_and_keeps_row_order()
    {
        // Mixed matrix: the first row writes Excel error ints into three
        // columns (Start #N/A, Finish #DIV/0!, StackIndex #VALUE! - codes from
        // the installed PIA layout pinned in ExcelCellConverterTests) while its
        // text fields stay readable; the second row is clean. Every error field
        // must read null and the row must survive in body order with its
        // RowNumber and readable fields intact.
        var table = new TableGraph(GanttTableSchema.TableName, SchemaHeaders());
        var sheet = new SheetGraph(table);
        var matrix = BodyMatrix(
            FullRow(id: "G-error", stackIndex: -2146826273, type: "As-Planned Activity", start: -2146826246, finish: -2146826281, visible: true),
            FullRow(id: "G-clean", stackIndex: 2.0, type: "Milestone", start: 44932.0, visible: "TRUE"));

        var reader = Build(new Mock<Excel.Application>(), new Mock<Excel.Workbook>(), [sheet], _ => matrix);

        var outcome = reader.Read();

        Assert.True(outcome.Succeeded);
        Assert.Equal(2, outcome.Rows.Count);

        GanttRowDto errorRow = outcome.Rows[0];
        Assert.Equal(1, errorRow.RowNumber);
        Assert.Equal("G-error", errorRow.Id);
        Assert.Equal("As-Planned Activity", errorRow.TypeText);
        Assert.Null(errorRow.Start);
        Assert.Null(errorRow.Finish);
        Assert.Null(errorRow.StackIndex);
        Assert.True(errorRow.Visible);

        GanttRowDto cleanRow = outcome.Rows[1];
        Assert.Equal(2, cleanRow.RowNumber);
        Assert.Equal("G-clean", cleanRow.Id);
        Assert.Equal(new DateOnly(2023, 1, 6), cleanRow.Start);
        Assert.Equal(2, cleanRow.StackIndex);
    }


}
