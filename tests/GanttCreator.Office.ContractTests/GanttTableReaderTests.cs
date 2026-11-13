using System.Runtime.InteropServices;
using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

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
            IExcelDateSystemConverter? dateSystemConverter = null,
            Func<Excel.Range, int>? bodyFirstRow = null,
            Func<Excel.Range, int>? bodyRowCount = null,
            Func<Excel.Range, int, Excel.Range>? bodyRow = null,
            Func<Excel.Application, object?>? activeSheet = null
        )
            : base(application, dateSystemConverter)
        {
            SheetAt = sheetAt ?? ((_, _) => throw new InvalidOperationException("should not be called"));
            TableAt = tableAt ?? ((_, _) => throw new InvalidOperationException("should not be called"));
            ColumnAt = columnAt ?? ((_, _) => throw new InvalidOperationException("should not be called"));
            Date1904 = date1904 ?? (_ => throw new InvalidOperationException("should not be called"));
            BodyValue = bodyValues ?? (_ => throw new InvalidOperationException("should not be called"));
            BodyFirstRow = bodyFirstRow ?? (_ => throw new InvalidOperationException("should not be called"));
            BodyRowCount = bodyRowCount ?? (_ => throw new InvalidOperationException("should not be called"));
            BodyRow = bodyRow ?? ((_, _) => throw new InvalidOperationException("should not be called"));
            ActiveSheet = activeSheet;
        }

        public Func<Excel.Sheets, int, object> SheetAt { get; }
        public Func<Excel.ListObjects, int, Excel.ListObject> TableAt { get; }
        public Func<Excel.ListColumns, int, Excel.ListColumn> ColumnAt { get; }
        public Func<Excel.Workbook, bool> Date1904 { get; }
        public Func<Excel.Range, object?> BodyValue { get; }
        public Func<Excel.Range, int> BodyFirstRow { get; }
        public Func<Excel.Range, int> BodyRowCount { get; }
        public Func<Excel.Range, int, Excel.Range> BodyRow { get; }

        public Func<Excel.Application, object?>? ActiveSheet { get; }

        internal override bool GetDate1904(Excel.Workbook workbook) => Date1904(workbook);

        internal override int GetBodyFirstRow(Excel.Range body) => BodyFirstRow(body);

        internal override int GetBodyRowCount(Excel.Range body) => BodyRowCount(body);

        internal override Excel.Range GetBodyRow(Excel.Range body, int index) => BodyRow(body, index);

        internal override object GetSheetAt(Excel.Sheets sheets, int index) => SheetAt(sheets, index);

        internal override Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index) => TableAt(listObjects, index);

        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index) => ColumnAt(columns, index);

        internal override object? GetBodyValues(Excel.Range body) => BodyValue(body);

        internal override object? GetActiveSheet(Excel.Application application) =>
            ActiveSheet is null ? null : ActiveSheet(application);
    }

    private static List<string> SchemaHeaders() => GanttTableSchema.Default.Columns.Select(c => c.Name).ToList();

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

    /// <summary>
    /// Builds one full-width body row. Built from the schema's own column names
    /// rather than a positional literal, because a hard-coded array silently
    /// breaks when a column is added -- R4.7A's <c>SiblingOrder</c> made 20 tests
    /// fail with an IndexOutOfRange deep inside the matrix helper, which named no
    /// column. Naming the columns means a future addition fails here, legibly.
    /// </summary>
    private static object?[] FullRow(
        object? id = null,
        object? stackIndex = null,
        object? type = null,
        object? start = null,
        object? finish = null,
        object? visible = null)
    {
        Dictionary<string, object?> values = new(StringComparer.Ordinal)
        {
            ["Id"] = id,
            ["StackIndex"] = stackIndex,
            ["Type"] = type,
            ["Start"] = start,
            ["Finish"] = finish,
            ["Visible"] = visible,
        };

        return [.. GanttTableSchema.Default.Columns.Select(c => values.GetValueOrDefault(c.Name))];
    }

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
        Func<Excel.Range, object?> bodyValues
    )
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
            bodyValues
        );
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
    public void Read_reports_the_active_sheet_as_the_gantt_sheet_when_it_carries_the_table()
    {
        var table = new TableGraph(GanttTableSchema.TableName, SchemaHeaders());
        var graph = new SheetGraph(table);
        Excel.Worksheet ganttSheet = graph.Worksheet.Object;
        _ = graph.Worksheet.SetupGet(w => w.Name).Returns(GanttWorkbookContract.GanttSheetLabel);
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(CreateSheetsMock([graph]).Object);

        var reader = new TestableReader(
            application.Object,
            (_, _) => ganttSheet,
            (_, _) => table.Table.Object,
            (_, index) => table.ColumnAt(index),
            _ => false,
            _ => BodyMatrix(FullRow()),
            activeSheet: _ => ganttSheet);

        GanttTableReadOutcome outcome = reader.Read();

        Assert.True(outcome.Succeeded);
        Assert.True(outcome.ActiveSheetIsGanttSheet);
    }

    [Fact]
    public void Read_reports_the_active_sheet_as_not_the_gantt_sheet_when_another_sheet_is_active()
    {
        var table = new TableGraph(GanttTableSchema.TableName, SchemaHeaders());
        var graph = new SheetGraph(table);
        Excel.Worksheet ganttSheet = graph.Worksheet.Object;
        _ = graph.Worksheet.SetupGet(w => w.Name).Returns(GanttWorkbookContract.GanttSheetLabel);
        var otherSheet = new Mock<Excel.Worksheet>();
        _ = otherSheet.SetupGet(w => w.Name).Returns("Notes");
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(CreateSheetsMock([graph]).Object);

        var reader = new TestableReader(
            application.Object,
            (_, _) => ganttSheet,
            (_, _) => table.Table.Object,
            (_, index) => table.ColumnAt(index),
            _ => false,
            _ => BodyMatrix(FullRow()),
            activeSheet: _ => otherSheet.Object);

        GanttTableReadOutcome outcome = reader.Read();

        Assert.True(outcome.Succeeded);
        Assert.False(outcome.ActiveSheetIsGanttSheet);
    }

    [Fact]
    public void Read_reports_a_non_worksheet_active_sheet_as_not_the_gantt_sheet()
    {
        var table = new TableGraph(GanttTableSchema.TableName, SchemaHeaders());
        var graph = new SheetGraph(table);
        Excel.Worksheet ganttSheet = graph.Worksheet.Object;
        _ = graph.Worksheet.SetupGet(w => w.Name).Returns(GanttWorkbookContract.GanttSheetLabel);
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(CreateSheetsMock([graph]).Object);

        var reader = new TestableReader(
            application.Object,
            (_, _) => ganttSheet,
            (_, _) => table.Table.Object,
            (_, index) => table.ColumnAt(index),
            _ => false,
            _ => BodyMatrix(FullRow()),
            activeSheet: _ => new object());

        GanttTableReadOutcome outcome = reader.Read();

        Assert.True(outcome.Succeeded);
        Assert.False(outcome.ActiveSheetIsGanttSheet);
    }

    [Fact]
    public void Read_reports_a_chart_sheet_active_sheet_as_not_the_gantt_sheet()
    {
        // State matrix, chart sheet active (2026-11-13). A chart sheet is a real
        // non-worksheet sheet type, not merely "some object": the cast guard must
        // refuse it so the whole ribbon stays gated on a chart sheet.
        var table = new TableGraph(GanttTableSchema.TableName, SchemaHeaders());
        var graph = new SheetGraph(table);
        Excel.Worksheet ganttSheet = graph.Worksheet.Object;
        _ = graph.Worksheet.SetupGet(w => w.Name).Returns(GanttWorkbookContract.GanttSheetLabel);
        Mock<Excel.Chart> chartSheet = new();
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(CreateSheetsMock([graph]).Object);

        var reader = new TestableReader(
            application.Object,
            (_, _) => ganttSheet,
            (_, _) => table.Table.Object,
            (_, index) => table.ColumnAt(index),
            _ => false,
            _ => BodyMatrix(FullRow()),
            activeSheet: _ => chartSheet.Object);

        GanttTableReadOutcome outcome = reader.Read();

        Assert.True(outcome.Succeeded);
        Assert.False(outcome.ActiveSheetIsGanttSheet);
    }

    [Fact]
    public void Read_does_not_fall_back_to_the_gantt_label_when_a_chart_sheet_is_active_and_the_table_is_missing()
    {
        // State matrix, chart sheet active with a missing table (2026-11-13). The
        // label fallback exists so the repair commands keep their context when the
        // table is gone; a chart sheet must not borrow that context merely because
        // a Gantt-labelled worksheet still exists.
        var table = new TableGraph("SomeOtherTable", SchemaHeaders());
        var graph = new SheetGraph(table);
        Excel.Worksheet ganttSheet = graph.Worksheet.Object;
        _ = graph.Worksheet.SetupGet(w => w.Name).Returns(GanttWorkbookContract.GanttSheetLabel);
        Mock<Excel.Chart> chartSheet = new();
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(CreateSheetsMock([graph]).Object);

        var reader = new TestableReader(
            application.Object,
            (_, _) => ganttSheet,
            (_, _) => table.Table.Object,
            (_, index) => table.ColumnAt(index),
            _ => false,
            bodyValues: _ => throw new InvalidOperationException("body must not be read"),
            activeSheet: _ => chartSheet.Object);

        GanttTableReadOutcome outcome = reader.Read();

        Assert.Equal(GanttTableReadRefusalReason.TableMissing, outcome.Refusal);
        Assert.False(outcome.ActiveSheetIsGanttSheet);
    }

    [Fact]
    public void Read_never_consults_sheet_or_workbook_protection()
    {
        // State matrix, protected sheet (2026-11-13). Protection blocks edits, not
        // reads, so the read port must stay protection-independent -- Refresh and
        // Validate run on protected Gantt sheets by design.
        var table = new TableGraph(GanttTableSchema.TableName, SchemaHeaders());
        var graph = new SheetGraph(table);
        Excel.Worksheet ganttSheet = graph.Worksheet.Object;
        _ = graph.Worksheet.SetupGet(w => w.Name).Returns(GanttWorkbookContract.GanttSheetLabel);
        _ = graph.Worksheet.SetupGet(w => w.ProtectContents).Returns(true);
        _ = graph.Worksheet.SetupGet(w => w.ProtectDrawingObjects).Returns(true);
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        _ = workbook.SetupGet(w => w.ProtectStructure).Returns(true);
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(CreateSheetsMock([graph]).Object);

        var reader = new TestableReader(
            application.Object,
            (_, _) => ganttSheet,
            (_, _) => table.Table.Object,
            (_, index) => table.ColumnAt(index),
            _ => false,
            _ => BodyMatrix(FullRow()),
            activeSheet: _ => ganttSheet);

        GanttTableReadOutcome outcome = reader.Read();

        Assert.True(outcome.Succeeded);
        graph.Worksheet.VerifyGet(w => w.ProtectContents, Times.Never());
        graph.Worksheet.VerifyGet(w => w.ProtectDrawingObjects, Times.Never());
        workbook.VerifyGet(w => w.ProtectStructure, Times.Never());
    }

    [Fact]
    public void Read_propagates_an_unexpected_com_failure_from_the_active_workbook_probe()
    {
        // State matrix, closed workbook (2026-11-13). IGanttTableReader is explicit:
        // "unexpected COM failures propagate to the command boundary". This is the
        // deliberate asymmetry with ISelectionIdentifier, which reports None: a read
        // that cannot even name its workbook must fail loudly, not report an empty
        // schedule that the renderer reconciles to nothing.
        var application = new Mock<Excel.Application>();
        _ = application.SetupGet(a => a.ActiveWorkbook)
            .Throws(ComFailure(unchecked((int)0x80010108)));

        var reader = new TestableReader(application.Object);

        _ = Assert.Throws<COMException>(() => reader.Read());
        _ = Assert.Throws<COMException>(() => reader.Read(Array.Empty<int>()));
    }

    [Fact]
    public void Read_falls_back_to_the_gantt_label_when_the_table_is_missing()
    {
        var table = new TableGraph("SomeOtherTable", SchemaHeaders());
        var graph = new SheetGraph(table);
        Excel.Worksheet ganttSheet = graph.Worksheet.Object;
        _ = graph.Worksheet.SetupGet(w => w.Name).Returns(GanttWorkbookContract.GanttSheetLabel);
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(CreateSheetsMock([graph]).Object);

        var reader = new TestableReader(
            application.Object,
            (_, _) => ganttSheet,
            (_, _) => table.Table.Object,
            (_, index) => table.ColumnAt(index),
            _ => false,
            bodyValues: _ => throw new InvalidOperationException("body must not be read"),
            activeSheet: _ => ganttSheet);

        GanttTableReadOutcome outcome = reader.Read();

        Assert.Equal(GanttTableReadRefusalReason.TableMissing, outcome.Refusal);
        Assert.True(outcome.ActiveSheetIsGanttSheet);
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
            (body) => throw new InvalidOperationException("should not be called")
        );

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
            dateSystemConverter: converter.Object
        );
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

        var reader = Build(
            new Mock<Excel.Application>(),
            new Mock<Excel.Workbook>(),
            [sheet],
            _ => throw new InvalidOperationException("should not be called")
        );

        GanttTableReadOutcome outcome = reader.Read();

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttTableReadRefusalReason.TableMissing, outcome.Refusal);
        Assert.Empty(outcome.Rows);
        Assert.False(outcome.ConfigurationPresent);
    }

    [Fact]
    public void A_missing_table_reports_the_configuration_sheet_when_it_exists()
    {
        // R5.16 / ADR-0056: this fact is what separates NeedsRepair from a
        // first run, so the reader must report it on the refusal path.
        var table = new TableGraph("SomeOtherTable", SchemaHeaders());
        var ganttSheet = new SheetGraph(table);
        var configSheet = new SheetGraph();
        _ = configSheet.Worksheet.SetupGet(w => w.Name).Returns(GanttWorkbookContract.ConfigSheetName);

        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        Mock<Excel.Sheets> sheetsMock = CreateSheetsMock([ganttSheet, configSheet]);
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheetsMock.Object);

        var reader = new TestableReader(
            application.Object,
            sheetAt: (_, index) => index == 1 ? ganttSheet.Worksheet.Object : configSheet.Worksheet.Object,
            tableAt: (_, _) => table.Table.Object,
            columnAt: (columns, index) => table.ColumnAt(index),
            date1904: _ => false,
            bodyValues: _ => throw new InvalidOperationException("body must not be read"));

        GanttTableReadOutcome outcome = reader.Read();

        Assert.Equal(GanttTableReadRefusalReason.TableMissing, outcome.Refusal);
        Assert.True(outcome.ConfigurationPresent);
    }

    [Fact]
    public void A_missing_table_reports_no_configuration_sheet_when_it_is_absent()
    {
        var table = new TableGraph("SomeOtherTable", SchemaHeaders());
        var ganttSheet = new SheetGraph(table);

        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        Mock<Excel.Sheets> sheetsMock = CreateSheetsMock([ganttSheet]);
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheetsMock.Object);

        var reader = new TestableReader(
            application.Object,
            sheetAt: (_, _) => ganttSheet.Worksheet.Object,
            tableAt: (_, _) => table.Table.Object,
            columnAt: (columns, index) => table.ColumnAt(index),
            date1904: _ => false,
            bodyValues: _ => throw new InvalidOperationException("body must not be read"));

        GanttTableReadOutcome outcome = reader.Read();

        Assert.Equal(GanttTableReadRefusalReason.TableMissing, outcome.Refusal);
        Assert.False(outcome.ConfigurationPresent);
    }

    [Fact]
    public void Read_refuses_when_a_required_header_is_missing()
    {
        var headers = SchemaHeaders();
        headers.Remove("Type");
        var table = new TableGraph(GanttTableSchema.TableName, headers);
        var sheet = new SheetGraph(table);

        var reader = Build(
            new Mock<Excel.Application>(),
            new Mock<Excel.Workbook>(),
            [sheet],
            _ => throw new InvalidOperationException("should not be called")
        );

        GanttTableReadOutcome outcome = reader.Read();

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttTableReadRefusalReason.TableUnreadable, outcome.Refusal);
        Assert.True(outcome.TablePresent);
        Assert.Empty(outcome.Rows);
    }

    [Fact]
    public void Read_reads_a_table_saved_before_the_date_label_column()
    {
        // R5.14 defect regression: a workbook saved before R5.14 has no
        // DateLabelPosition column. The column is optional (the SortOrder
        // precedent), so the legacy table still reads and its date-label
        // cells come back blank (= Auto). The ribbon-state gate keys on
        // this read; a TableMissing refusal here greyed the whole Gantt
        // group on every pre-R5.14 workbook.
        var headers = SchemaHeaders();
        headers.Remove("DateLabelPosition");
        var table = new TableGraph(GanttTableSchema.TableName, headers);
        var sheet = new SheetGraph(table);
        var row = FullRow(id: "G-legacy", type: "As-Planned Activity").ToList();
        row.RemoveAt(GanttTableSchema.Default.Columns.Select(c => c.Name).ToList().IndexOf("DateLabelPosition"));

        // The legacy body is one column narrower than the running
        // schema, so the matrix is built at the row's own width
        // rather than through BodyMatrix (which sizes to the schema).
        var matrix = Array.CreateInstance(typeof(object), [1, row.Count], [1, 1]);
        for (var column = 1; column <= row.Count; column++)
        {
            matrix.SetValue(row[column - 1], 1, column);
        }

        var reader = Build(
            new Mock<Excel.Application>(),
            new Mock<Excel.Workbook>(),
            [sheet],
            _ => matrix);

        GanttTableReadOutcome outcome = reader.Read();

        Assert.True(outcome.Succeeded);
        Assert.Null(outcome.Refusal);
        GanttRowDto read = Assert.Single(outcome.Rows);
        Assert.Equal("G-legacy", read.Id);
        Assert.Equal("As-Planned Activity", read.TypeText);
        Assert.Equal(GanttCellState.Empty, read.DateLabelPositionCell.State);
        Assert.Null(read.DateLabelPositionText);
    }

    [Fact]
    public void ExcelValue2Matrix_normalizes_one_based_and_zero_based_matrices()
    {
        var oneBased = Array.CreateInstance(typeof(object), [1, 2], [1, 1]);
        var zeroBased = Array.CreateInstance(typeof(object), [1, 2]);
        oneBased.SetValue("one", 1, 1);
        oneBased.SetValue(2.0, 1, 2);
        zeroBased.SetValue("one", 0, 0);
        zeroBased.SetValue(2.0, 0, 1);

        List<object?[]> oneBasedRows = ExcelValue2Matrix.ReadRows(oneBased);
        List<object?[]> zeroBasedRows = ExcelValue2Matrix.ReadRows(zeroBased);

        Assert.Equal(zeroBasedRows, oneBasedRows);
        Assert.Equal(["one", 2.0], oneBasedRows[0]);
    }

    [Fact]
    public void ExcelValue2Matrix_normalizes_nonstandard_positive_lower_bounds()
    {
        var matrix = Array.CreateInstance(typeof(object), [2, 2], [2, 3]);
        matrix.SetValue("a", 2, 3);
        matrix.SetValue("b", 2, 4);
        matrix.SetValue("c", 3, 3);
        matrix.SetValue("d", 3, 4);

        List<object?[]> rows = ExcelValue2Matrix.ReadRows(matrix);

        Assert.Equal(
            [
                ["a", "b"],
                ["c", "d"],
            ],
            rows
        );
    }

    [Fact]
    public void ExcelValue2Matrix_returns_no_rows_for_empty_one_based_matrix()
    {
        var matrix = Array.CreateInstance(typeof(object), [0, 0], [1, 1]);

        Assert.Empty(ExcelValue2Matrix.ReadRows(matrix));
    }

    [Fact]
    public void ExcelValue2Matrix_returns_empty_for_non_two_dimensional_shapes()
    {
        Assert.Empty(ExcelValue2Matrix.ReadRows(null));
        Assert.Empty(ExcelValue2Matrix.ReadRows("not a matrix"));
        Assert.Empty(ExcelValue2Matrix.ReadRows(Array.CreateInstance(typeof(object), [1])));
    }

    [Fact]
    public void Read_accepts_zero_based_body_matrices()
    {
        var table = new TableGraph(GanttTableSchema.TableName, SchemaHeaders());
        var sheet = new SheetGraph(table);
        var matrix = Array.CreateInstance(typeof(object), [1, GanttTableSchema.Default.Columns.Count]);
        for (int column = 0; column < GanttTableSchema.Default.Columns.Count; column++)
        {
            matrix.SetValue(null, 0, column);
        }

        matrix.SetValue("G-zero", 0, 0);
        matrix.SetValue("As-Planned Activity", 0, GanttTableSchema.Default.Columns.ToList().FindIndex(c => c.Name == "Type"));
        matrix.SetValue(0.0, 0, GanttTableSchema.Default.Columns.ToList().FindIndex(c => c.Name == "StackIndex"));
        matrix.SetValue(44927.0, 0, GanttTableSchema.Default.Columns.ToList().FindIndex(c => c.Name == "Start"));
        matrix.SetValue(44927.0, 0, GanttTableSchema.Default.Columns.ToList().FindIndex(c => c.Name == "Finish"));

        var reader = Build(new Mock<Excel.Application>(), new Mock<Excel.Workbook>(), [sheet], _ => matrix);
        GanttTableReadOutcome outcome = reader.Read();

        Assert.True(outcome.Succeeded);
        GanttRowDto row = Assert.Single(outcome.Rows);
        Assert.Equal("G-zero", row.Id);
        Assert.Equal("As-Planned Activity", row.TypeText);
        Assert.Equal(new DateOnly(2023, 1, 1), row.Start);
    }

    [Fact]
    public void Read_carries_the_duration_text_into_its_own_cell()
    {
        // R4.7C D1: Duration is read as text, because the contract carries a
        // non-duration marker for milestones and a blank for structural rows. A
        // numeric cell state could represent only the day count.
        var table = new TableGraph(GanttTableSchema.TableName, SchemaHeaders());
        var sheet = new SheetGraph(table);
        object?[] row = FullRow(id: "D-one", type: "As-Planned Activity");
        var columns = GanttTableSchema.Default.Columns.ToList();
        row[columns.FindIndex(c => c.Name == "Duration")] = "12";

        var reader = Build(
            new Mock<Excel.Application>(),
            new Mock<Excel.Workbook>(),
            [sheet],
            _ => BodyMatrix(row));

        GanttTableReadOutcome outcome = reader.Read();

        Assert.True(outcome.Succeeded);
        GanttRowDto read = Assert.Single(outcome.Rows);
        Assert.Equal(GanttCellState.Value, read.DurationCell.State);
        Assert.Equal("12", read.Duration);
    }

    [Fact]
    public void Read_leaves_duration_empty_when_the_cell_is_blank()
    {
        // A row the engine has not written yet must read as empty, not as an
        // unsupported cell: R4.7F computes the value on Refresh, and a blank is the
        // honest pre-Refresh state.
        var table = new TableGraph(GanttTableSchema.TableName, SchemaHeaders());
        var sheet = new SheetGraph(table);

        var reader = Build(
            new Mock<Excel.Application>(),
            new Mock<Excel.Workbook>(),
            [sheet],
            _ => BodyMatrix(FullRow(id: "D-two", type: "As-Planned Activity")));

        GanttTableReadOutcome outcome = reader.Read();

        Assert.True(outcome.Succeeded);
        GanttRowDto read = Assert.Single(outcome.Rows);
        Assert.Equal(GanttCellState.Empty, read.DurationCell.State);
        Assert.Null(read.Duration);
    }

    [Fact]
    public void Read_returns_an_empty_list_when_the_table_has_no_body()
    {
        var table = new TableGraph(GanttTableSchema.TableName, SchemaHeaders());
        table.WithNoBody();
        var sheet = new SheetGraph(table);

        var reader = Build(
            new Mock<Excel.Application>(),
            new Mock<Excel.Workbook>(),
            [sheet],
            _ => throw new InvalidOperationException("should not be called")
        );

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
            FullRow(id: "G-2", stackIndex: "2", type: "Delineator", start: 44932.0, visible: "TRUE")
        );

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
        Assert.Equal(GanttCellState.ExcelError, row.StartCell.State);
        Assert.Equal(GanttExcelErrorCode.DivisionByZero, row.StartCell.ErrorCode);
        Assert.Equal(GanttCellState.Empty, row.VisibleCell.State);
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
            FullRow(
                id: "G-error",
                stackIndex: -2146826273,
                type: "As-Planned Activity",
                start: -2146826246,
                finish: -2146826281,
                visible: true
            ),
            FullRow(id: "G-clean", stackIndex: 2.0, type: "Milestone", start: 44932.0, visible: "TRUE")
        );

        var reader = Build(new Mock<Excel.Application>(), new Mock<Excel.Workbook>(), [sheet], _ => matrix);

        var outcome = reader.Read();

        Assert.True(outcome.Succeeded);
        Assert.Equal(2, outcome.Rows.Count);

        GanttRowDto errorRow = outcome.Rows[0];
        Assert.Equal(1, errorRow.RowNumber);
        Assert.Equal("G-error", errorRow.Id);
        Assert.Equal("As-Planned Activity", errorRow.TypeText);
        Assert.Equal(GanttCellState.ExcelError, errorRow.StartCell.State);
        Assert.Equal(GanttExcelErrorCode.NotAvailable, errorRow.StartCell.ErrorCode);
        Assert.Equal(GanttCellState.ExcelError, errorRow.FinishCell.State);
        Assert.Equal(GanttExcelErrorCode.DivisionByZero, errorRow.FinishCell.ErrorCode);
        Assert.Equal(GanttCellState.ExcelError, errorRow.StackIndexCell.State);
        Assert.Equal(GanttExcelErrorCode.Value, errorRow.StackIndexCell.ErrorCode);
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

    // ---- Read(IReadOnlyList<int>): the delta read the ChangeTracker drives ----

    /// <summary>
    /// A reader wired for the worksheet-row overload: a real table and column map,
    /// with the body span reported by the fixture rather than by COM.
    /// </summary>
    private static TestableReader DeltaReader(
        Mock<Excel.Application> application,
        Mock<Excel.Workbook> workbook,
        TableGraph table,
        Func<Excel.Range, object?> bodyValues,
        int firstBodyRow = 4,
        int bodyRowCount = 3)
    {
        SheetGraph sheet = new(table);
        Mock<Excel.Sheets> sheets = CreateSheetsMock([sheet]);
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);
        Excel.Range rowRange = new Mock<Excel.Range>().Object;
        return new TestableReader(
            application.Object,
            sheetAt: (_, _) => sheet.Worksheet.Object,
            tableAt: (_, _) => table.Table.Object,
            columnAt: (_, index) => table.ColumnAt(index),
            date1904: _ => false,
            bodyValues: bodyValues,
            bodyFirstRow: _ => firstBodyRow,
            bodyRowCount: _ => bodyRowCount,
            bodyRow: (_, _) => rowRange);
    }

    [Fact]
    public void Read_by_row_numbers_refuses_without_an_application_object()
    {
        GanttTableReadOutcome outcome = new TestableReader(null).Read([5]);

        Assert.Equal(GanttTableReadRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }

    [Fact]
    public void Read_by_row_numbers_refuses_without_an_active_workbook()
    {
        Mock<Excel.Application> application = new();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns((Excel.Workbook)null!);

        GanttTableReadOutcome outcome = new TestableReader(application.Object).Read([5]);

        Assert.Equal(GanttTableReadRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }

    [Fact]
    public void Read_by_row_numbers_refuses_an_unsupported_date_system()
    {
        TableGraph table = new(GanttTableSchema.TableName, SchemaHeaders());
        SheetGraph sheet = new(table);
        Mock<Excel.Application> application = new();
        Mock<Excel.Workbook> workbook = new();
        Mock<Excel.Sheets> sheets = CreateSheetsMock([sheet]);
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);

        TestableReader reader = new(
            application.Object,
            sheetAt: (_, _) => sheet.Worksheet.Object,
            tableAt: (_, _) => table.Table.Object,
            columnAt: (_, index) => table.ColumnAt(index),
            date1904: _ => true);

        Assert.Equal(GanttTableReadRefusalReason.DateSystemUnsupported, reader.Read([5]).Refusal);
    }

    [Fact]
    public void Read_by_row_numbers_refuses_when_the_table_is_missing()
    {
        Mock<Excel.Application> application = new();
        Mock<Excel.Workbook> workbook = new();
        SheetGraph sheet = new();
        Mock<Excel.Sheets> sheets = CreateSheetsMock([sheet]);
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);

        TestableReader reader = new(
            application.Object,
            sheetAt: (_, _) => sheet.Worksheet.Object,
            date1904: _ => false);

        Assert.Equal(GanttTableReadRefusalReason.TableMissing, reader.Read([5]).Refusal);
    }

    [Fact]
    public void Read_by_row_numbers_refuses_when_a_required_header_is_missing()
    {
        TableGraph table = new(GanttTableSchema.TableName, ["Id"]);
        Mock<Excel.Application> application = new();
        Mock<Excel.Workbook> workbook = new();

        TestableReader reader = DeltaReader(
            application,
            workbook,
            table,
            _ => throw new InvalidOperationException("the body must not be read"));

        GanttTableReadOutcome outcome = reader.Read([5]);

        Assert.Equal(GanttTableReadRefusalReason.TableUnreadable, outcome.Refusal);
    }

    [Fact]
    public void Read_by_row_numbers_returns_empty_when_the_table_has_no_body()
    {
        TableGraph table = new(GanttTableSchema.TableName, SchemaHeaders());
        table.WithNoBody();
        Mock<Excel.Application> application = new();
        Mock<Excel.Workbook> workbook = new();

        TestableReader reader = DeltaReader(
            application,
            workbook,
            table,
            _ => throw new InvalidOperationException("the body must not be read"));

        GanttTableReadOutcome outcome = reader.Read([5]);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Empty(outcome.Rows);
    }

    [Fact]
    public void Read_by_row_numbers_returns_empty_for_an_empty_selection()
    {
        TableGraph table = new(GanttTableSchema.TableName, SchemaHeaders());
        Mock<Excel.Application> application = new();
        Mock<Excel.Workbook> workbook = new();

        TestableReader reader = DeltaReader(
            application,
            workbook,
            table,
            _ => throw new InvalidOperationException("the body must not be read"));

        Assert.Empty(reader.Read([]).Rows);
    }

    [Fact]
    public void Read_by_row_numbers_reads_only_the_requested_worksheet_rows()
    {
        // The body starts at worksheet row 4, so rows 5 and 6 are body indexes 2 and
        // 3. The read must translate, not treat the numbers as body offsets.
        TableGraph table = new(GanttTableSchema.TableName, SchemaHeaders());
        Mock<Excel.Application> application = new();
        Mock<Excel.Workbook> workbook = new();

        TestableReader reader = DeltaReader(
            application,
            workbook,
            table,
            _ => BodyMatrix(FullRow(id: "G-aaaa")),
            firstBodyRow: 4,
            bodyRowCount: 3);

        GanttTableReadOutcome outcome = reader.Read([5, 6]);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal([2, 3], outcome.Rows.Select(row => row.RowNumber));
    }

    [Fact]
    public void Read_by_row_numbers_skips_numbers_outside_the_body_span()
    {
        TableGraph table = new(GanttTableSchema.TableName, SchemaHeaders());
        Mock<Excel.Application> application = new();
        Mock<Excel.Workbook> workbook = new();

        TestableReader reader = DeltaReader(
            application,
            workbook,
            table,
            _ => BodyMatrix(FullRow(id: "G-aaaa")),
            firstBodyRow: 4,
            bodyRowCount: 3);

        // 1 is above the body and 99 is past its end; neither may widen the read.
        Assert.Empty(reader.Read([1, 99]).Rows);
    }

    [Fact]
    public void Read_by_row_numbers_skips_a_row_the_host_reports_no_values_for()
    {
        TableGraph table = new(GanttTableSchema.TableName, SchemaHeaders());
        Mock<Excel.Application> application = new();
        Mock<Excel.Workbook> workbook = new();

        TestableReader reader = DeltaReader(
            application,
            workbook,
            table,
            _ => null,
            firstBodyRow: 4,
            bodyRowCount: 3);

        Assert.Empty(reader.Read([5]).Rows);
    }

    /// <summary>
    /// Produces the exception a disconnected COM object raises, without
    /// constructing the runtime-reserved <see cref="COMException"/> directly
    /// (CA2201). The same production is used by the other Office contract tests.
    /// </summary>
    /// <param name="hresult">The HRESULT the host would report.</param>
    /// <returns>A live <see cref="COMException"/> carrying that HRESULT.</returns>
    private static COMException ComFailure(int hresult)
    {
        var exception = (COMException?)Marshal.GetExceptionForHR(hresult);
        return exception ?? throw new InvalidOperationException(
            $"The runtime did not produce a COMException for HRESULT 0x{hresult:X8}.");
    }
}
