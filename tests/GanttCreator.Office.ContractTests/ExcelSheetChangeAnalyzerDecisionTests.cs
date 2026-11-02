using System.Collections;
using System.Text;
using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for the parts of <see cref="ExcelSheetChangeAnalyzer"/> the
/// existing test class did not reach: the pure address parser, and the decision
/// path through a mocked application/workbook/table graph.
/// </summary>
/// <remarks>
/// Unlike the mutating adapters this one has no indexer seams, because it reads
/// only plain properties (<c>Row</c>, <c>Count</c>, <c>Name</c>) plus collection
/// enumerators, all of which Moq can stub - so the whole decision path runs
/// without a host and without changing production code.
/// </remarks>
public sealed class ExcelSheetChangeAnalyzerDecisionTests
{
    [Theory]
    [InlineData("A1", 1, 1, 1, 1)]
    [InlineData("A1:B2", 1, 2, 1, 2)]
    [InlineData("$A$1:$B$2", 1, 2, 1, 2)]
    [InlineData("AA10:AA12", 10, 12, 27, 27)]
    public void A_well_formed_address_reports_its_bounds(
        string address, int firstRow, int lastRow, int firstColumn, int lastColumn)
    {
        bool parsed = ExcelSheetChangeAnalyzer.TryParseRange(
            address, out int r1, out int r2, out int c1, out int c2);

        Assert.True(parsed);
        Assert.Equal(firstRow, r1);
        Assert.Equal(lastRow, r2);
        Assert.Equal(firstColumn, c1);
        Assert.Equal(lastColumn, c2);
    }

    [Fact]
    public void A_multi_area_address_reports_the_enclosing_bounds()
    {
        bool parsed = ExcelSheetChangeAnalyzer.TryParseRange(
            "A1:B2,D4:D6", out int firstRow, out int lastRow, out int firstColumn, out int lastColumn);

        Assert.True(parsed);
        Assert.Equal(1, firstRow);
        Assert.Equal(6, lastRow);
        Assert.Equal(1, firstColumn);
        Assert.Equal(4, lastColumn);
    }

    [Theory]
    [InlineData("")]
    [InlineData("  ")]
    [InlineData("nonsense")]
    [InlineData("A")]
    [InlineData("1")]
    [InlineData("A0")]
    [InlineData("A1:")]
    [InlineData("A1:B")]
    public void A_malformed_address_is_refused(string address)
    {
        bool parsed = ExcelSheetChangeAnalyzer.TryParseRange(
            address, out int firstRow, out int lastRow, out int firstColumn, out int lastColumn);

        Assert.False(parsed);
        Assert.Equal(0, firstRow);
        Assert.Equal(0, lastRow);
        Assert.Equal(0, firstColumn);
        Assert.Equal(0, lastColumn);
    }

    [Theory]
    [InlineData(3, 1, 3, true)]
    [InlineData(1, 1, 3, true)]
    [InlineData(4, 1, 3, false)]
    [InlineData(0, 1, 3, false)]
    public void A_column_is_included_only_inside_the_changed_span(
        int column, int firstColumn, int lastColumn, bool expected) =>
        Assert.Equal(expected, ExcelSheetChangeAnalyzer.IsColumnIncluded(column, firstColumn, lastColumn));

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void An_absent_address_never_refreshes(string? address) =>
        Assert.False(new ExcelSheetChangeAnalyzer(new Mock<Excel.Application>().Object).ShouldAutoRefresh(address));

    [Fact]
    public void No_application_never_refreshes()
    {
        ExcelSheetChangeAnalyzer analyzer = new(null);

        Assert.False(analyzer.ShouldAutoRefresh("'Book.xlsx|Sheet1'!A1"));
    }

    [Fact]
    public void No_active_workbook_never_refreshes()
    {
        Mock<Excel.Application> application = new();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns((Excel.Workbook)null!);

        Assert.False(new ExcelSheetChangeAnalyzer(application.Object).ShouldAutoRefresh("'Book.xlsx|Sheet1'!A1"));
    }

    [Theory]
    [InlineData("A1")]
    [InlineData("Sheet1!A1")]
    public void An_address_without_a_workbook_qualifier_never_refreshes(string address) =>
        Assert.False(Graph().Analyzer.ShouldAutoRefresh(address));

    [Fact]
    public void An_address_naming_another_workbook_never_refreshes() =>
        Assert.False(Graph().Analyzer.ShouldAutoRefresh("'Other.xlsx|Sheet1'!A6"));

    [Fact]
    public void An_address_naming_another_sheet_never_refreshes() =>
        Assert.False(Graph().Analyzer.ShouldAutoRefresh("'Book.xlsx|Other'!A6"));

    [Fact]
    public void A_change_inside_the_body_on_a_chart_affecting_column_refreshes()
    {
        AnalyzerGraph graph = Graph();

        Assert.True(graph.Analyzer.ShouldAutoRefresh("'Book.xlsx|Gantt'!A6"));
    }

    [Fact]
    public void A_change_outside_the_body_rows_never_refreshes()
    {
        AnalyzerGraph graph = Graph();

        // The body starts at row 5 and covers three rows, so row 9 is below it.
        Assert.False(graph.Analyzer.ShouldAutoRefresh("'Book.xlsx|Gantt'!A9"));
    }

    [Fact]
    public void A_change_on_the_derived_duration_column_never_refreshes()
    {
        AnalyzerGraph graph = Graph();
        int durationIndex = GanttTableSchema.Default.Columns
            .Select((column, index) => (column, index))
            .First(pair => pair.column.Name == "Duration").index;
        string column = ColumnLetters(durationIndex + 1);

        Assert.False(graph.Analyzer.ShouldAutoRefresh($"'Book.xlsx|Gantt'!{column}6"));
    }

    [Fact]
    public void A_workbook_without_the_data_table_never_refreshes()
    {
        AnalyzerGraph graph = Graph(tablePresent: false);

        Assert.False(graph.Analyzer.ShouldAutoRefresh("'Book.xlsx|Gantt'!A6"));
    }

    [Fact]
    public void A_table_missing_a_required_column_never_refreshes()
    {
        AnalyzerGraph graph = Graph(completeColumns: false);

        Assert.False(graph.Analyzer.ShouldAutoRefresh("'Book.xlsx|Gantt'!A6"));
    }

    [Fact]
    public void A_table_without_a_body_never_refreshes()
    {
        AnalyzerGraph graph = Graph(bodyPresent: false);

        Assert.False(graph.Analyzer.ShouldAutoRefresh("'Book.xlsx|Gantt'!A6"));
    }

    [Fact]
    public void A_sheet_collection_that_throws_never_refreshes()
    {
        AnalyzerGraph graph = Graph(sheetEnumerationFails: true);

        Assert.False(graph.Analyzer.ShouldAutoRefresh("'Book.xlsx|Gantt'!A6"));
    }

    [Fact]
    public void A_chart_sheet_before_the_gantt_sheet_is_skipped()
    {
        // The graph exposes a chart sheet first, which cannot hold a ListObject.
        AnalyzerGraph graph = Graph(chartSheetFirst: true);

        Assert.True(graph.Analyzer.ShouldAutoRefresh("'Book.xlsx|Gantt'!A6"));
    }

    private static string ColumnLetters(int column)
    {
        StringBuilder builder = new();
        int remaining = column;
        while (remaining > 0)
        {
            int digit = (remaining - 1) % 26;
            _ = builder.Insert(0, (char)('A' + digit));
            remaining = (remaining - 1) / 26;
        }

        return builder.ToString();
    }

    private static AnalyzerGraph Graph(
        bool tablePresent = true,
        bool completeColumns = true,
        bool bodyPresent = true,
        bool sheetEnumerationFails = false,
        bool chartSheetFirst = false)
    {
        const string WorkbookName = "Book.xlsx";
        const string SheetName = "Gantt";

        Mock<Excel.Application> application = new();
        Mock<Excel.Workbook> workbook = new();
        Mock<Excel.Sheets> sheets = new();
        Mock<Excel.Worksheet> worksheet = new();
        Mock<Excel.Worksheet> chartSheet = new();
        Mock<Excel.ListObjects> listObjects = new();
        Mock<Excel.ListObject> table = new();
        Mock<Excel.ListColumns> listColumns = new();
        Mock<Excel.Range> body = new();
        Mock<Excel.Range> bodyRows = new();

        int columnCount = GanttTableSchema.Default.Columns.Count;
        Mock<Excel.ListColumn>[] columns = new Mock<Excel.ListColumn>[columnCount];
        for (int index = 0; index < columnCount; index++)
        {
            int captured = index;
            columns[index] = new Mock<Excel.ListColumn>();
            _ = columns[index]
                .SetupGet(c => c.Name)
                .Returns(completeColumns ? GanttTableSchema.Default.Columns[captured].Name : "NotAColumn");
        }

        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Name).Returns(WorkbookName);
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);
        _ = worksheet.SetupGet(w => w.Name).Returns(SheetName);
        _ = worksheet.SetupGet(w => w.ListObjects).Returns(listObjects.Object);
        // A chart sheet reports an EMPTY, enumerable table collection. Leaving it
        // as a loose mock would throw from foreach and the adapter's outer catch
        // would abandon the whole sheet search instead of moving on - which no real
        // COM collection does.
        Mock<Excel.ListObjects> chartTables = new();
        _ = chartTables.Setup(l => l.GetEnumerator())
            .Returns(() => ((IEnumerable)new List<Excel.ListObject>()).GetEnumerator());
        _ = chartSheet.SetupGet(w => w.ListObjects).Returns(chartTables.Object);
        _ = listObjects.SetupGet(l => l.Count).Returns(1);
        _ = table.SetupGet(t => t.Name).Returns(GanttTableSchema.TableName);
        _ = table.SetupGet(t => t.ListColumns).Returns(listColumns.Object);
        _ = listColumns.SetupGet(c => c.Count).Returns(columnCount);
        _ = listColumns
            .Setup(c => c[It.IsAny<int>()])
            .Returns<int>(index => columns[index - 1].Object);
        _ = table.SetupGet(t => t.DataBodyRange).Returns(bodyPresent ? body.Object : (Excel.Range)null!);
        _ = body.SetupGet(b => b.Row).Returns(5);
        _ = body.SetupGet(b => b.Column).Returns(1);
        _ = body.SetupGet(b => b.Rows).Returns(bodyRows.Object);
        _ = bodyRows.SetupGet(r => r.Count).Returns(3);

        List<object> sheetEntries = chartSheetFirst
            ? [chartSheet.Object, worksheet.Object]
            : [worksheet.Object];
        if (sheetEnumerationFails)
        {
            _ = sheets.Setup(s => s.GetEnumerator())
                .Throws(new InvalidOperationException("COM proxy is dead."));
        }
        else
        {
            _ = sheets.Setup(s => s.GetEnumerator())
                .Returns(() => ((IEnumerable)sheetEntries).GetEnumerator());
        }

        if (!tablePresent)
        {
            _ = worksheet.SetupGet(w => w.ListObjects).Returns(new Mock<Excel.ListObjects>().Object);
        }
        else
        {
            List<Excel.ListObject> tables = [table.Object];
            _ = listObjects.Setup(l => l.GetEnumerator())
                .Returns(() => ((IEnumerable)tables).GetEnumerator());
        }

        return new AnalyzerGraph(new ExcelSheetChangeAnalyzer(application.Object));
    }

    /// <summary>The mocked graph and the analyzer over it.</summary>
    private sealed record AnalyzerGraph(ExcelSheetChangeAnalyzer Analyzer);
}
