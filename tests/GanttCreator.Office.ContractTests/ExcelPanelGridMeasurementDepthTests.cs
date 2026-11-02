using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Depth tests for <see cref="ExcelPanelGridMeasurement"/>: the refusal arms of
/// <c>Measure</c>, the sheet/table/column searches the adapter performs for
/// itself, and each read seam's own body.
/// </summary>
/// <remarks>
/// A different cut from <see cref="ExcelPanelGridMeasurementTests"/>, which
/// substitutes <c>FindGanttTable</c>, <c>FindColumn</c> and the whole read layer
/// so its assertions are about the measurements that reach Core. This file
/// substitutes only the collection indexers, so the searches and the read
/// bodies run, and calls the read seams directly where a full
/// <c>Measure</c> would need the worksheet-row plumbing a contract test cannot
/// model.
/// </remarks>
public sealed class ExcelPanelGridMeasurementDepthTests
{
    /// <summary>
    /// The adapter with only the collection indexers substituted, so
    /// <c>FindGanttTable</c> and <c>FindColumn</c> run their own loops.
    /// </summary>
    private sealed class IndexedMeasurement(
        object? application,
        IReadOnlyList<Excel._Worksheet?> sheets,
        IReadOnlyList<Excel.ListObject?> tables,
        IReadOnlyList<Excel.ListColumn?> columns) : ExcelPanelGridMeasurement(application)
    {
        internal override int GetSheetCount(Excel.Sheets sheets1) => sheets.Count;

        internal override Excel._Worksheet? GetSheetAt(Excel.Sheets sheets1, int index) =>
            index >= 1 && index <= sheets.Count ? sheets[index - 1] : null;

        internal override int GetListObjectCount(Excel.ListObjects objects) => tables.Count;

        internal override Excel.ListObject? GetListObjectAt(Excel.ListObjects objects, int index) =>
            index >= 1 && index <= tables.Count ? tables[index - 1] : null;

        internal override int GetColumnCount(Excel.ListColumns columns1) => columns.Count;

        internal override Excel.ListColumn? GetColumnAt(Excel.ListColumns columns1, int index) =>
            index >= 1 && index <= columns.Count ? columns[index - 1] : null;
    }

    private static Excel.Sheets Sheets() => new Mock<Excel.Sheets>().Object;

    private static Excel.ListObjects TableCollection() => new Mock<Excel.ListObjects>().Object;

    private static Excel.ListColumns ColumnCollection() => new Mock<Excel.ListColumns>().Object;

    private static Excel.ListObject Table(string name, Excel.ListColumns? columnCollection = null)
    {
        Mock<Excel.ListObject> table = new();
        _ = table.SetupGet(t => t.Name).Returns(name);
        if (columnCollection is not null)
        {
            _ = table.SetupGet(t => t.ListColumns).Returns(columnCollection);
        }

        return table.Object;
    }

    private static Excel.ListColumn Column(string name, object? widthPt = null)
    {
        Mock<Excel.Range> range = new();
        _ = range.SetupGet(r => r.Width).Returns(widthPt!);
        Mock<Excel.ListColumn> column = new();
        _ = column.SetupGet(c => c.Name).Returns(name);
        _ = column.SetupGet(c => c.Range).Returns(range.Object);
        return column.Object;
    }

    private static Excel._Worksheet Sheet(Excel.ListObjects? tables)
    {
        Mock<Excel._Worksheet> sheet = new();
        if (tables is not null)
        {
            _ = sheet.SetupGet(s => s.ListObjects).Returns(tables);
        }

        return sheet.Object;
    }

    private static Excel.Application Application(Excel.Sheets? sheets)
    {
        Mock<Excel.Workbook> workbook = new();
        if (sheets is not null)
        {
            _ = workbook.SetupGet(w => w.Sheets).Returns(sheets);
        }

        Mock<Excel.Application> application = new();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        return application.Object;
    }

    private static Excel.Range Range(double top, double left, object? rowHeight = null, object? width = null)
    {
        Mock<Excel.Range> range = new();
        _ = range.SetupGet(r => r.Top).Returns(top);
        _ = range.SetupGet(r => r.Left).Returns(left);
        if (rowHeight is not null)
        {
            _ = range.SetupGet(r => r.RowHeight).Returns(rowHeight);
        }

        if (width is not null)
        {
            _ = range.SetupGet(r => r.Width).Returns(width);
        }

        return range.Object;
    }

    [Fact]
    public void Measure_refuses_when_the_host_supplied_no_application_object()
    {
        var measurement = new ExcelPanelGridMeasurement(null);

        PanelGridOutcome outcome = measurement.Measure(["Id"]);

        Assert.Equal(PanelGridRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }

    [Fact]
    public void Measure_refuses_an_empty_column_list_before_touching_the_host()
    {
        var measurement = new ExcelPanelGridMeasurement(Application(Sheets()));

        PanelGridOutcome outcome = measurement.Measure([]);

        Assert.Equal(PanelGridRefusalReason.InvalidMeasurement, outcome.Refusal);
    }

    [Fact]
    public void Measure_refuses_when_the_active_workbook_reports_no_sheet_collection()
    {
        var measurement = new ExcelPanelGridMeasurement(Application(sheets: null));

        PanelGridOutcome outcome = measurement.Measure(["Id"]);

        Assert.Equal(PanelGridRefusalReason.TableMissing, outcome.Refusal);
    }

    [Fact]
    public void Measure_refuses_when_no_sheet_carries_the_gantt_table()
    {
        IndexedMeasurement measurement = new(
            Application(Sheets()),
            [Sheet(TableCollection()), Sheet(tables: null)],
            [Table("tblSomethingElse")],
            [Column("Id", widthPt: 40.0)]);

        PanelGridOutcome outcome = measurement.Measure(["Id"]);

        Assert.Equal(PanelGridRefusalReason.TableMissing, outcome.Refusal);
    }

    [Fact]
    public void Measure_refuses_when_the_table_does_not_carry_a_requested_column()
    {
        IndexedMeasurement measurement = new(
            Application(Sheets()),
            [Sheet(TableCollection())],
            [Table(GanttTableSchema.TableName)],
            [Column("Other", widthPt: 40.0)]);

        PanelGridOutcome outcome = measurement.Measure(["Id"]);

        Assert.Equal(PanelGridRefusalReason.TableMissing, outcome.Refusal);
    }

    [Fact]
    public void The_table_search_skips_a_null_sheet_a_sheet_without_tables_and_a_foreign_table()
    {
        // Sheet 1 does not resolve, sheet 2 has no ListObjects, sheet 3 carries a
        // table the schema does not name, sheet 4 carries the Gantt table. The
        // search must walk past all three rather than stopping at the first miss.
        Excel.ListObject gantt = Table(GanttTableSchema.TableName);
        IndexedMeasurement measurement = new(
            Application(Sheets()),
            [null, Sheet(tables: null), Sheet(TableCollection()), Sheet(TableCollection())],
            [Table("tblSomethingElse"), gantt],
            [Column("Id", widthPt: 40.0)]);

        Excel.ListObject? found = measurement.FindGanttTable(Sheets());

        Assert.Same(gantt, found);
    }

    [Fact]
    public void The_table_search_returns_null_when_every_sheet_misses()
    {
        IndexedMeasurement measurement = new(
            Application(Sheets()),
            [Sheet(TableCollection())],
            [Table("tblSomethingElse")],
            []);

        Assert.Null(measurement.FindGanttTable(Sheets()));
    }

    [Fact]
    public void The_column_search_returns_null_when_the_table_reports_no_column_collection()
    {
        IndexedMeasurement measurement = new(Application(Sheets()), [], [], []);

        Assert.Null(measurement.FindColumn(Table("tblGanttData", columnCollection: null), "Id"));
    }

    [Fact]
    public void The_column_search_skips_a_null_column_and_matches_the_name_ordinally()
    {
        Excel.ListColumn wanted = Column("Id", widthPt: 40.0);
        IndexedMeasurement measurement = new(
            Application(Sheets()),
            [],
            [],
            [null, Column("id", widthPt: 10.0), wanted]);

        Excel.ListColumn? found = measurement.FindColumn(Table(GanttTableSchema.TableName, ColumnCollection()), "Id");

        // "id" is not "Id": the comparison is ordinal, because the schema name is
        // the reconciliation key.
        Assert.Same(wanted, found);
    }

    [Fact]
    public void The_origin_is_read_from_the_body_range_top_and_left()
    {
        Mock<Excel.ListObject> table = new();
        _ = table.SetupGet(t => t.DataBodyRange).Returns(Range(top: 120.5, left: 48.25));
        var measurement = new ExcelPanelGridMeasurement(null);

        Assert.Equal(120.5, measurement.ReadOriginTop(table.Object));
        Assert.Equal(48.25, measurement.ReadOriginLeft(table.Object));
    }

    [Fact]
    public void The_origin_is_absent_when_the_table_has_no_body_range()
    {
        Mock<Excel.ListObject> table = new();
        var measurement = new ExcelPanelGridMeasurement(null);

        Assert.Null(measurement.ReadOriginTop(table.Object));
        Assert.Null(measurement.ReadOriginLeft(table.Object));
    }

    [Fact]
    public void The_header_height_is_read_from_the_header_range()
    {
        Mock<Excel.ListObject> table = new();
        _ = table.SetupGet(t => t.HeaderRowRange).Returns(Range(top: 0, left: 0, rowHeight: 15.75));
        var measurement = new ExcelPanelGridMeasurement(null);

        Assert.Equal(15.75, measurement.ReadHeaderRowHeight(table.Object));
    }

    [Fact]
    public void The_header_height_is_absent_when_the_table_reports_no_header_range()
    {
        var measurement = new ExcelPanelGridMeasurement(null);

        Assert.Null(measurement.ReadHeaderRowHeight(new Mock<Excel.ListObject>().Object));
    }

    [Fact]
    public void The_column_width_is_read_through_the_column_range()
    {
        var measurement = new ExcelPanelGridMeasurement(null);

        Assert.Equal(42.5, measurement.ReadColumnWidth(Column("Id", widthPt: 42.5)));
    }

    [Fact]
    public void The_column_width_is_absent_when_the_column_reports_no_range()
    {
        Mock<Excel.ListColumn> column = new();
        var measurement = new ExcelPanelGridMeasurement(null);

        Assert.Null(measurement.ReadColumnWidth(column.Object));
    }

    [Fact]
    public void The_body_row_count_is_read_from_the_rows_collection()
    {
        Mock<Excel.Range> rows = new();
        _ = rows.SetupGet(r => r.Count).Returns(7);
        Mock<Excel.Range> body = new();
        _ = body.SetupGet(r => r.Rows).Returns(rows.Object);
        var measurement = new ExcelPanelGridMeasurement(null);

        Assert.Equal(7, measurement.GetBodyRowCount(body.Object));
    }

    [Fact]
    public void The_body_row_count_is_zero_when_the_host_reports_no_rows_collection()
    {
        var measurement = new ExcelPanelGridMeasurement(null);

        Assert.Equal(0, measurement.GetBodyRowCount(new Mock<Excel.Range>().Object));
    }

    [Fact]
    public void The_first_body_row_index_is_read_from_the_range_row()
    {
        Mock<Excel.Range> body = new();
        _ = body.SetupGet(r => r.Row).Returns(4);
        var measurement = new ExcelPanelGridMeasurement(null);

        Assert.Equal(4, measurement.GetRangeRow(body.Object));
    }
}
