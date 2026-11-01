using GanttCreator.Core;
using Xunit;

namespace GanttCreator.Office.ContractTests;

public sealed class ExcelSheetChangeAnalyzerTests
{
    [Theory]
    [InlineData("$C$5", 5, 5, 3, 3)]
    [InlineData("D7:F12", 7, 12, 4, 6)]
    [InlineData("$AA$9:$AB$8", 8, 9, 27, 28)]
    [InlineData("A1:B2,C4", 1, 4, 1, 3)]
    public void TryParseRange_ParsesCellAndRectangle(
        string address, int expectedFirstRow, int expectedLastRow, int expectedFirstColumn, int expectedLastColumn)
    {
        bool parsed = ExcelSheetChangeAnalyzer.TryParseRange(
            address, out int firstRow, out int lastRow, out int firstColumn, out int lastColumn);

        Assert.True(parsed);
        Assert.Equal(expectedFirstRow, firstRow);
        Assert.Equal(expectedLastRow, lastRow);
        Assert.Equal(expectedFirstColumn, firstColumn);
        Assert.Equal(expectedLastColumn, lastColumn);
    }

    [Theory]
    [InlineData("A1:B2,C4?", false)]
    [InlineData("A0", false)]
    [InlineData("A", false)]
    public void TryParseRange_RejectsMalformedOrMultiAreaAddress(string address, bool expected)
    {
        bool parsed = ExcelSheetChangeAnalyzer.TryParseRange(
            address, out _, out _, out _, out _);

        Assert.Equal(expected, parsed);
    }

    [Fact]
    public void IsColumnIncluded_UsesInclusiveRange()
    {
        Assert.True(ExcelSheetChangeAnalyzer.IsColumnIncluded(4, 4, 6));
        Assert.True(ExcelSheetChangeAnalyzer.IsColumnIncluded(6, 4, 6));
        Assert.False(ExcelSheetChangeAnalyzer.IsColumnIncluded(7, 4, 6));
    }

    [Fact]
    public void Every_schema_column_triggers_except_derived_duration()
    {
        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        int[] columnMap = [.. Enumerable.Range(1, schema.Count)];
        int bodyFirstColumn = 3;

        for (int index = 0; index < schema.Count; index++)
        {
            int absoluteColumn = bodyFirstColumn + columnMap[index] - 1;
            bool shouldTrigger = ExcelSheetChangeAnalyzer.HasChartAffectingColumn(
                schema, columnMap, bodyFirstColumn, absoluteColumn, absoluteColumn);

            Assert.Equal(!schema[index].Name.Equals("Duration", StringComparison.OrdinalIgnoreCase), shouldTrigger);
        }
    }

    [Theory]
    [InlineData(1, 4, 2, 8, 5, 10, 1, 8, false)]
    [InlineData(5, 5, 9, 9, 5, 10, 4, 8, false)]
    [InlineData(5, 5, 4, 4, 5, 10, 4, 8, true)]
    public void RangeIntersectsBody_RejectsHeadersAndOutsideRanges(
        int firstRow, int lastRow, int firstColumn, int lastColumn,
        int bodyFirstRow, int bodyLastRow, int bodyFirstColumn, int bodyLastColumn, bool expected)
    {
        Assert.Equal(expected, ExcelSheetChangeAnalyzer.RangeIntersectsBody(
            firstRow, lastRow, firstColumn, lastColumn,
            bodyFirstRow, bodyLastRow, bodyFirstColumn, bodyLastColumn));
    }

    [Fact]
    public void HasChartAffectingColumn_RejectsMapWithWrongSchemaLength()
    {
        Assert.False(ExcelSheetChangeAnalyzer.HasChartAffectingColumn(
            GanttTableSchema.Default.Columns, [], 1, 1, 1));
    }

    [Fact]
    public void HasChartAffectingColumn_RejectsNullSchema()
    {
        Assert.Throws<ArgumentNullException>(() => ExcelSheetChangeAnalyzer.HasChartAffectingColumn(
            null!, [], 1, 1, 1));
    }

    [Fact]
    public void HasChartAffectingColumn_RejectsNullColumnMap()
    {
        Assert.Throws<ArgumentNullException>(() => ExcelSheetChangeAnalyzer.HasChartAffectingColumn(
            GanttTableSchema.Default.Columns, null!, 1, 1, 1));
    }

    [Fact]
    public void Recalculation_has_no_automatic_refresh_path()
    {
        // R5.11 trigger matrix, recalc cell: recalculation is a NON-trigger by
        // design. Duration is engine-written, so no calculate event is needed;
        // pinning the absence makes adding one a deliberate, reviewed change
        // rather than a silent new trigger.
        Assert.DoesNotContain(
            typeof(ISheetChangeAnalyzer).GetMembers(),
            member => member.Name.Contains("Calculate", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(
            typeof(IExcelApplicationAdapter).GetMembers(),
            member => member.Name.Contains("Calculate", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void A_change_in_the_row_below_the_body_does_not_trigger_until_the_table_claims_it()
    {
        // R5.11 trigger matrix, table-resize cell. Typing below the table makes
        // Excel auto-expand it; the classifier reads the CURRENT body bounds, so
        // before the table claims the row the change is outside the body and the
        // change is not a trigger. Once the body includes the row (post-expand),
        // the same address is inside the body and does trigger.
        const int bodyFirstRow = 5;
        const int bodyLastRowBeforeExpand = 10;
        const int changedRow = 11;
        const int column = 4;

        Assert.False(ExcelSheetChangeAnalyzer.RangeIntersectsBody(
            changedRow, changedRow, column, column,
            bodyFirstRow, bodyLastRowBeforeExpand, column, column));

        Assert.True(ExcelSheetChangeAnalyzer.RangeIntersectsBody(
            changedRow, changedRow, column, column,
            bodyFirstRow, bodyLastRowBeforeExpand + 1, column, column));
    }

    [Fact]
    public void A_change_on_the_header_row_does_not_trigger()
    {
        // R5.11 trigger matrix, header cell: the body starts one row below the
        // header, so a header edit is outside the body even in the same column.
        Assert.False(ExcelSheetChangeAnalyzer.RangeIntersectsBody(
            firstRow: 4, lastRow: 4, firstColumn: 4, lastColumn: 4,
            bodyFirstRow: 5, bodyLastRow: 10, bodyFirstColumn: 4, bodyLastColumn: 8));
    }

    [Fact]
    public void A_multi_area_change_spanning_the_body_and_a_row_below_triggers()
    {
        // A paste that covers the last body row and the row below arrives as a
        // multi-area address; the analyzer classifies it by its bounding box, so
        // the chart-affecting body column still triggers.
        Assert.True(ExcelSheetChangeAnalyzer.RangeIntersectsBody(
            firstRow: 10, lastRow: 11, firstColumn: 4, lastColumn: 4,
            bodyFirstRow: 5, bodyLastRow: 10, bodyFirstColumn: 4, bodyLastColumn: 8));
    }

    [Fact]
    public void ShouldAutoRefresh_refuses_a_null_or_empty_address()
    {
        ExcelSheetChangeAnalyzer analyzer = new(application: null);

        Assert.False(analyzer.ShouldAutoRefresh(null));
        Assert.False(analyzer.ShouldAutoRefresh(string.Empty));
    }

    [Fact]
    public void ShouldAutoRefresh_refuses_when_no_workbook_is_available()
    {
        // A different sheet or an unqualified address can never reach the table
        // classifier without an active workbook; the guard refuses before any
        // COM access, which is what keeps the scheduler inert outside Excel.
        ExcelSheetChangeAnalyzer analyzer = new(application: null);

        Assert.False(analyzer.ShouldAutoRefresh("Sheet1!$C$5"));
        Assert.False(analyzer.ShouldAutoRefresh("$C$5"));
    }
}
