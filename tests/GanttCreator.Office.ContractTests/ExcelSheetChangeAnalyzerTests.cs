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
}
