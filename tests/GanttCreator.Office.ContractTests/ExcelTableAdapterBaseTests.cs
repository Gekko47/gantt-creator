using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for the shared table traversal and its D1 release policy.
/// </summary>
public sealed class ExcelTableAdapterBaseTests
{
    /// <summary>
    /// Supplies the traversal's seams from lists so no COM host is needed, and
    /// records every release the base asks for.
    /// </summary>
    private sealed class RecordingAdapter : ExcelTableAdapterBase
    {
        internal List<object> Sheets { get; } = [];

        internal Dictionary<object, List<Excel.ListObject>> Tables { get; } = [];

        internal List<object?> Released { get; } = [];

        internal override int GetSheetCount(Excel.Sheets sheets) => Sheets.Count;

        internal override object? GetSheetAt(Excel.Sheets sheets, int index) =>
            index >= 1 && index <= Sheets.Count ? Sheets[index - 1] : null;

        internal override int GetListObjectCount(Excel.ListObjects objects) =>
            Tables.TryGetValue(objects, out List<Excel.ListObject>? tables) ? tables.Count : 0;

        internal override Excel.ListObject? GetTableAt(Excel.ListObjects objects, int index) =>
            Tables.TryGetValue(objects, out List<Excel.ListObject>? tables)
            && index >= 1
            && index <= tables.Count
                ? tables[index - 1]
                : null;

        internal override void ReleaseAcquired(object? proxy) => Released.Add(proxy);
    }

    [Fact]
    public void Returned_table_and_sheet_are_not_released_and_skipped_proxies_are()
    {
        var skippedTable = new Mock<Excel.ListObject>();
        _ = skippedTable.SetupGet(table => table.Name).Returns("Notes");
        var matchingTable = new Mock<Excel.ListObject>();
        _ = matchingTable.SetupGet(table => table.Name).Returns("tblGanttData");
        var listObjects = new Mock<Excel.ListObjects>();
        var ganttSheet = new Mock<Excel.Worksheet>();
        _ = ganttSheet.SetupGet(sheet => sheet.ListObjects).Returns(listObjects.Object);

        var adapter = new RecordingAdapter();
        adapter.Sheets.Add(new object());
        adapter.Sheets.Add(ganttSheet.Object);
        adapter.Tables[listObjects.Object] = [skippedTable.Object, matchingTable.Object];

        bool found = adapter.TryLocateTable(
            new Mock<Excel.Sheets>().Object,
            out Excel.Worksheet? worksheet,
            out Excel.ListObject? table);

        Assert.True(found);
        Assert.Same(ganttSheet.Object, worksheet);
        Assert.Same(matchingTable.Object, table);
        Assert.Equal(2, adapter.Released.Count);
        Assert.Contains(skippedTable.Object, adapter.Released);
        Assert.Contains(listObjects.Object, adapter.Released);
        Assert.DoesNotContain(matchingTable.Object, adapter.Released);
        Assert.DoesNotContain(ganttSheet.Object, adapter.Released);
    }

    [Fact]
    public void Missing_table_releases_the_skipped_proxies_and_the_non_owning_sheet()
    {
        var skippedTable = new Mock<Excel.ListObject>();
        _ = skippedTable.SetupGet(table => table.Name).Returns("Notes");
        var listObjects = new Mock<Excel.ListObjects>();
        var sheet = new Mock<Excel.Worksheet>();
        _ = sheet.SetupGet(candidate => candidate.ListObjects).Returns(listObjects.Object);

        var adapter = new RecordingAdapter();
        adapter.Sheets.Add(sheet.Object);
        adapter.Tables[listObjects.Object] = [skippedTable.Object];

        bool found = adapter.TryLocateTable(
            new Mock<Excel.Sheets>().Object,
            out Excel.Worksheet? worksheet,
            out Excel.ListObject? table);

        Assert.False(found);
        Assert.Null(worksheet);
        Assert.Null(table);
        Assert.Equal(3, adapter.Released.Count);
        Assert.Contains(skippedTable.Object, adapter.Released);
        Assert.Contains(listObjects.Object, adapter.Released);
        Assert.Contains(sheet.Object, adapter.Released);
    }

    [Fact]
    public void A_non_worksheet_entry_is_skipped_without_being_cast()
    {
        var adapter = new RecordingAdapter();
        adapter.Sheets.Add(new object());

        bool found = adapter.TryLocateTable(
            new Mock<Excel.Sheets>().Object,
            out Excel.Worksheet? worksheet,
            out Excel.ListObject? table);

        Assert.False(found);
        Assert.Null(worksheet);
        Assert.Null(table);
        Assert.Empty(adapter.Released);
    }

    [Fact]
    public void Null_sheets_throws()
    {
        var adapter = new RecordingAdapter();

        ArgumentNullException thrown = Assert.Throws<ArgumentNullException>(() =>
        {
            adapter.TryLocateTable(null!, out _, out _);
        });

        Assert.Equal("sheets", thrown.ParamName);
    }
}
