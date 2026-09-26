using System.Runtime.InteropServices;
using GanttCreator.Core;
using GanttCreator.Office;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contract tests for the Excel inserted-row selector. The audit findings were
/// that the selector derived the target row from an absolute worksheet row — which
/// agrees with the body index only when the table happens to start at row 1 — and
/// that it selected the row without activating its worksheet first.
/// </summary>
public class ExcelInsertedRowSelectorTests
{
    /// <summary>
    /// The selector with its COM indexer seams substituted; the requested body
    /// index and the activation are recorded so the assertions observe the
    /// behaviour rather than a proxy.
    /// </summary>
    private sealed class TestableSelector(
        object? application,
        Excel.Worksheet worksheet,
        Excel.ListObject table,
        Func<int, Excel.ListRow> rowAt)
        : ExcelInsertedRowSelector(application)
    {
        public List<int> RequestedBodyIndexes { get; } = [];

        internal override object GetSheetAt(Excel.Sheets sheets, int index) => worksheet;

        internal override Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index) => table;

        internal override Excel.ListRows GetListRows(Excel.ListObject candidate)
        {
            Assert.Same(table, candidate);
            return new Mock<Excel.ListRows>().Object;
        }

        internal override Excel.ListRow GetListRowAt(Excel.ListRows rows, int bodyIndex)
        {
            RequestedBodyIndexes.Add(bodyIndex);
            return rowAt(bodyIndex);
        }

        internal override Excel.Range GetRowRange(Excel.ListRow row) => row.Range;
    }

    private sealed class Graph
    {
        public Mock<Excel.Application> Application { get; } = new();

        public Mock<Excel.Worksheet> Worksheet { get; } = new();

        public Mock<Excel.ListObject> Table { get; } = new();

        public Mock<Excel.ListRow>[] Rows { get; } = [new(), new(), new(), new()];

        public Mock<Excel.Range>[] RowRanges { get; } = [new(), new(), new(), new()];

        public int Activations { get; private set; }

        /// <summary>
        /// One ordered record of the calls the selector makes. Activation and row
        /// selection share this single sequence so their relative order is
        /// observable, not just their counts.
        /// </summary>
        public List<string> Calls { get; } = [];

        public Graph()
        {
            var workbook = new Mock<Excel.Workbook>();
            var sheets = new Mock<Excel.Sheets>();
            var listObjects = new Mock<Excel.ListObjects>();

            _ = Application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
            _ = workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);
            _ = sheets.SetupGet(s => s.Count).Returns(1);
            _ = Worksheet.SetupGet(w => w.ListObjects).Returns(listObjects.Object);
            _ = listObjects.SetupGet(l => l.Count).Returns(1);
            _ = Table.SetupGet(t => t.Name).Returns(GanttTableSchema.TableName);
            _ = Worksheet
                .Setup(w => w.Activate())
                .Callback(() =>
                {
                    Activations++;
                    Calls.Add("activate");
                });

            for (var index = 0; index < Rows.Length; index++)
            {
                _ = Rows[index].SetupGet(r => r.Range).Returns(RowRanges[index].Object);
                var bodyIndex = index + 1;
                _ = RowRanges[index]
                    .Setup(r => r.Select())
                    .Callback(() => Calls.Add($"select:{bodyIndex}"));
            }
        }

        public TestableSelector Build() =>
            new(
                Application.Object,
                Worksheet.Object,
                Table.Object,
                bodyIndex => Rows[bodyIndex - 1].Object);
    }

    [Fact]
    public void Selects_the_requested_table_relative_body_row()
    {
        // Body row 3 is ListRows[3], whatever worksheet row the table occupies.
        var graph = new Graph();
        TestableSelector selector = graph.Build();

        selector.SelectBodyRow(3);

        Assert.Equal(3, Assert.Single(selector.RequestedBodyIndexes));
        graph.RowRanges[2].Verify(r => r.Select(), Times.Once);
        graph.RowRanges[0].Verify(r => r.Select(), Times.Never);
        graph.RowRanges[1].Verify(r => r.Select(), Times.Never);
        graph.RowRanges[3].Verify(r => r.Select(), Times.Never);
    }

    [Fact]
    public void Activates_the_worksheet_before_selecting()
    {
        // Range.Select only works on the active sheet, so the Gantt worksheet
        // must be activated first. Asserting the two counts separately would
        // pass even with the order reversed, so the single recorded sequence is
        // what proves activation came first.
        var graph = new Graph();
        TestableSelector selector = graph.Build();

        selector.SelectBodyRow(2);

        Assert.Equal(1, graph.Activations);
        graph.RowRanges[1].Verify(r => r.Select(), Times.Once);
        Assert.Equal(["activate", "select:2"], graph.Calls);
        Assert.True(
            graph.Calls.IndexOf("activate") < graph.Calls.IndexOf("select:2"),
            "Worksheet activation must precede the row selection.");
    }

    [Fact]
    public void Does_nothing_without_an_active_workbook()
    {
        var graph = new Graph();
        _ = graph.Application.SetupGet(a => a.ActiveWorkbook).Returns((Excel.Workbook)null!);
        TestableSelector selector = graph.Build();

        Assert.Null(Record.Exception(() => selector.SelectBodyRow(1)));

        Assert.Empty(selector.RequestedBodyIndexes);
        Assert.Equal(0, graph.Activations);
        graph.RowRanges[0].Verify(r => r.Select(), Times.Never);
    }

    [Fact]
    public void A_host_failure_reading_the_active_workbook_degrades_to_no_selection()
    {
        // The workbook is captured in the constructor, which runs *before* the
        // insert. A host failure there must therefore be contained: the row is
        // still added, only the cursor stays put, and nothing escapes into Excel.
        var graph = new Graph();
        _ = graph
            .Application.SetupGet(a => a.ActiveWorkbook)
            .Throws<COMException>();

        TestableSelector? selector = null;
        Assert.Null(Record.Exception(() => selector = graph.Build()));
        Assert.NotNull(selector);

        Assert.Null(Record.Exception(() => selector!.SelectBodyRow(2)));

        Assert.Empty(selector!.RequestedBodyIndexes);
        Assert.Equal(0, graph.Activations);
        graph.RowRanges[1].Verify(r => r.Select(), Times.Never);
    }
}
