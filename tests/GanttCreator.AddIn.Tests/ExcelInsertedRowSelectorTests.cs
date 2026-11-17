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
        Func<int, Excel.ListRow> rowAt,
        IReadOnlyDictionary<Excel.Range, int> rangeRows)
        : ExcelInsertedRowSelector(application)
    {
        public List<int> RequestedBodyIndexes { get; } = [];

        /// <summary>
        /// The cells the column-aware path selected, as (worksheet row, worksheet column).
        /// </summary>
        /// <remarks>
        /// A cell select is a different call from a row select, and the whole point of
        /// <c>SelectBodyCell</c> is that the row's full width is NOT selected. Recording
        /// the address is what proves the column came from the active cell.
        /// </remarks>
        public List<(int Row, int Column)> SelectedCells { get; } = [];

        /// <summary>The column the host reports for the active cell, or null for none.</summary>
        public int? ActiveColumn { get; set; }

        internal override int? ReadActiveColumn() => ActiveColumn;

        internal override int GetRangeRow(Excel.Range range) =>
            rangeRows.TryGetValue(range, out int row) ? row : 0;

        internal override Excel.Range GetCell(Excel.Worksheet source, int worksheetRow, int worksheetColumn)
        {
            Assert.Same(worksheet, source);
            SelectedCells.Add((worksheetRow, worksheetColumn));
            Mock<Excel.Range> cell = new();
            _ = cell.Setup(r => r.Select());
            return cell.Object;
        }

        /// <summary>
        /// Whether the list-row lookup seam surfaces the host failure, simulating a
        /// refused COM indexer. The adapter's best-effort guard has to cover the
        /// walk, not only the final <c>Activate</c>/<c>Select</c> pair.
        /// <para>
        /// The thrown exception is the one Moq raises for a configured
        /// <c>Throws</c>, so no <c>new COMException</c> is constructed — CA2201
        /// reserves that type for the runtime.
        /// </para>
        /// </summary>
        public bool FailRowLookup { get; set; }

        /// <summary>
        /// A row whose range read is configured to throw, used only to raise the
        /// host failure the guard has to absorb.
        /// </summary>
        public Mock<Excel.ListRow> FailingRow { get; } = new();

        internal override object GetSheetAt(Excel.Sheets sheets, int index) => worksheet;

        internal override Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index) => table;

        internal override Excel.ListRows GetListRows(Excel.ListObject candidate)
        {
            Assert.Same(table, candidate);
            return new Mock<Excel.ListRows>().Object;
        }

        internal override Excel.ListRow GetListRowAt(Excel.ListRows rows, int bodyIndex)
        {
            if (FailRowLookup)
            {
                // Raise the host failure the same way the other tests inject one:
                // a configured mock throwing through its own member access. If the
                // configuration were ever lost this would return normally and the
                // guard would be proven by nothing, so the test also asserts no
                // selection was made.
                _ = FailingRow.Object.Range;
            }

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

        /// <summary>
        /// The worksheet row each body row's range reports, so the cell address the
        /// selector builds can be asserted against the real layout rather than a constant.
        /// </summary>
        public Dictionary<Excel.Range, int> RangeRows { get; } = [];

        /// <summary>The worksheet row of the table's first body row.</summary>
        private const int _firstBodyRow = 4;

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
                RangeRows[RowRanges[index].Object] = _firstBodyRow + index;
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
                bodyIndex => Rows[bodyIndex - 1].Object,
                RangeRows);
    }

    /// <summary>
    /// A cell selection keeps the user's column and moves only the row.
    /// </summary>
    /// <remarks>
    /// This is the relocation contract: promote, demote, move up and move down all rotate
    /// or re-home content, so the address the user was on holds a different activity
    /// afterwards. Selecting the row's full width would keep the row but throw the column
    /// away; the cell address is what the user expects to still be on.
    /// </remarks>
    [Fact]
    public void SelectBodyCell_selects_the_active_column_of_the_requested_row()
    {
        var graph = new Graph();
        TestableSelector selector = graph.Build();
        selector.ActiveColumn = 3;

        selector.SelectBodyCell(2);

        Assert.Equal(1, graph.Activations);
        Assert.Equal([(5, 3)], selector.SelectedCells);
        foreach (Mock<Excel.Range> rowRange in graph.RowRanges)
        {
            rowRange.Verify(r => r.Select(), Times.Never);
        }
    }

    /// <summary>
    /// A host that reports no active cell falls back to selecting the whole row.
    /// </summary>
    /// <remarks>
    /// The fallback is the behaviour this port had before cell selection existed, so a
    /// host error degrades to a row selection rather than to no selection at all.
    /// </remarks>
    [Fact]
    public void SelectBodyCell_without_an_active_cell_selects_the_whole_row()
    {
        var graph = new Graph();
        TestableSelector selector = graph.Build();
        selector.ActiveColumn = null;

        selector.SelectBodyCell(2);

        Assert.Empty(selector.SelectedCells);
        graph.RowRanges[1].Verify(r => r.Select(), Times.Once);
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
    public void A_host_failure_activating_the_worksheet_degrades_to_no_selection()
    {
        // The row has already been inserted by the time SelectBodyRow runs, so a
        // refused activation must not escape: doing so would surface the insert as
        // a failure to the user even though the row is there.
        var graph = new Graph();
        _ = graph.Worksheet.Setup(w => w.Activate()).Throws<COMException>();
        TestableSelector selector = graph.Build();

        Assert.Null(Record.Exception(() => selector.SelectBodyRow(2)));

        graph.RowRanges[1].Verify(r => r.Select(), Times.Never);
    }

    [Fact]
    public void A_host_failure_selecting_the_row_degrades_to_no_selection()
    {
        // Same guarantee for the select itself, which is the second COM call at
        // this boundary and the one most likely to be refused (a protected or
        // non-active sheet).
        var graph = new Graph();
        _ = graph.RowRanges[1].Setup(r => r.Select()).Throws<COMException>();
        TestableSelector selector = graph.Build();

        Assert.Null(Record.Exception(() => selector.SelectBodyRow(2)));

        Assert.Equal(2, Assert.Single(selector.RequestedBodyIndexes));
    }

    [Fact]
    public void A_host_failure_reading_the_list_row_degrades_to_no_selection()
    {
        // The guard covers the walk, not just the two selection calls. The list-row
        // lookup is a COM indexer and can be refused like any other, and a refusal
        // that escaped would surface the insert as a failure when the row was in
        // fact added.
        var graph = new Graph();
        TestableSelector selector = graph.Build();
        selector.FailRowLookup = true;
        _ = selector.FailingRow.SetupGet(r => r.Range).Throws<COMException>();

        Assert.Null(Record.Exception(() => selector.SelectBodyRow(2)));

        graph.RowRanges[1].Verify(r => r.Select(), Times.Never);
        Assert.Equal(0, graph.Activations);
    }

    [Fact]
    public void A_host_failure_reading_the_table_name_degrades_to_no_selection()
    {
        // Same for the table-name read, which decides whether the walk keeps
        // going. A refusal here is a failure to select, not a failure of the row
        // that has already been inserted.
        var graph = new Graph();
        _ = graph.Table.SetupGet(t => t.Name).Throws<COMException>();
        TestableSelector selector = graph.Build();

        Assert.Null(Record.Exception(() => selector.SelectBodyRow(2)));

        Assert.Empty(selector.RequestedBodyIndexes);
        Assert.Equal(0, graph.Activations);
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
