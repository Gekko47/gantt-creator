using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for the R5.5 standalone selection identifier (owner ruling
/// 2026-10-14): the one-expanded-row facts the ribbon gating consumes, and
/// the inert subscription when no application exists.
/// </summary>
public class ExcelSelectionIdentifierTests
{
    private sealed class Graph
    {
        private readonly Dictionary<Excel.Range, (int Row, int Count)> _spans = new();
        private readonly Dictionary<int, Mock<Excel.Range>> _overlaps = new();

        public Mock<Excel.Application> Application { get; } = new();
        public Mock<Excel.Workbook> Workbook { get; } = new();
        public Mock<Excel.Sheets> Sheets { get; } = new();
        public Mock<Excel.Worksheet> Worksheet { get; } = new();
        public Mock<Excel.ListObjects> ListObjects { get; } = new();
        public Mock<Excel.ListObject> Table { get; } = new();
        public Mock<Excel.Range> HeaderRow { get; } = new();
        public Mock<Excel.Range> DataBody { get; } = new();
        public Mock<Excel.Range> ActiveCell { get; } = new();
        public Mock<Excel.Range> Selection { get; } = new();
        public bool HasSelection { get; set; } = true;

        public Graph()
        {
            Application.SetupGet(a => a.ActiveWorkbook).Returns(Workbook.Object);
            Workbook.SetupGet(w => w.Sheets).Returns(Sheets.Object);
            Worksheet.SetupGet(w => w.ListObjects).Returns(ListObjects.Object);
            ListObjects.SetupGet(lo => lo.Count).Returns(1);
            Table.SetupGet(t => t.Name).Returns("tblGanttData");
            Table.SetupGet(t => t.HeaderRowRange).Returns(HeaderRow.Object);
            Table.SetupGet(t => t.DataBodyRange).Returns(DataBody.Object);
        }

        /// <summary>Registers a range as [row, row + count - 1].</summary>
        public Excel.Range Register(Excel.Range range, int row, int count)
        {
            _spans[range] = (row, count);
            return range;
        }

        public (int Row, int Count) Span(Excel.Range range) => _spans[range];

        /// <summary>Returns the tracked overlap range starting at a row.</summary>
        public Excel.Range OverlapAt(int row)
        {
            if (!_overlaps.TryGetValue(row, out Mock<Excel.Range>? value))
            {
                value = new Mock<Excel.Range>();
                value.SetupGet(r => r.Row).Returns(row);
                _overlaps[row] = value;
            }

            _spans[value.Object] = (row, 1);
            return value.Object;
        }

        public TestableIdentifier Build() => new(Application.Object, this);
    }

    private sealed class TestableIdentifier(Excel.Application? application, Graph graph)
        : ExcelSelectionIdentifier(application)
    {
        private readonly Graph _graph = graph;

        internal override int GetSheetCount(Excel.Sheets sheets) => 1;

        internal override object GetSheetAt(Excel.Sheets sheets, int index) => _graph.Worksheet.Object;

        internal override Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index) =>
            _graph.Table.Object;

        internal override Excel.Range? GetHeaderRow(Excel.ListObject table) => _graph.HeaderRow.Object;

        internal override Excel.Range? GetDataBodyRange(Excel.ListObject table) => _graph.DataBody.Object;

        internal override Excel.Range? GetActiveCell(Excel.Application application) =>
            _graph.ActiveCell.Object;

        internal override object? GetSelection(Excel.Application application) =>
            _graph.HasSelection ? _graph.Selection.Object : null;

        internal override int GetRangeRow(Excel.Range range) => _graph.Span(range).Row;

        internal override int GetRowCount(Excel.Range range) => _graph.Span(range).Count;

        internal override Excel.Range? Intersect(Excel.Range first, Excel.Range second)
        {
            (int Row, int Count) a = _graph.Span(first);
            (int Row, int Count) b = _graph.Span(second);
            var start = Math.Max(a.Row, b.Row);
            var end = Math.Min(a.Row + a.Count - 1, b.Row + b.Count - 1);
            return start > end ? null : _graph.OverlapAt(start);
        }
    }

    private static Graph Seeded(int bodyRowCount, int activeRow, int selectionRow, int selectionCount)
    {
        var graph = new Graph();
        graph.Register(graph.HeaderRow.Object, row: 1, count: 1);
        graph.Register(graph.DataBody.Object, row: 2, count: bodyRowCount);
        graph.Register(graph.ActiveCell.Object, row: activeRow, count: 1);
        graph.Register(graph.Selection.Object, row: selectionRow, count: selectionCount);
        return graph;
    }

    [Fact]
    public void Identify_reports_one_expanded_body_row()
    {
        Graph graph = Seeded(bodyRowCount: 3, activeRow: 3, selectionRow: 3, selectionCount: 1);

        SelectionFacts facts = graph.Build().Identify();

        Assert.True(facts.HasSelection);
        Assert.True(facts.IsSingleBodyRow);
        Assert.Equal(2, facts.BodyIndex);
        Assert.False(facts.IsHeaderRow);
        Assert.Equal(1, facts.SelectedRowCount);
    }

    [Fact]
    public void Identify_reports_a_multi_row_selection_as_not_single()
    {
        // Ruling 1: a multi-row selection must never enable a bespoke edit.
        Graph graph = Seeded(bodyRowCount: 3, activeRow: 3, selectionRow: 3, selectionCount: 2);

        SelectionFacts facts = graph.Build().Identify();

        Assert.True(facts.HasSelection);
        Assert.False(facts.IsSingleBodyRow);
        Assert.Equal(2, facts.SelectedRowCount);
    }

    [Fact]
    public void Identify_reports_the_header_row_as_header()
    {
        Graph graph = Seeded(bodyRowCount: 3, activeRow: 1, selectionRow: 1, selectionCount: 1);

        SelectionFacts facts = graph.Build().Identify();

        Assert.True(facts.HasSelection);
        Assert.True(facts.IsHeaderRow);
        Assert.False(facts.IsSingleBodyRow);
    }

    [Fact]
    public void Identify_reports_none_when_the_selection_misses_the_table()
    {
        Graph graph = Seeded(bodyRowCount: 3, activeRow: 3, selectionRow: 3, selectionCount: 1);
        graph.Register(graph.Selection.Object, row: 10, count: 1);

        SelectionFacts facts = graph.Build().Identify();

        Assert.False(facts.HasSelection);
        Assert.False(facts.IsSingleBodyRow);
        Assert.Null(facts.BodyIndex);
    }

    [Fact]
    public void Identify_reports_none_without_an_application()
    {
        SelectionFacts facts = new ExcelSelectionIdentifier(null).Identify();

        Assert.False(facts.HasSelection);
        Assert.Equal(SelectionFacts.None, facts);
    }

    [Fact]
    public void SubscribeSelectionChanged_is_inert_without_an_application()
    {
        var identifier = new ExcelSelectionIdentifier(null);
        var invoked = false;

        IDisposable subscription = identifier.SubscribeSelectionChanged(() => invoked = true);
        subscription.Dispose();
        subscription.Dispose();

        Assert.False(invoked);
    }

    [Fact]
    public void SubscribeSelectionChanged_rejects_a_null_handler()
    {
        var identifier = new ExcelSelectionIdentifier(null);

        Assert.Throws<ArgumentNullException>(() => identifier.SubscribeSelectionChanged(null!));
    }
}
