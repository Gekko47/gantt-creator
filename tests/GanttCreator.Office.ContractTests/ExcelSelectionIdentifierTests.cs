using System.Runtime.InteropServices;
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
        public Mock<Excel.Chart> Chart { get; } = new();
        public bool HasSelection { get; set; } = true;

        /// <summary>
        /// False models a chart-sheet or shape selection: <c>Application.Selection</c>
        /// returns a live object that is not a <c>Range</c>.
        /// </summary>
        public bool SelectionIsRange { get; set; } = true;

        /// <summary>
        /// True models a chart sheet being active. Excel has no active cell there,
        /// so <c>Application.ActiveCell</c> raises a COM error instead of returning
        /// a range (state matrix, 2026-11-13).
        /// </summary>
        public bool ActiveCellProbeFails { get; set; }

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

        internal override Excel.Range? GetActiveCell(Excel.Application application)
        {
            if (_graph.ActiveCellProbeFails)
            {
                // COMException is constructed through Marshal, never with `new`
                // (CA2201 reserves it): a chart sheet has no active cell, so the
                // live host raises an application-defined error here.
                throw ComFailure(unchecked((int)0x800A03EC));
            }

            return _graph.ActiveCell.Object;
        }

        internal override object? GetSelection(Excel.Application application)
        {
            if (!_graph.HasSelection)
            {
                return null;
            }

            return _graph.SelectionIsRange ? _graph.Selection.Object : _graph.Chart.Object;
        }

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
    public void Identify_reports_the_active_cell_row_inside_a_multi_row_selection()
    {
        // R5.9b: the structural commands act on the row under the ACTIVE cell,
        // so a multi-row selection must still name that row; the bespoke-edit
        // facts stay false, and the two answers do not collide.
        Graph graph = Seeded(bodyRowCount: 3, activeRow: 3, selectionRow: 2, selectionCount: 2);

        SelectionFacts facts = graph.Build().Identify();

        Assert.False(facts.IsSingleBodyRow);
        Assert.Equal(2, facts.ActiveCellBodyIndex);
    }

    [Fact]
    public void Identify_reports_no_active_cell_row_when_the_active_cell_is_the_header()
    {
        Graph graph = Seeded(bodyRowCount: 3, activeRow: 1, selectionRow: 1, selectionCount: 1);

        SelectionFacts facts = graph.Build().Identify();

        Assert.True(facts.IsHeaderRow);
        Assert.Null(facts.ActiveCellBodyIndex);
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

    [Fact]
    public void Identify_reports_none_when_the_active_workbook_is_closed()
    {
        // State matrix, closed workbook (2026-11-13). Excel reports "no active
        // workbook" as null, not as a stale proxy.
        Graph graph = Seeded(bodyRowCount: 3, activeRow: 3, selectionRow: 3, selectionCount: 1);
        _ = graph.Application.SetupGet(a => a.ActiveWorkbook).Returns((Excel.Workbook)null!);

        SelectionFacts facts = graph.Build().Identify();

        Assert.Equal(SelectionFacts.None, facts);
    }

    [Fact]
    public void Identify_reports_none_when_the_active_workbook_probe_fails()
    {
        // State matrix, closed workbook (2026-11-13). A workbook that is closing
        // can make ActiveWorkbook throw rather than return null; the port is
        // documented total ("Never throws"), so the catch must absorb it.
        Graph graph = Seeded(bodyRowCount: 3, activeRow: 3, selectionRow: 3, selectionCount: 1);
        _ = graph.Application.SetupGet(a => a.ActiveWorkbook)
            .Throws(ComFailure(unchecked((int)0x80010108)));

        SelectionFacts facts = graph.Build().Identify();

        Assert.Equal(SelectionFacts.None, facts);
    }

    [Fact]
    public void Identify_reports_none_when_a_chart_sheet_is_active()
    {
        // State matrix, chart sheet active (2026-11-13). Excel has no active cell
        // on a chart sheet, so Application.ActiveCell raises a COM error. Reporting
        // None keeps the ribbon on its previous snapshot instead of faulting on a
        // sheet switch -- this is the positive test for the catch in Identify().
        Graph graph = Seeded(bodyRowCount: 3, activeRow: 3, selectionRow: 3, selectionCount: 1);
        graph.SelectionIsRange = false;
        graph.ActiveCellProbeFails = true;

        SelectionFacts facts = graph.Build().Identify();

        Assert.Equal(SelectionFacts.None, facts);
    }

    [Fact]
    public void Identify_falls_back_to_the_active_cell_when_a_non_range_object_is_selected()
    {
        // State matrix, chart/shape selection (2026-11-13). A non-Range selection
        // casts to null, and the active cell still names the row the structural
        // commands act on (R5.9b).
        Graph graph = Seeded(bodyRowCount: 3, activeRow: 3, selectionRow: 3, selectionCount: 1);
        graph.SelectionIsRange = false;

        SelectionFacts facts = graph.Build().Identify();

        Assert.True(facts.HasSelection);
        Assert.True(facts.IsSingleBodyRow);
        Assert.Equal(2, facts.BodyIndex);
    }

    [Fact]
    public void Identify_never_consults_sheet_or_workbook_protection()
    {
        // State matrix, protected sheet (2026-11-13). Protection blocks edits, not
        // reads, so identifying a selection must be protection-independent: if it
        // branched on protection, Remove/structural gating would silently disable
        // itself the moment the user protected the Gantt sheet.
        Graph graph = Seeded(bodyRowCount: 3, activeRow: 3, selectionRow: 3, selectionCount: 1);
        _ = graph.Worksheet.SetupGet(w => w.ProtectContents).Returns(true);
        _ = graph.Worksheet.SetupGet(w => w.ProtectDrawingObjects).Returns(true);
        _ = graph.Workbook.SetupGet(w => w.ProtectStructure).Returns(true);

        SelectionFacts facts = graph.Build().Identify();

        Assert.True(facts.IsSingleBodyRow);
        graph.Worksheet.VerifyGet(w => w.ProtectContents, Times.Never());
        graph.Worksheet.VerifyGet(w => w.ProtectDrawingObjects, Times.Never());
        graph.Workbook.VerifyGet(w => w.ProtectStructure, Times.Never());
    }

    /// <summary>
    /// Produces the exception a failed COM probe raises, without constructing the
    /// runtime-reserved <see cref="COMException"/> directly (CA2201). The same
    /// production is used by the other Office contract tests.
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
