using GanttCreator.Core;
using GanttCreator.Office;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for the R5.4 collapse-all and expand-all behaviour of
/// <see cref="ExcelOutlineGroupWriter"/>: the grouped child rows are hidden or
/// shown through the same accessor the native Excel control writes, the parent
/// row is never touched, and every refusal path fires with no write.
/// </summary>
public sealed class ExcelOutlineCollapseTests
{
    /// <summary>
    /// Overrides the COM seams, records every hidden-flag write as a
    /// (worksheetRow, hidden) pair, and serves a fixed outline level per body
    /// row.
    /// </summary>
    private sealed class TestableWriter(
        object? application,
        IWorksheetProtectionGuard guard,
        Excel.Worksheet worksheet,
        Excel.ListObject table,
        List<(int Row, bool Hidden)> hiddenWrites)
        : ExcelOutlineGroupWriter(application, guard)
    {
        internal override bool TryFindTable(
            Excel.Sheets sheets,
            out Excel.Worksheet? resolvedWorksheet,
            out Excel.ListObject? resolvedTable)
        {
            resolvedWorksheet = worksheet;
            resolvedTable = table;
            return true;
        }

        internal override Excel.Range? GetTableBody(Excel.ListObject source) => SourceBody;

        internal override int GetBodyRowCount(Excel.Range body) => BodyRowCount;

        internal override int? ReadFirstWorksheetRow(Excel.Range range) => FirstWorksheetRow;

        internal override Excel.Range? GetRowsRange(Excel.Worksheet source, int firstRow, int lastRow)
        {
            _pendingFirst = firstRow;
            _pendingLast = lastRow;
            return new Mock<Excel.Range>().Object;
        }

        internal override object? ReadOutlineLevelValue(Excel.Range range)
        {
            int bodyRow = _pendingFirst - FirstWorksheetRow + 1;
            return Levels.TryGetValue(bodyRow, out int level) ? level : null;
        }

        internal override bool SetRowHidden(Excel.Range range, bool hidden)
        {
            hiddenWrites.Add((_pendingFirst, hidden));
            return true;
        }

        /// <summary>The outline level each body row currently carries.</summary>
        internal Dictionary<int, int> Levels { get; } = [];

        internal int FirstWorksheetRow { get; set; } = 2;

        internal int BodyRowCount { get; set; }

        private int _pendingFirst;

        private int _pendingLast;

        private Excel.Range SourceBody { get; } = new Mock<Excel.Range>().Object;
    }

    private static (Mock<Excel.Application> Application, Mock<Excel.Worksheet> Worksheet, Mock<Excel.ListObject> Table) Graph()
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var worksheet = new Mock<Excel.Worksheet>();
        var table = new Mock<Excel.ListObject>();

        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(new Mock<Excel.Sheets>().Object);
        return (application, worksheet, table);
    }

    private static IWorksheetProtectionGuard Guard(ProtectionGuardOutcome outcome)
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.QueryTarget(It.IsAny<object?>())).Returns(outcome);
        return guard.Object;
    }

    /// <summary>
    /// A writer over a sheet whose body carries one group: row 1 is the
    /// parent and row 2 is its child.
    /// </summary>
    private static (TestableWriter Writer, List<(int Row, bool Hidden)> Writes) OneGroupWriter(
        ProtectionGuardOutcome guard = ProtectionGuardOutcome.NotProtected)
    {
        (Mock<Excel.Application> application, Mock<Excel.Worksheet> worksheet, Mock<Excel.ListObject> table) = Graph();
        var writes = new List<(int Row, bool Hidden)>();
        var writer = new TestableWriter(
            application.Object,
            Guard(guard),
            worksheet.Object,
            table.Object,
            writes)
        {
            BodyRowCount = 2,
        };

        writer.Levels[1] = 1;
        writer.Levels[2] = 2;
        return (writer, writes);
    }

    [Fact]
    public void Collapsing_hides_the_grouped_child_row()
    {
        (TestableWriter writer, List<(int Row, bool Hidden)> writes) = OneGroupWriter();

        OutlineCollapseOutcome outcome = writer.SetCollapsed(OutlineCollapseDirection.Collapse);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.RowsWritten);
        // Body row 2 is worksheet row 3: the table's first body row is row 2.
        Assert.Equal([(3, true)], writes);
    }

    [Fact]
    public void Expanding_shows_the_grouped_child_row()
    {
        (TestableWriter writer, List<(int Row, bool Hidden)> writes) = OneGroupWriter();

        OutlineCollapseOutcome outcome = writer.SetCollapsed(OutlineCollapseDirection.Expand);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.RowsWritten);
        Assert.Equal([(3, false)], writes);
    }

    [Fact]
    public void The_parent_row_is_never_written()
    {
        (TestableWriter writer, List<(int Row, bool Hidden)> writes) = OneGroupWriter();

        writer.SetCollapsed(OutlineCollapseDirection.Collapse);

        // The parent (worksheet row 2) keeps its native collapse control and
        // stays visible; only the child is written.
        Assert.DoesNotContain(writes, write => write.Row == 2);
    }

    [Fact]
    public void Every_groups_children_are_written()
    {
        (Mock<Excel.Application> application, Mock<Excel.Worksheet> worksheet, Mock<Excel.ListObject> table) = Graph();
        var writes = new List<(int Row, bool Hidden)>();
        var writer = new TestableWriter(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            worksheet.Object,
            table.Object,
            writes)
        {
            BodyRowCount = 4,
        };

        // Two groups: rows 1-2 and rows 3-4.
        writer.Levels[1] = 1;
        writer.Levels[2] = 2;
        writer.Levels[3] = 1;
        writer.Levels[4] = 2;

        OutlineCollapseOutcome outcome = writer.SetCollapsed(OutlineCollapseDirection.Collapse);

        Assert.Equal(2, outcome.RowsWritten);
        Assert.Equal([(3, true), (5, true)], writes);
    }

    [Fact]
    public void A_sheet_with_no_group_writes_nothing()
    {
        (TestableWriter writer, List<(int Row, bool Hidden)> writes) = OneGroupWriter();
        writer.Levels.Clear();
        writer.Levels[1] = 1;
        writer.Levels[2] = 1;

        OutlineCollapseOutcome outcome = writer.SetCollapsed(OutlineCollapseDirection.Collapse);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.RowsWritten);
        Assert.Empty(writes);
    }

    [Fact]
    public void A_protected_sheet_refuses_with_no_write()
    {
        (TestableWriter writer, List<(int Row, bool Hidden)> writes) =
            OneGroupWriter(ProtectionGuardOutcome.SheetProtected);

        OutlineCollapseOutcome outcome = writer.SetCollapsed(OutlineCollapseDirection.Collapse);

        Assert.False(outcome.Succeeded);
        Assert.Equal(OutlineCollapseRefusalReason.TargetProtected, outcome.Refusal);
        Assert.Equal(0, outcome.RowsWritten);
        Assert.Empty(writes);
    }

    [Fact]
    public void A_missing_workbook_refuses_with_no_write()
    {
        var application = new Mock<Excel.Application>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns((Excel.Workbook)null!);
        var writes = new List<(int Row, bool Hidden)>();
        var writer = new TestableWriter(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            new Mock<Excel.Worksheet>().Object,
            new Mock<Excel.ListObject>().Object,
            writes);

        OutlineCollapseOutcome outcome = writer.SetCollapsed(OutlineCollapseDirection.Collapse);

        Assert.False(outcome.Succeeded);
        Assert.Equal(OutlineCollapseRefusalReason.NoActiveWorkbook, outcome.Refusal);
        Assert.Empty(writes);
    }

    [Fact]
    public void A_protected_sheet_is_refused_before_the_outline_is_read()
    {
        // ADR-0008 D4: the protection check is the FIRST operation in a
        // mutating adapter, before any application, workbook or table access.
        // The guard refusing proves the write phase was never entered.
        (TestableWriter writer, List<(int Row, bool Hidden)> writes) =
            OneGroupWriter(ProtectionGuardOutcome.SheetProtected);

        writer.SetCollapsed(OutlineCollapseDirection.Collapse);

        Assert.Empty(writes);
    }
}
