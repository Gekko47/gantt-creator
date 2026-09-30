using GanttCreator.Core;
using GanttCreator.Office;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ExcelOutlineGroupWriter"/> (R4.7D,
/// ADR-0026 D5): outline groups are rebuilt from <c>ParentId</c> alone, the parent
/// row itself is never moved into a group, and every refusal path fires with no
/// write.
/// </summary>
public sealed class ExcelOutlineGroupWriterTests
{
    /// <summary>
    /// Overrides the COM seams, records every outline-level write as a
    /// (firstRow, lastRow, level) triple, and lets the test resolve ranges to
    /// placeholder mocks.
    /// </summary>
    private sealed class TestableWriter(
        object? application,
        IWorksheetProtectionGuard guard,
        Excel.Worksheet worksheet,
        Excel.ListObject table,
        List<(int First, int Last, int Level)> writes)
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

        internal override int? ReadFirstWorksheetRow(Excel.Range range) => FirstWorksheetRow;

        internal override Excel.Range? GetRowsRange(Excel.Worksheet source, int firstRow, int lastRow)
        {
            // The adapter passes the resolved span straight through to
            // WriteOutlineLevel, so the bounds are captured here and asserted by
            // the test. The stub infers nothing about the range.
            _pendingFirst = firstRow;
            _pendingLast = lastRow;
            return new Mock<Excel.Range>().Object;
        }

        internal override bool WriteOutlineLevel(Excel.Range range, int level)
        {
            writes.Add((_pendingFirst, _pendingLast, level));
            return true;
        }

        internal int FirstWorksheetRow { get; set; } = 2;

        private int _pendingFirst;

        private int _pendingLast;

        private Excel.Range SourceBody { get; } = new Mock<Excel.Range>().Object;
    }

    private static GanttRowId NewId() => GanttRowId.New();

    private static GanttEvent Event(int rowNumber, GanttRowId id, GanttRowId? parentId = null) =>
        new(
            rowNumber,
            id,
            NewId(),
            0,
            GanttEntityType.AsPlannedActivity,
            "Build frame",
            new DateOnly(2026, 1, 1),
            new DateOnly(2026, 12, 31),
            parentId,
            "AsPlannedActivity",
            null,
            null,
            null,
            true,
            null);

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
    /// The row span is named by its whole-row A1 address, so the write covers every
    /// column of every row in the group rather than one cell.
    /// </summary>
    /// <remarks>
    /// The adapter's other seams are stubbed, so this pins the part the stub
    /// bypasses. The previous form indexed <c>Rows</c> with two integers, which the
    /// PIA defines as <c>Item(RowIndex, ColumnIndex)</c> -- a single cell -- so a
    /// group of several children outlined one cell and the other rows kept their old
    /// level. A single-row group is the same address as the cell it used to select,
    /// which is why the defect was invisible for a one-child group.
    /// </remarks>
    [Theory]
    [InlineData(3, 4, "3:4")]
    [InlineData(5, 5, "5:5")]
    [InlineData(1, 16384, "1:16384")]
    public void The_row_span_address_covers_every_row_between_first_and_last(
        int firstRow,
        int lastRow,
        string expected)
    {
        Assert.Equal(expected, ExcelOutlineGroupWriter.RowSpanAddress(firstRow, lastRow));
    }

    /// <summary>
    /// A parent with two contiguous children writes exactly one range at the child
    /// outline level. The parent row is excluded, which is what keeps the
    /// <c>-</c>/<c>+</c> control on the parent rather than consuming it.
    /// </summary>
    [Fact]
    public void A_parent_with_contiguous_children_writes_one_child_range()
    {
        var (application, worksheet, table) = Graph();
        List<(int First, int Last, int Level)> writes = [];
        GanttRowId parent = NewId();
        var writer = new TestableWriter(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            worksheet.Object,
            table.Object,
            writes);

        // Body row 1 is worksheet row 2, so body rows 2-3 are worksheet rows 3-4.
        OutlineGroupOutcome outcome = writer.Apply(
            [Event(1, parent), Event(2, NewId(), parent), Event(3, NewId(), parent)]);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.GroupsApplied);
        var write = Assert.Single(writes);
        Assert.Equal(3, write.First);
        Assert.Equal(4, write.Last);
        Assert.Equal(ExcelOutlineGroupWriter.ChildOutlineLevel, write.Level);
    }

    /// <summary>
    /// A flat table writes nothing at all: the common case must not add outline
    /// controls or dirty the workbook.
    /// </summary>
    [Fact]
    public void A_flat_table_writes_nothing()
    {
        var (application, worksheet, table) = Graph();
        List<(int First, int Last, int Level)> writes = [];
        var writer = new TestableWriter(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            worksheet.Object,
            table.Object,
            writes);

        OutlineGroupOutcome outcome = writer.Apply([Event(1, NewId()), Event(2, NewId())]);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.GroupsApplied);
        Assert.Equal(0, outcome.RowsUngrouped);
        Assert.Empty(writes);
    }

    /// <summary>
    /// An empty event list succeeds with nothing written rather than refusing.
    /// </summary>
    [Fact]
    public void An_empty_hierarchy_succeeds_with_no_writes()
    {
        var (application, worksheet, table) = Graph();
        List<(int First, int Last, int Level)> writes = [];
        var writer = new TestableWriter(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            worksheet.Object,
            table.Object,
            writes);

        OutlineGroupOutcome outcome = writer.Apply([]);

        Assert.True(outcome.Succeeded);
        Assert.Empty(writes);
    }

    /// <summary>
    /// A protected target is refused before any write (ADR-0008 D4). The empty
    /// write list proves the refusal preceded the COM call rather than following it.
    /// </summary>
    [Fact]
    public void A_protected_target_is_refused_without_writing()
    {
        var (application, worksheet, table) = Graph();
        List<(int First, int Last, int Level)> writes = [];
        GanttRowId parent = NewId();
        var writer = new TestableWriter(
            application.Object,
            Guard(ProtectionGuardOutcome.SheetProtected),
            worksheet.Object,
            table.Object,
            writes);

        OutlineGroupOutcome outcome = writer.Apply([Event(1, parent), Event(2, NewId(), parent)]);

        Assert.False(outcome.Succeeded);
        Assert.Equal(OutlineGroupRefusalReason.TargetProtected, outcome.Refusal);
        Assert.Empty(writes);
    }

    /// <summary>
    /// A guard reporting no active workbook is distinguished from a protected sheet.
    /// </summary>
    [Fact]
    public void A_missing_workbook_from_the_guard_is_refused_as_such()
    {
        var (application, worksheet, table) = Graph();
        var writer = new TestableWriter(
            application.Object,
            Guard(ProtectionGuardOutcome.NoActiveWorkbook),
            worksheet.Object,
            table.Object,
            []);

        OutlineGroupOutcome outcome = writer.Apply([Event(1, NewId())]);

        Assert.False(outcome.Succeeded);
        Assert.Equal(OutlineGroupRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }

    /// <summary>
    /// A null application object degrades to the no-active-workbook refusal with no
    /// mutation.
    /// </summary>
    [Fact]
    public void A_null_application_is_refused()
    {
        var writer = new TestableWriter(
            null,
            Guard(ProtectionGuardOutcome.NotProtected),
            new Mock<Excel.Worksheet>().Object,
            new Mock<Excel.ListObject>().Object,
            []);

        OutlineGroupOutcome outcome = writer.Apply([Event(1, NewId())]);

        Assert.False(outcome.Succeeded);
        Assert.Equal(OutlineGroupRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }

    /// <summary>
    /// A null event list is a caller bug and throws, so it cannot be mistaken for
    /// "nothing to group".
    /// </summary>
    [Fact]
    public void A_null_hierarchy_throws()
    {
        var (application, worksheet, table) = Graph();
        var writer = new TestableWriter(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            worksheet.Object,
            table.Object,
            []);

        Assert.Throws<ArgumentNullException>(() => writer.Apply(null!));
    }
}
