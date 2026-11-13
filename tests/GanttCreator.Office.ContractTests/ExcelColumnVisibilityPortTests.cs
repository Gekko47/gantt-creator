using GanttCreator.Core;
using GanttCreator.Office;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ExcelColumnVisibilityPort"/> (R5.18h): the read reports
/// the live state, the write touches only drifted columns, and every refusal path fires
/// with no write.
/// </summary>
/// <remarks>
/// <see cref="ExcelColumnVisibilityPort.Apply"/> runs whenever the user picks an item in
/// the ribbon's dropdown, so a write it does not need marks the workbook dirty and a
/// write it does need but skips leaves a hidden column visible. Both directions are
/// pinned below: the idempotence count, and the unhide of a column dropped from the set.
/// </remarks>
public sealed class ExcelColumnVisibilityPortTests
{
    /// <summary>
    /// Overrides the COM seams so a test can supply a table, control each column's live
    /// hidden flag, and record the writes with no live Excel host.
    /// </summary>
    /// <param name="application">The application object.</param>
    /// <param name="guard">The protection guard.</param>
    /// <param name="table">The table, or null when the table is missing.</param>
    /// <param name="worksheet">The resolved worksheet.</param>
    /// <param name="hidden">Each offered column's live hidden flag.</param>
    /// <param name="writes">The flag writes the test double records.</param>
    /// <param name="liveColumnCount">How many columns the fake live table reports.</param>
    private sealed class TestablePort(
        object? application,
        IWorksheetProtectionGuard guard,
        Excel.ListObject? table,
        Excel.Worksheet? worksheet,
        Dictionary<string, bool> hidden,
        List<string> writes,
        int liveColumnCount = 5)
        : ExcelColumnVisibilityPort(application, guard)
    {
        internal override bool TryFindGanttTable(
            Excel.Sheets sheets,
            out Excel.Worksheet? resolvedWorksheet,
            out Excel.ListObject? resolvedTable)
        {
            resolvedTable = table;
            resolvedWorksheet = table is null ? null : worksheet;
            return table is not null;
        }

        internal override Excel.ListColumns GetTableColumns(Excel.ListObject target)
        {
            Mock<Excel.ListColumns> columns = new();
            columns.Setup(c => c.Count).Returns(liveColumnCount);
            return columns.Object;
        }

        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index)
        {
            string name = GanttColumnVisibility.ToggleableColumnNames[index - 1];

            Mock<Excel.Range> entire = new();
            entire.SetupGet(r => r.Hidden).Returns(hidden[name]);
            entire.SetupSet(r => r.Hidden = It.IsAny<object>())
                .Callback((object value) => writes.Add(name + ":" + value));

            Mock<Excel.Range> range = new();
            range.SetupGet(r => r.EntireColumn).Returns(entire.Object);

            Mock<Excel.ListColumn> column = new();
            column.SetupGet(c => c.Name).Returns(name);
            column.SetupGet(c => c.Range).Returns(range.Object);
            return column.Object;
        }

        internal override Excel.Range GetEntireColumn(Excel.Range range) => range.EntireColumn;
    }

    /// <summary>Every offered column's live flag, with the named ones hidden.</summary>
    /// <param name="hiddenNames">The columns that start hidden.</param>
    /// <returns>The live flag map.</returns>
    private static Dictionary<string, bool> Flags(params string[] hiddenNames)
    {
        Dictionary<string, bool> flags = new(StringComparer.Ordinal);
        foreach (string name in GanttColumnVisibility.ToggleableColumnNames)
        {
            flags[name] = hiddenNames.Contains(name, StringComparer.Ordinal);
        }

        return flags;
    }

    private static Mock<Excel.Application> Application()
    {
        Mock<Excel.Application> application = new();
        Mock<Excel.Workbook> workbook = new();
        Mock<Excel.Sheets> sheets = new();
        workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);
        application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        return application;
    }

    private static Mock<IWorksheetProtectionGuard> Guard(ProtectionGuardOutcome outcome)
    {
        Mock<IWorksheetProtectionGuard> guard = new();
        guard.Setup(g => g.Query()).Returns(outcome);
        guard.Setup(g => g.QueryTarget(It.IsAny<Excel.Worksheet>())).Returns(outcome);
        return guard;
    }

    private static TestablePort Port(
        Dictionary<string, bool> hidden,
        List<string> writes,
        Excel.ListObject? table = null,
        ProtectionGuardOutcome protection = ProtectionGuardOutcome.NotProtected,
        int liveColumnCount = 5) =>
        new(
            Application().Object,
            Guard(protection).Object,
            table ?? new Mock<Excel.ListObject>().Object,
            new Mock<Excel.Worksheet>().Object,
            hidden,
            writes,
            liveColumnCount);

    /// <summary>The read reports exactly the offered columns the worksheet hides.</summary>
    [Fact]
    public void ReadHidden_reports_only_the_hidden_offered_columns()
    {
        List<string> writes = [];
        TestablePort port = Port(Flags("Start"), writes);

        ColumnVisibilityPortOutcome outcome = port.ReadHidden();

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.ColumnsWritten);
        Assert.Equal("Start", Assert.Single(outcome.Hidden!));
        Assert.Empty(writes);
    }

    /// <summary>The read refuses rather than reporting "nothing hidden" with no workbook.</summary>
    [Fact]
    public void ReadHidden_refuses_when_there_is_no_active_workbook()
    {
        List<string> writes = [];
        TestablePort port = new(
            new Mock<Excel.Application>().Object,
            Guard(ProtectionGuardOutcome.NotProtected).Object,
            new Mock<Excel.ListObject>().Object,
            new Mock<Excel.Worksheet>().Object,
            Flags(),
            writes);

        ColumnVisibilityPortOutcome outcome = port.ReadHidden();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ColumnVisibilityRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }

    /// <summary>
    /// The read refuses with a typed reason when the data table is absent, so the caller
    /// cannot mistake "no table" for "nothing hidden".
    /// </summary>
    [Fact]
    public void ReadHidden_refuses_when_the_table_is_missing()
    {
        List<string> writes = [];
        TestablePort port = new(
            Application().Object,
            Guard(ProtectionGuardOutcome.NotProtected).Object,
            null,
            null,
            Flags(),
            writes);

        ColumnVisibilityPortOutcome outcome = port.ReadHidden();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ColumnVisibilityRefusalReason.TableMissing, outcome.Refusal);
    }

    /// <summary>
    /// Picking an item writes that one column and leaves the other four alone, so the
    /// workbook is not dirtied for a choice that changed nothing.
    /// </summary>
    [Fact]
    public void Apply_writes_only_the_column_whose_state_differs()
    {
        List<string> writes = [];
        TestablePort port = Port(Flags(), writes);

        ColumnVisibilityPortOutcome outcome = port.Apply(new HashSet<string>(["Description"], StringComparer.Ordinal));

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.ColumnsWritten);
        Assert.Equal("Description:True", Assert.Single(writes));
    }

    /// <summary>Applying the state the sheet already carries writes nothing.</summary>
    [Fact]
    public void Apply_leaves_an_identical_sheet_untouched()
    {
        List<string> writes = [];
        TestablePort port = Port(Flags("Type"), writes);

        ColumnVisibilityPortOutcome outcome = port.Apply(new HashSet<string>(["Type"], StringComparer.Ordinal));

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.ColumnsWritten);
        Assert.Empty(writes);
    }

    /// <summary>
    /// A column dropped from the set is unhidden — without this direction a toggle could
    /// only ever hide, and the column would never come back.
    /// </summary>
    [Fact]
    public void Apply_unhides_a_column_dropped_from_the_set()
    {
        List<string> writes = [];
        TestablePort port = Port(Flags("Type"), writes);

        ColumnVisibilityPortOutcome outcome = port.Apply(new HashSet<string>(StringComparer.Ordinal));

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.ColumnsWritten);
        Assert.Equal("Type:False", Assert.Single(writes));
        Assert.Empty(outcome.Hidden!);
    }

    /// <summary>
    /// A name outside the Core contract is refused before any write, so one bad name
    /// cannot leave the sheet half-applied.
    /// </summary>
    [Fact]
    public void Apply_refuses_a_column_the_dropdown_does_not_offer()
    {
        List<string> writes = [];
        TestablePort port = Port(Flags(), writes);

        ColumnVisibilityPortOutcome outcome = port.Apply(new HashSet<string>(["Id"], StringComparer.Ordinal));

        Assert.False(outcome.Succeeded);
        Assert.Equal(ColumnVisibilityRefusalReason.UnknownColumn, outcome.Refusal);
        Assert.Empty(writes);
    }

    /// <summary>A protected target is refused with no write (ADR-0008 D4).</summary>
    [Fact]
    public void Apply_refuses_a_protected_target_without_writing()
    {
        List<string> writes = [];
        TestablePort port = Port(Flags(), writes, protection: ProtectionGuardOutcome.SheetProtected);

        ColumnVisibilityPortOutcome outcome = port.Apply(new HashSet<string>(["Type"], StringComparer.Ordinal));

        Assert.False(outcome.Succeeded);
        Assert.Equal(ColumnVisibilityRefusalReason.TargetProtected, outcome.Refusal);
        Assert.Empty(writes);
    }

    /// <summary>
    /// A table missing an offered column refuses with NOTHING written: every column is
    /// resolved before the first flag is set, so a broken table cannot produce a
    /// partially applied choice.
    /// </summary>
    [Fact]
    public void Apply_refuses_when_the_table_lacks_an_offered_column()
    {
        List<string> writes = [];
        TestablePort port = Port(Flags(), writes, liveColumnCount: 4);

        ColumnVisibilityPortOutcome outcome = port.Apply(new HashSet<string>(["Type"], StringComparer.Ordinal));

        Assert.False(outcome.Succeeded);
        Assert.Equal(ColumnVisibilityRefusalReason.UnknownColumn, outcome.Refusal);
        Assert.Empty(writes);
    }

    /// <summary>The write refuses with no workbook rather than throwing.</summary>
    [Fact]
    public void Apply_refuses_when_there_is_no_active_workbook()
    {
        List<string> writes = [];
        TestablePort port = new(
            new Mock<Excel.Application>().Object,
            Guard(ProtectionGuardOutcome.NotProtected).Object,
            new Mock<Excel.ListObject>().Object,
            new Mock<Excel.Worksheet>().Object,
            Flags(),
            writes);

        ColumnVisibilityPortOutcome outcome = port.Apply(new HashSet<string>(["Type"], StringComparer.Ordinal));

        Assert.False(outcome.Succeeded);
        Assert.Equal(ColumnVisibilityRefusalReason.NoActiveWorkbook, outcome.Refusal);
        Assert.Empty(writes);
    }

    /// <summary>A null set is a caller defect and is rejected before any COM call.</summary>
    [Fact]
    public void Apply_rejects_a_null_set() =>
        _ = Assert.Throws<ArgumentNullException>(
            () => Port(Flags(), []).Apply(null!));
}
