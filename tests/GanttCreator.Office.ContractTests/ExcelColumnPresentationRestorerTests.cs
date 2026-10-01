using GanttCreator.Core;
using GanttCreator.Office;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ExcelColumnPresentationRestorer"/> (R4.7C D1/D2,
/// R4.8A D2): a correct sheet writes nothing, and every refusal path fires with no
/// write.
/// </summary>
/// <remarks>
/// <para>
/// <b>Idempotence is the load-bearing claim.</b> This adapter runs on every Refresh,
/// so an implementation that rewrote a column whose flag already matched would mark
/// the workbook dirty every time — a defect a user meets as a workbook that always
/// asks to save. <c>Range.Hidden</c> is declared <see cref="object"/> in the PIA, so
/// a naive <c>!= true</c> compares boxed references and is always unequal. That is
/// exactly the bug the write-count assertion below pins.
/// </para>
/// </remarks>
public sealed class ExcelColumnPresentationRestorerTests
{
    /// <summary>
    /// Overrides the COM seams so the test can supply a table and record writes with
    /// no live Excel host.
    /// </summary>
    /// <param name="application">The application object.</param>
    /// <param name="guard">The protection guard.</param>
    /// <param name="table">The table, or null when the table is missing.</param>
    /// <param name="worksheet">The resolved worksheet.</param>
    /// <param name="recorded">The flag writes the test double records.</param>
    /// <remarks>
    /// Not sealed, so <c>TestableRestorerWithDrift</c> can override
    /// <c>GetColumnAt</c> and produce a genuinely wrong flag to repair.
    /// </remarks>
    private class TestableRestorer(
        object? application,
        IWorksheetProtectionGuard guard,
        Excel.ListObject? table,
        Excel.Worksheet? worksheet,
        List<string> recorded)
        : ExcelColumnPresentationRestorer(application, guard)
    {
        /// <summary>The shared write log, so a subclass records into the test's list.</summary>
        protected List<string> Recorded { get; } = recorded;

        internal override bool TryFindGanttTable(
            Excel.Sheets sheets,
            out Excel.ListObject? resolvedTable,
            out Excel.Worksheet? resolvedWorksheet)
        {
            resolvedTable = table;
            resolvedWorksheet = table is null ? null : worksheet;
            return table is not null;
        }

        internal override Excel.ListColumns GetTableColumns(Excel.ListObject target)
        {
            Mock<Excel.ListColumns> columns = new();
            columns.Setup(c => c.Count).Returns(GanttTableSchema.Default.Columns.Count);
            return columns.Object;
        }

        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index)
        {
            GanttTableColumn definition = GanttTableSchema.Default.Columns[index - 1];

            // Both flags start already-correct, so only a genuinely wrong value
            // becomes a write. That makes the write-count assertion below a statement
            // about the comparison rather than about the loop.
            Mock<Excel.Range> entire = new();
            entire.SetupGet(c => c.Hidden).Returns(definition.IsHidden);
            entire.SetupSet(c => c.Hidden = It.IsAny<object>())
                .Callback((object _) => Recorded.Add("Hidden:" + index));

            Mock<Excel.Range> range = new();
            range.SetupGet(r => r.Locked).Returns(definition.IsLocked);
            range.SetupSet(r => r.Locked = It.IsAny<object>())
                .Callback((object _) => Recorded.Add("Locked:" + index));
            range.SetupGet(r => r.EntireColumn).Returns(entire.Object);

            Mock<Excel.ListColumn> column = new();
            column.SetupGet(c => c.Range).Returns(range.Object);
            return column.Object;
        }

        internal override Excel.Range GetEntireColumn(Excel.Range range) => range.EntireColumn;
    }

    private static Excel.Worksheet Worksheet()
    {
        Mock<Excel.Worksheet> worksheet = new();
        return worksheet.Object;
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

    /// <summary>
    /// A sheet already carrying the schema's classification writes nothing, so a
    /// Refresh does not mark the workbook dirty.
    /// </summary>
    /// <remarks>
    /// This is the positive test for the comparison the adapter exists to get right.
    /// A mutation that replaced the unboxed <c>ReadHiddenFlag</c> call with a direct
    /// <c>!= wanted</c> against the boxed property would make every column appear
    /// drifted, and this count would become the schema's column count rather than 0.
    /// </remarks>
    [Fact]
    public void An_already_correct_sheet_writes_nothing()
    {
        List<string> recorded = [];
        TestableRestorer restorer = new(
            Application().Object,
            Guard(ProtectionGuardOutcome.NotProtected).Object,
            new Mock<Excel.ListObject>().Object,
            Worksheet(),
            recorded);

        ColumnPresentationOutcome outcome = restorer.EnsureClassification();

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.ColumnsRestored);
        Assert.Empty(recorded);
    }

    /// <summary>
    /// A column whose flag has drifted is rewritten, so the restore is not vacuous.
    /// </summary>
    /// <remarks>
    /// The counterpart to the idempotence test. Without it, an adapter that wrote
    /// nothing at all would satisfy "writes nothing" and the guarantee would be
    /// untested — the same vacuity this review found twice already.
    /// </remarks>
    [Fact]
    public void A_drifted_hidden_flag_is_restored()
    {
        List<string> recorded = [];
        int firstEngineColumn = IndexOfFirstEngineColumn();

        // Every column reports Hidden = false, so each engine column is drifted.
        TestableRestorerWithDrift restorer = new(
            Application().Object,
            Guard(ProtectionGuardOutcome.NotProtected).Object,
            new Mock<Excel.ListObject>().Object,
            Worksheet(),
            recorded,
            firstEngineColumn);

        ColumnPresentationOutcome outcome = restorer.EnsureClassification();

        Assert.True(outcome.Succeeded);

        // The drifted column is rewritten, and written to the schema's value rather
        // than to whatever it already held.
        Assert.Contains(recorded, entry => entry.StartsWith("Hidden:" + firstEngineColumn + "=", StringComparison.Ordinal));
        Assert.Contains("Hidden:" + firstEngineColumn + "=True", recorded);
    }

    /// <summary>The one-based index of the first engine-hidden column in the schema.</summary>
    private static int IndexOfFirstEngineColumn()
    {
        IReadOnlyList<GanttTableColumn> columns = GanttTableSchema.Default.Columns;
        for (var index = 0; index < columns.Count; index++)
        {
            if (columns[index].IsHidden)
            {
                return index + 1;
            }
        }

        throw new InvalidOperationException("The schema declares no engine-hidden column.");
    }

    /// <summary>
    /// A restorer whose columns all report <c>Hidden = false</c>, so every
    /// engine-hidden column is genuinely drifted and must be rewritten.
    /// </summary>
    /// <param name="application">The application object.</param>
    /// <param name="guard">The protection guard.</param>
    /// <param name="table">The table.</param>
    /// <param name="worksheet">The resolved worksheet.</param>
    /// <param name="recorded">The flag writes recorded.</param>
    /// <param name="driftedColumn">The one-based index reporting a wrong flag.</param>
    /// <remarks>
    /// The write log is taken from the base class's <c>Recorded</c> property rather
    /// than captured here: capturing a primary-constructor parameter that is also
    /// passed to the base constructor is the double-capture CS9107 warns about, and
    /// passing a fresh list would have recorded into a list the test never reads.
    /// </remarks>
    private sealed class TestableRestorerWithDrift(
        object? application,
        IWorksheetProtectionGuard guard,
        Excel.ListObject table,
        Excel.Worksheet worksheet,
        List<string> recorded,
        int driftedColumn)
        : TestableRestorer(application, guard, table, worksheet, recorded)
    {
        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index)
        {
            if (index != driftedColumn)
            {
                return base.GetColumnAt(columns, index);
            }

            Mock<Excel.Range> entire = new();
            entire.SetupGet(c => c.Hidden).Returns(false);
            entire.SetupSet(c => c.Hidden = It.IsAny<object>())
                .Callback((object value) => Recorded.Add("Hidden:" + index + "=" + value));

            Mock<Excel.Range> range = new();
            range.SetupGet(r => r.Locked)
                .Returns(GanttTableSchema.Default.Columns[index - 1].IsLocked);
            range.SetupGet(r => r.EntireColumn).Returns(entire.Object);

            Mock<Excel.ListColumn> column = new();
            column.SetupGet(c => c.Range).Returns(range.Object);
            return column.Object;
        }
    }

    /// <summary>A protected target refuses and writes nothing (ADR-0008 D4).</summary>
    [Fact]
    public void A_protected_target_refuses_without_writing()
    {
        List<string> recorded = [];
        TestableRestorer restorer = new(
            Application().Object,
            Guard(ProtectionGuardOutcome.SheetProtected).Object,
            new Mock<Excel.ListObject>().Object,
            Worksheet(),
            recorded);

        ColumnPresentationOutcome outcome = restorer.EnsureClassification();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ColumnPresentationRefusalReason.TargetProtected, outcome.Refusal);
        Assert.Empty(recorded);
    }

    /// <summary>A missing table refuses with its own typed reason.</summary>
    [Fact]
    public void A_missing_table_refuses_without_writing()
    {
        List<string> recorded = [];
        TestableRestorer restorer = new(
            Application().Object,
            Guard(ProtectionGuardOutcome.NotProtected).Object,
            table: null,
            Worksheet(),
            recorded);

        ColumnPresentationOutcome outcome = restorer.EnsureClassification();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ColumnPresentationRefusalReason.TableMissing, outcome.Refusal);
        Assert.Empty(recorded);
    }

    /// <summary>A missing workbook refuses with the no-workbook reason.</summary>
    [Fact]
    public void A_missing_workbook_refuses_without_writing()
    {
        Mock<Excel.Application> application = new();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns((Excel.Workbook)null!);
        List<string> recorded = [];
        TestableRestorer restorer = new(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected).Object,
            new Mock<Excel.ListObject>().Object,
            Worksheet(),
            recorded);

        ColumnPresentationOutcome outcome = restorer.EnsureClassification();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ColumnPresentationRefusalReason.NoActiveWorkbook, outcome.Refusal);
        Assert.Empty(recorded);
    }

    /// <summary>A foreign application object degrades rather than throwing.</summary>
    [Fact]
    public void A_foreign_application_object_degrades_to_no_workbook()
    {
        ExcelColumnPresentationRestorer restorer = new(new object());

        ColumnPresentationOutcome outcome = restorer.EnsureClassification();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ColumnPresentationRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }
}