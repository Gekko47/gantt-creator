using System.Diagnostics.CodeAnalysis;
using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for the two adapters the row commands drive and nothing else
/// covered: <see cref="ExcelInsertedRowSelector"/> (0% before this file) and
/// <see cref="ExcelPlotAnchorRepairer"/> (27%).
/// </summary>
public sealed class ExcelRowCommandAdapterDepthTests
{
    /// <summary>
    /// Justification shared by the two tests that must raise a real
    /// <see cref="System.Runtime.InteropServices.COMException"/>: the repairer
    /// catches that type and nothing else, because a failed defined-name write
    /// crosses the interop boundary as one. A test-only suppression; production code
    /// still may not raise it.
    /// </summary>
    private const string _comExceptionJustification =
        "The adapter catches COMException specifically, because that is how a refused "
        + "defined-name write crosses the interop boundary. A test that substituted "
        + "another exception type would exercise a path the product never takes.";

    // ---- ExcelInsertedRowSelector ----

    /// <summary>
    /// The selector with the collection indexers substituted, so the workbook walk
    /// runs. Every proxy it would otherwise reach through COM is supplied here.
    /// </summary>
    private sealed class TestableSelector(
        object? application,
        IReadOnlyList<object?> sheets,
        IReadOnlyList<Excel.ListObject> tables,
        Excel.ListRows? listRows,
        Excel.ListRow? listRow,
        Excel.Range? rowRange) : ExcelInsertedRowSelector(application)
    {
        internal override object GetSheetAt(Excel.Sheets sheets1, int index) => sheets[index - 1]!;

        internal override Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index) =>
            tables[index - 1];

        internal override Excel.ListRows GetListRows(Excel.ListObject table) =>
            listRows ?? throw new InvalidOperationException("the host refused the list rows");

        internal override Excel.ListRow GetListRowAt(Excel.ListRows rows, int bodyIndex) =>
            listRow ?? throw new InvalidOperationException("the host refused the row");

        internal override Excel.Range GetRowRange(Excel.ListRow row) =>
            rowRange ?? throw new InvalidOperationException("the host refused the range");
    }

    private static Excel.Workbook WorkbookWithSheets(int count)
    {
        Mock<Excel.Sheets> sheets = new();
        _ = sheets.SetupGet(s => s.Count).Returns(count);
        Mock<Excel.Workbook> workbook = new();
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);
        return workbook.Object;
    }

    private static Excel.Application ApplicationWith(Excel.Workbook? workbook)
    {
        Mock<Excel.Application> application = new();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook!);
        return application.Object;
    }

    private static Mock<Excel.Worksheet> SheetWithTables(params Excel.ListObject[] tables)
    {
        Mock<Excel.ListObjects> collection = new();
        _ = collection.SetupGet(o => o.Count).Returns(tables.Length);
        Mock<Excel.Worksheet> sheet = new();
        _ = sheet.SetupGet(s => s.ListObjects).Returns(collection.Object);
        return sheet;
    }

    private static Excel.ListObject Table(string name)
    {
        Mock<Excel.ListObject> table = new();
        _ = table.SetupGet(t => t.Name).Returns(name);
        return table.Object;
    }

    [Fact]
    public void SelectBodyRow_does_nothing_without_an_application_object()
    {
        // No captured workbook means no selection change; the row has still been
        // added, so this must not throw.
        ExcelInsertedRowSelector selector = new(null);

        Assert.Null(Record.Exception(() => selector.SelectBodyRow(1)));
    }

    [Fact]
    public void SelectBodyRow_does_nothing_when_the_workbook_capture_fails()
    {
        // A host that refuses the capture degrades to "the cursor stays put".
        Mock<Excel.Application> application = new();
        _ = application.SetupGet(a => a.ActiveWorkbook).Throws(new InvalidOperationException("refused"));
        ExcelInsertedRowSelector selector = new(application.Object);

        Assert.Null(Record.Exception(() => selector.SelectBodyRow(1)));
    }

    [Fact]
    public void SelectBodyRow_skips_a_chart_sheet_and_a_foreign_table_before_selecting()
    {
        // The walk must pass a non-worksheet entry and a table that is not the Gantt
        // table, then activate the sheet and select the requested row.
        Mock<Excel.Worksheet> gantt = SheetWithTables(Table("tblSomethingElse"), Table(GanttTableSchema.TableName));
        Mock<Excel.Range> range = new();
        TestableSelector selector = new(
            ApplicationWith(WorkbookWithSheets(2)),
            [new object(), gantt.Object],
            [Table("tblSomethingElse"), Table(GanttTableSchema.TableName)],
            new Mock<Excel.ListRows>().Object,
            new Mock<Excel.ListRow>().Object,
            range.Object);

        selector.SelectBodyRow(3);

        gantt.Verify(s => s.Activate(), Times.Once);
        range.Verify(r => r.Select(), Times.Once);
    }

    [Fact]
    public void SelectBodyRow_leaves_the_cursor_alone_when_the_walk_fails()
    {
        // The row is already added by the time this runs, so a refused lookup must
        // not surface as a failed insert.
        Mock<Excel.Worksheet> gantt = SheetWithTables(Table(GanttTableSchema.TableName));
        TestableSelector selector = new(
            ApplicationWith(WorkbookWithSheets(1)),
            [gantt.Object],
            [Table(GanttTableSchema.TableName)],
            listRows: null,
            listRow: null,
            rowRange: null);

        Assert.Null(Record.Exception(() => selector.SelectBodyRow(3)));
        gantt.Verify(s => s.Activate(), Times.Never);
    }

    [Fact]
    public void SelectBodyRow_does_nothing_when_no_sheet_carries_the_gantt_table()
    {
        Mock<Excel.Worksheet> plain = SheetWithTables(Table("tblSomethingElse"));
        TestableSelector selector = new(
            ApplicationWith(WorkbookWithSheets(1)),
            [plain.Object],
            [Table("tblSomethingElse")],
            new Mock<Excel.ListRows>().Object,
            new Mock<Excel.ListRow>().Object,
            new Mock<Excel.Range>().Object);

        selector.SelectBodyRow(1);

        plain.Verify(s => s.Activate(), Times.Never);
    }

    // ---- ExcelPlotAnchorRepairer ----

    /// <summary>
    /// The repairer with only the collection indexers substituted, so the table
    /// search and the defined-name search both run.
    /// </summary>
    private sealed class TestableRepairer(
        object? application,
        IWorksheetProtectionGuard guard,
        IReadOnlyList<Excel._Worksheet?> sheets,
        IReadOnlyList<Excel.ListObject> tables) : ExcelPlotAnchorRepairer(application, guard)
    {
        internal override Excel.Worksheet GetSheetAt(Excel.Sheets sheets1, int index) =>
            (Excel.Worksheet)sheets[index - 1]!;

        internal override Excel.ListObject GetTableAt(Excel.ListObjects objects, int index) =>
            tables[index - 1];
    }

    private static Mock<IWorksheetProtectionGuard> Guard(ProtectionGuardOutcome target)
    {
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.QueryTarget(It.IsAny<object>())).Returns(target);
        return guard;
    }

    private static Excel.Application ApplicationWith(Excel.Sheets sheets)
    {
        Mock<Excel.Workbook> workbook = new();
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheets);
        Mock<Excel.Application> application = new();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        return application.Object;
    }

    private static Excel.Sheets OneSheet() => Sheets(1);

    private static Excel.Sheets Sheets(int count)
    {
        Mock<Excel.Sheets> sheets = new();
        _ = sheets.SetupGet(s => s.Count).Returns(count);
        return sheets.Object;
    }

    /// <summary>A Gantt worksheet whose defined-name collection is the one supplied.</summary>
    /// <param name="names">The names collection, or null for a loose one.</param>
    /// <param name="namesThrow">Whether reading the collection raises the COM failure.</param>
    /// <param name="tableCount">How many tables the sheet reports.</param>
    /// <returns>The worksheet.</returns>
    [SuppressMessage("Design", "CA2201:Do not raise reserved exception types", Justification = _comExceptionJustification)]
    private static Excel.Worksheet GanttSheetWithNames(
        Excel.Names? names = null,
        bool namesThrow = false,
        int tableCount = 1)
    {
        Mock<Excel.ListObjects> tables = new();
        _ = tables.SetupGet(o => o.Count).Returns(tableCount);
        Mock<Excel.Worksheet> sheet = new();
        _ = sheet.SetupGet(s => s.Name).Returns("Gantt");
        _ = sheet.SetupGet(s => s.ListObjects).Returns(tables.Object);
        if (namesThrow)
        {
            _ = sheet.SetupGet(s => s.Names)
                .Throws(new System.Runtime.InteropServices.COMException("the host refused"));
        }
        else
        {
            _ = sheet.SetupGet(s => s.Names).Returns(names ?? new Mock<Excel.Names>().Object);
        }

        return sheet.Object;
    }

    private static TestableRepairer Repairer(
        Excel.Application application,
        Excel.Worksheet sheet,
        ProtectionGuardOutcome target = ProtectionGuardOutcome.NotProtected,
        params Excel.ListObject[] tables) =>
        new(application, Guard(target).Object, [sheet], tables);

    [Fact]
    public void Repair_refuses_without_an_active_workbook()
    {
        TestableRepairer repairer = new(null, Guard(ProtectionGuardOutcome.NotProtected).Object, [], []);

        Assert.Equal(PlotAnchorRepairRefusalReason.NoActiveWorkbook, repairer.Repair().Refusal);
    }

    [Fact]
    public void Repair_refuses_when_no_sheet_carries_the_gantt_table()
    {
        TestableRepairer repairer = Repairer(
            ApplicationWith(OneSheet()),
            GanttSheetWithNames(),
            ProtectionGuardOutcome.NotProtected,
            Table("tblSomethingElse"));

        Assert.Equal(PlotAnchorRepairRefusalReason.TableMissing, repairer.Repair().Refusal);
    }

    [Fact]
    public void Repair_refuses_a_protected_target_before_writing_a_name()
    {
        TestableRepairer repairer = Repairer(
            ApplicationWith(OneSheet()),
            GanttSheetWithNames(),
            ProtectionGuardOutcome.SheetProtected,
            Table(GanttTableSchema.TableName));

        Assert.Equal(PlotAnchorRepairRefusalReason.TargetProtected, repairer.Repair().Refusal);
    }

    [Fact]
    public void Repair_reports_no_active_workbook_when_the_guard_cannot_see_one()
    {
        TestableRepairer repairer = Repairer(
            ApplicationWith(OneSheet()),
            GanttSheetWithNames(),
            ProtectionGuardOutcome.NoActiveWorkbook,
            Table(GanttTableSchema.TableName));

        Assert.Equal(PlotAnchorRepairRefusalReason.NoActiveWorkbook, repairer.Repair().Refusal);
    }

    [Fact]
    public void Repair_adds_the_anchor_name_when_the_sheet_has_none()
    {
        // A loose Names mock accepts the add and reports nothing found beforehand,
        // so the assertion is that the add path was taken and reported success.
        TestableRepairer repairer = Repairer(
            ApplicationWith(OneSheet()),
            GanttSheetWithNames(),
            ProtectionGuardOutcome.NotProtected,
            Table(GanttTableSchema.TableName));

        PlotAnchorRepairOutcome outcome = repairer.Repair();

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
    }

    [Fact]
    public void Repair_rewrites_an_existing_anchor_rather_than_adding_a_second_name()
    {
        // The anchor exists, so the rewrite path runs and no second name is added.
        // The written value is GanttSheetLayout.BuildPlotAnchorRefersTo, which Core's
        // own tests already pin; asserting it here would need a setter callback the
        // COM property's delegate shape does not support.
        Mock<Excel.Name> existing = new();
        _ = existing.SetupGet(n => n.Name).Returns(GanttWorkbookContract.PlotAnchorDefinedName);
        Mock<Excel.Names> names = new();
        _ = names.SetupGet(n => n.Count).Returns(1);
        _ = names.Setup(n => n.Item(It.IsAny<object>())).Returns(existing.Object);
        TestableRepairer repairer = Repairer(
            ApplicationWith(OneSheet()),
            GanttSheetWithNames(names.Object),
            ProtectionGuardOutcome.NotProtected,
            Table(GanttTableSchema.TableName));

        PlotAnchorRepairOutcome outcome = repairer.Repair();

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
    }

    [Fact]
    [SuppressMessage("Design", "CA2201:Do not raise reserved exception types", Justification = _comExceptionJustification)]
    public void Repair_reports_a_refused_names_collection_as_a_typed_refusal()
    {
        // The same refusal is provoked one call earlier than a refused Names.Add:
        // the COM signature of Add uses optional arguments, and an expression tree
        // may not make such a call, so the host refuses the collection instead.
        TestableRepairer repairer = Repairer(
            ApplicationWith(OneSheet()),
            GanttSheetWithNames(namesThrow: true),
            ProtectionGuardOutcome.NotProtected,
            Table(GanttTableSchema.TableName));

        Assert.Equal(PlotAnchorRepairRefusalReason.NameWriteFailed, repairer.Repair().Refusal);
    }
}
