using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Depth tests for <see cref="ExcelSelectionIdentifier"/>: the "no selection"
/// arms of <see cref="ExcelSelectionIdentifier.Identify"/> and the
/// <c>SheetSelectionChange</c> subscription.
/// </summary>
/// <remarks>
/// <see cref="ExcelSelectionIdentifierTests"/> covers the resolved facts — one
/// body row, a header, a multi-row span. This file covers what happens when the
/// probe cannot answer at all, including the two catches that exist so a COM
/// failure reports "no selection" rather than escaping the Ribbon state path, and
/// the event connection that keeps the dropdowns fresh.
/// </remarks>
public sealed class ExcelSelectionIdentifierDepthTests
{
    /// <summary>
    /// The identifier with every probe substituted, so a test controls exactly one
    /// answer and the rest stay static.
    /// </summary>
    private sealed class DepthIdentifier(Excel.Application? application) : ExcelSelectionIdentifier(application)
    {
        public int SheetCount { get; set; } = 1;

        public bool SheetCountThrows { get; set; }

        public Excel.Range? HeaderRow { get; set; }

        public Excel.Range? DataBody { get; set; }

        public Excel.Range? Active { get; set; }

        public bool SelectionThrows { get; set; }

        public Excel.Range? HeaderIntersection { get; set; }

        public Excel.Range? BodyIntersection { get; set; }

        public int BodyFirstRow { get; set; } = 2;

        public int ActiveRow { get; set; }

        public int RowCount { get; set; } = 1;

        private List<Excel.ListObject> Tables { get; } = [];

        public DepthIdentifier WithTable()
        {
            Mock<Excel.ListObject> table = new();
            _ = table.SetupGet(t => t.Name).Returns(GanttTableSchema.TableName);
            Tables.Add(table.Object);
            return this;
        }

        internal override int GetSheetCount(Excel.Sheets sheets) =>
            SheetCountThrows ? throw new InvalidOperationException("the host refused") : SheetCount;

        internal override object GetSheetAt(Excel.Sheets sheets, int index)
        {
            Mock<Excel.ListObjects> listObjects = new();
            _ = listObjects.SetupGet(o => o.Count).Returns(Tables.Count);
            Mock<Excel.Worksheet> sheet = new();
            _ = sheet.SetupGet(s => s.ListObjects).Returns(listObjects.Object);
            return sheet.Object;
        }

        internal override Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index) =>
            Tables[index - 1];

        internal override Excel.Range? GetHeaderRow(Excel.ListObject table) => HeaderRow;

        internal override Excel.Range? GetDataBodyRange(Excel.ListObject table) => DataBody;

        internal override Excel.Range? GetActiveCell(Excel.Application application) => Active;

        internal override object? GetSelection(Excel.Application application) =>
            SelectionThrows ? throw new InvalidOperationException("the host refused") : null;

        internal override int GetRangeRow(Excel.Range range) =>
            ReferenceEquals(range, DataBody) ? BodyFirstRow : ActiveRow;

        internal override int GetRowCount(Excel.Range range) => RowCount;

        internal override Excel.Range? Intersect(Excel.Range first, Excel.Range second) =>
            ReferenceEquals(first, HeaderRow) ? HeaderIntersection : BodyIntersection;
    }

    private static Excel.Range Range() => new Mock<Excel.Range>().Object;

    /// <summary>An identifier over a workbook, so <c>Identify</c> gets past the null check.</summary>
    private static DepthIdentifier Identifier()
    {
        Mock<Excel.Workbook> workbook = new();
        _ = workbook.SetupGet(w => w.Sheets).Returns(new Mock<Excel.Sheets>().Object);
        Mock<Excel.Application> application = new();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        return new DepthIdentifier(application.Object);
    }

    [Fact]
    public void Identify_reports_none_when_no_workbook_is_active()
    {
        Mock<Excel.Application> application = new();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns((Excel.Workbook)null!);

        Assert.Equal(SelectionFacts.None, new DepthIdentifier(application.Object).Identify());
    }

    [Fact]
    public void Identify_reports_none_without_an_application_object()
    {
        Assert.Equal(SelectionFacts.None, new DepthIdentifier(null).Identify());
    }

    [Fact]
    public void Identify_reports_none_when_no_sheet_carries_the_gantt_table()
    {
        DepthIdentifier identifier = Identifier();
        identifier.SheetCount = 0;

        Assert.Equal(SelectionFacts.None, identifier.Identify());
    }

    [Fact]
    public void Identify_reports_none_when_the_table_search_throws()
    {
        // A COM failure while looking for the table is caught inside the search, so
        // the caller sees "no selection" rather than an exception on the Ribbon path.
        DepthIdentifier identifier = Identifier();
        identifier.SheetCountThrows = true;

        Assert.Equal(SelectionFacts.None, identifier.Identify());
    }

    [Fact]
    public void Identify_reports_none_when_the_selection_probe_throws()
    {
        DepthIdentifier identifier = Identifier().WithTable();
        identifier.HeaderRow = Range();
        identifier.SelectionThrows = true;

        Assert.Equal(SelectionFacts.None, identifier.Identify());
    }

    [Fact]
    public void Identify_reports_none_when_the_table_has_no_header_row()
    {
        DepthIdentifier identifier = Identifier().WithTable();
        identifier.DataBody = Range();

        Assert.Equal(SelectionFacts.None, identifier.Identify());
    }

    [Fact]
    public void Identify_reports_none_when_the_selection_is_outside_the_header_and_body()
    {
        DepthIdentifier identifier = Identifier().WithTable();
        identifier.HeaderRow = Range();
        identifier.Active = Range();

        // Neither intersection resolves and no body range exists at all.
        Assert.Equal(SelectionFacts.None, identifier.Identify());
    }

    [Fact]
    public void Identify_reports_none_when_the_body_starts_at_an_unresolvable_row()
    {
        // A body whose first worksheet row reads as zero cannot be turned into a
        // one-based body index, so the fact is withheld rather than guessed.
        Excel.Range active = Range();
        Excel.Range body = Range();
        DepthIdentifier identifier = Identifier().WithTable();
        identifier.HeaderRow = Range();
        identifier.DataBody = body;
        identifier.Active = active;
        identifier.BodyFirstRow = 0;
        identifier.BodyIntersection = Range();

        Assert.Equal(SelectionFacts.None, identifier.Identify());
    }

    [Fact]
    public void Identify_resolves_a_single_body_row_as_the_one_expanded_fact()
    {
        // The positive control: the same harness must produce a resolved fact, or the
        // refusals above would pass vacuously.
        Excel.Range active = Range();
        Excel.Range body = Range();
        Excel.Range intersection = Range();
        DepthIdentifier identifier = Identifier().WithTable();
        identifier.HeaderRow = Range();
        identifier.DataBody = body;
        identifier.Active = active;
        identifier.BodyFirstRow = 4;
        identifier.ActiveRow = 6;
        identifier.RowCount = 1;
        identifier.BodyIntersection = intersection;

        SelectionFacts facts = identifier.Identify();

        Assert.True(facts.HasSelection);
        Assert.True(facts.IsSingleBodyRow);
        Assert.Equal(3, facts.BodyIndex);
        Assert.False(facts.IsHeaderRow);
        Assert.Equal(1, facts.SelectedRowCount);
    }

    [Fact]
    public void A_selection_change_subscription_attaches_and_detaches_one_handler()
    {
        Mock<Excel.Application> application = new();
        Mock<Excel.AppEvents_Event> events = application.As<Excel.AppEvents_Event>();

        using (new ExcelSelectionIdentifier(application.Object).SubscribeSelectionChanged(() => { }))
        {
            events.VerifyAdd(
                e => e.SheetSelectionChange += It.IsAny<Excel.AppEvents_SheetSelectionChangeEventHandler>(),
                Times.Once);
        }

        events.VerifyRemove(
            e => e.SheetSelectionChange -= It.IsAny<Excel.AppEvents_SheetSelectionChangeEventHandler>(),
            Times.Once);
    }

    [Fact]
    public void A_selection_change_invokes_the_handler()
    {
        Mock<Excel.Application> application = new();
        Excel.AppEvents_SheetSelectionChangeEventHandler? captured = null;
        _ = application.As<Excel.AppEvents_Event>()
            .SetupAdd(e => e.SheetSelectionChange += It.IsAny<Excel.AppEvents_SheetSelectionChangeEventHandler>())
            .Callback<Excel.AppEvents_SheetSelectionChangeEventHandler>(handler => captured = handler);
        int calls = 0;
        using IDisposable subscription = new ExcelSelectionIdentifier(application.Object)
            .SubscribeSelectionChanged(() => calls++);

        captured!(new object(), Range());

        Assert.Equal(1, calls);
    }

    [Fact]
    public void A_throwing_selection_handler_never_escapes_the_event_dispatch()
    {
        // A throw here would surface inside Excel's own event dispatch.
        Mock<Excel.Application> application = new();
        Excel.AppEvents_SheetSelectionChangeEventHandler? captured = null;
        _ = application.As<Excel.AppEvents_Event>()
            .SetupAdd(e => e.SheetSelectionChange += It.IsAny<Excel.AppEvents_SheetSelectionChangeEventHandler>())
            .Callback<Excel.AppEvents_SheetSelectionChangeEventHandler>(handler => captured = handler);
        using IDisposable subscription = new ExcelSelectionIdentifier(application.Object)
            .SubscribeSelectionChanged(() => throw new InvalidOperationException("handler failed"));

        Assert.Null(Record.Exception(() => captured!(new object(), Range())));
    }

    [Fact]
    public void A_second_dispose_of_a_selection_subscription_is_a_no_op()
    {
        Mock<Excel.Application> application = new();
        IDisposable subscription = new ExcelSelectionIdentifier(application.Object)
            .SubscribeSelectionChanged(() => { });

        subscription.Dispose();

        Assert.Null(Record.Exception(subscription.Dispose));
    }

    [Fact]
    public void SubscribeSelectionChanged_rejects_a_null_handler()
    {
        ExcelSelectionIdentifier identifier = new(new Mock<Excel.Application>().Object);

        Assert.Throws<ArgumentNullException>(() => identifier.SubscribeSelectionChanged(null!));
    }

    [Fact]
    public void Subscribing_without_an_application_object_returns_an_inert_subscription()
    {
        IDisposable subscription = new ExcelSelectionIdentifier(null).SubscribeSelectionChanged(() => { });

        Assert.Null(Record.Exception(subscription.Dispose));
        Assert.Null(Record.Exception(subscription.Dispose));
    }
}
