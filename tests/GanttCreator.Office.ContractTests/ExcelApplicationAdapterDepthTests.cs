using Excel = Microsoft.Office.Interop.Excel;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Depth tests for <see cref="ExcelApplicationAdapter"/>: the status-bar writes and
/// the <c>SheetChange</c> handler's guards. The existing tests attach and detach the
/// handler but never raise it.
/// </summary>
/// <remarks>
/// The event is reachable because the PIA exposes it as an interface member: a Moq
/// <c>SetupAdd</c> captures the delegate the adapter attached, and the test invokes
/// it the way Excel would. One branch stays out of reach and is recorded rather than
/// faked: <c>Range.Address</c> is a COM indexed property with optional arguments, so
/// it cannot appear in a Moq expression tree — the happy-path address composition is
/// exercised by the tagged live-Office test.
/// </remarks>
public sealed class ExcelApplicationAdapterDepthTests
{
    /// <summary>
    /// An adapter over a mock application whose <c>SheetChange</c> delegate is
    /// captured so a test can raise it.
    /// </summary>
    private sealed class Harness
    {
        public Harness()
        {
            Application = new Mock<Excel.Application>();
            Events = Application.As<Excel.AppEvents_Event>();
            _ = Events.SetupAdd(e => e.SheetChange += It.IsAny<Excel.AppEvents_SheetChangeEventHandler>())
                .Callback<Excel.AppEvents_SheetChangeEventHandler>(handler => SheetChangeHandler = handler);
        }

        public Mock<Excel.Application> Application { get; }

        public Mock<Excel.AppEvents_Event> Events { get; }

        public Excel.AppEvents_SheetChangeEventHandler? SheetChangeHandler { get; private set; }

        public ExcelApplicationAdapter Adapter() => new(Application.Object);

        /// <summary>Raises the captured handler the way Excel would.</summary>
        public void RaiseSheetChange(object sheet, Excel.Range target) => SheetChangeHandler!(sheet, target);
    }

    private static Excel.Worksheet SheetWhoseParentThrows()
    {
        Mock<Excel.Worksheet> sheet = new();
        _ = sheet.SetupGet(s => s.Parent).Throws(new InvalidOperationException("the host refused"));
        return sheet.Object;
    }

    [Fact]
    public void ShowStatusBar_rejects_a_null_message()
    {
        // Positive test for the argument guard on the adapter's port.
        ExcelApplicationAdapter adapter = new(new Mock<Excel.Application>().Object);

        Assert.Throws<ArgumentNullException>(() => adapter.ShowStatusBar(null!));
    }

    [Fact]
    public void ShowStatusBar_does_nothing_without_an_application_object()
    {
        ExcelApplicationAdapter adapter = new(null);

        Assert.Null(Record.Exception(() => adapter.ShowStatusBar("hello")));
    }

    [Fact]
    public void ShowStatusBar_writes_the_message_onto_the_host()
    {
        Mock<Excel.Application> application = new();
        object? written = null;
        _ = application.SetupSet(a => a.StatusBar = It.IsAny<object>())
            .Callback<object>(value => written = value);

        new ExcelApplicationAdapter(application.Object).ShowStatusBar("hello");

        Assert.Equal("hello", written);
    }

    [Fact]
    public void ShowStatusBar_swallows_a_host_that_refuses_the_write()
    {
        // A failed status-bar write loses the message and nothing more; it must not
        // escape into the Ribbon state path.
        Mock<Excel.Application> application = new();
        _ = application.SetupSet(a => a.StatusBar = It.IsAny<object>())
            .Throws(new InvalidOperationException("the host refused"));

        Assert.Null(Record.Exception(() => new ExcelApplicationAdapter(application.Object).ShowStatusBar("hello")));
    }

    [Fact]
    public void ClearStatusBar_does_nothing_without_an_application_object()
    {
        ExcelApplicationAdapter adapter = new(null);

        Assert.Null(Record.Exception(adapter.ClearStatusBar));
    }

    [Fact]
    public void ClearStatusBar_restores_the_host_automatic_text()
    {
        // False, not empty string: false is the host's documented "automatic".
        Mock<Excel.Application> application = new();
        object? written = null;
        _ = application.SetupSet(a => a.StatusBar = It.IsAny<object>())
            .Callback<object>(value => written = value);

        new ExcelApplicationAdapter(application.Object).ClearStatusBar();

        Assert.Equal(false, written);
    }

    [Fact]
    public void ClearStatusBar_swallows_a_host_that_refuses_the_write()
    {
        Mock<Excel.Application> application = new();
        _ = application.SetupSet(a => a.StatusBar = It.IsAny<object>())
            .Throws(new InvalidOperationException("the host refused"));

        Assert.Null(Record.Exception(() => new ExcelApplicationAdapter(application.Object).ClearStatusBar()));
    }

    [Fact]
    public void A_sheet_change_after_dispose_delivers_nothing()
    {
        Harness harness = new();
        List<string?> reported = [];
        IDisposable subscription = harness.Adapter().SubscribeSheetChanged(reported.Add);
        subscription.Dispose();

        harness.RaiseSheetChange(new Mock<Excel.Worksheet>().Object, new Mock<Excel.Range>().Object);

        Assert.Empty(reported);
    }

    [Fact]
    public void A_sheet_change_for_a_chart_sheet_delivers_nothing()
    {
        // A chart sheet is not an Excel.Worksheet, so it has no name to qualify and
        // must be skipped rather than reported as a change.
        Harness harness = new();
        List<string?> reported = [];
        using IDisposable subscription = harness.Adapter().SubscribeSheetChanged(reported.Add);

        harness.RaiseSheetChange(new object(), new Mock<Excel.Range>().Object);

        Assert.Empty(reported);
    }

    [Fact]
    public void A_sheet_change_whose_parent_probe_fails_delivers_nothing()
    {
        // The parent read is a COM call inside an event dispatch: it must not escape,
        // and it must not report a change whose target is unknown.
        Harness harness = new();
        List<string?> reported = [];
        using IDisposable subscription = harness.Adapter().SubscribeSheetChanged(reported.Add);

        Assert.Null(Record.Exception(
            () => harness.RaiseSheetChange(SheetWhoseParentThrows(), new Mock<Excel.Range>().Object)));
        Assert.Empty(reported);
    }

    [Fact]
    public void A_second_dispose_of_a_sheet_change_subscription_is_a_no_op()
    {
        Harness harness = new();
        IDisposable subscription = harness.Adapter().SubscribeSheetChanged(_ => { });

        subscription.Dispose();

        Assert.Null(Record.Exception(subscription.Dispose));
    }
}
