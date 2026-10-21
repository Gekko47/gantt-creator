using GanttCreator.Core;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Tests for the R5.1 mode-change warnings (WI-G): an invalid plot date and
/// an inverted plot range both produce a user-facing, actionable message
/// rather than a silent revert.
/// </summary>
[Collection("diagnostics-service")]
public sealed class PlotRangeWarningTests : IDisposable
{
    public PlotRangeWarningTests() => RibbonStateService.Reset();

    public void Dispose()
    {
        RibbonStateService.Reset();
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Builds a state service whose presenter collects the messages instead
    /// of showing a dialog.
    /// </summary>
    /// <param name="messages">The collected messages.</param>
    /// <returns>The service under test.</returns>
    private static RibbonStateService ServiceWith(List<string> messages)
    {
        RibbonStateService service = RibbonStateService.Instance;
        service.SetErrorPresenter(message => messages.Add(message));
        return service;
    }

    [Fact]
    public void An_unparsable_start_date_warns_instead_of_silently_reverting()
    {
        // R5.1 D3: a mode-change warning follows the R2.6 message discipline
        // -- it says what was wrong and what to do. A silent revert leaves the
        // user pressing OK on a box that never took their value, with nothing
        // to say why.
        var messages = new List<string>();
        RibbonStateService service = ServiceWith(messages);

        service.SetPlotStartDate("not-a-date");

        var message = Assert.Single(messages);
        Assert.Contains("not-a-date", message, StringComparison.Ordinal);
        Assert.Contains("dd-MMM-yy", message, StringComparison.Ordinal);
    }

    [Fact]
    public void An_unparsable_finish_date_warns_instead_of_silently_reverting()
    {
        var messages = new List<string>();
        RibbonStateService service = ServiceWith(messages);

        service.SetPlotFinishDate("not-a-date");

        var message = Assert.Single(messages);
        Assert.Contains("not-a-date", message, StringComparison.Ordinal);
        Assert.Contains("dd-MMM-yy", message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_valid_date_does_not_warn()
    {
        // The control for the two cases above: without it, "no warning" would
        // be indistinguishable from "the warning never fires".
        var messages = new List<string>();
        RibbonStateService service = ServiceWith(messages);

        service.SetPlotStartDate("15-Mar-26");

        Assert.Empty(messages);
    }

    [Fact]
    public void A_start_after_the_finish_warns()
    {
        // The second warning R5.1 D3 asks for: the range cannot be drawn, so
        // the start date is refused and the warning says so.
        var messages = new List<string>();
        RibbonStateService service = ServiceWith(messages);

        service.SetPlotFinishDate("01-Jan-26");
        service.SetPlotStartDate("15-Mar-26");

        var message = Assert.Single(messages);
        Assert.Contains("after", message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void An_inverted_range_keeps_the_previous_start_date()
    {
        // The refusal happens BEFORE the state is committed, so the workbook
        // and the in-memory state agree and the next Refresh does not refuse
        // the range the user was told about.
        var messages = new List<string>();
        RibbonStateService service = ServiceWith(messages);

        service.SetPlotFinishDate("01-Jan-26");
        service.SetPlotStartDate("15-Mar-26");

        // The refused start was never committed, so the start end still holds
        // the value it had before the attempt.
        Assert.NotEqual("15-Mar-26", service.GetPlotStartDate());
    }

    [Fact]
    public void A_finish_before_the_start_warns()
    {
        // The symmetric case: the finish end is the one being edited.
        var messages = new List<string>();
        RibbonStateService service = ServiceWith(messages);

        service.SetPlotStartDate("15-Mar-26");
        service.SetPlotFinishDate("01-Jan-26");

        var message = Assert.Single(messages);
        Assert.Contains("before", message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void A_range_in_the_right_order_does_not_warn()
    {
        // The control for the inverted-range warning: a start before the
        // finish is the ordinary case and must stay silent.
        var messages = new List<string>();
        RibbonStateService service = ServiceWith(messages);

        service.SetPlotStartDate("01-Jan-26");
        service.SetPlotFinishDate("15-Mar-26");

        Assert.Empty(messages);
    }
}
