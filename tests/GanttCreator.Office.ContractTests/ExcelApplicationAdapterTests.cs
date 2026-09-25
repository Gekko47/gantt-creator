using Excel = Microsoft.Office.Interop.Excel;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ExcelApplicationAdapter"/>. They need no live
/// Office: the application object is either absent, foreign, or a Moq proxy of
/// the Excel application interface (the PIA exposes it as an interface, so Moq
/// can implement it and its <c>AppEvents_Event</c> event interface).
/// </summary>
public class ExcelApplicationAdapterTests
{
    [Fact]
    public void HasActiveWorkbook_reports_not_determinable_without_an_application_object()
        => Assert.Null(new ExcelApplicationAdapter(null).HasActiveWorkbook());

    [Fact]
    public void HasActiveWorkbook_reports_not_determinable_for_a_foreign_object()
        => Assert.Null(new ExcelApplicationAdapter("not an Excel application").HasActiveWorkbook());

    [Fact]
    public void HasActiveWorkbook_is_false_when_excel_has_no_active_workbook()
    {
        var application = new Mock<Excel.Application>();
        application.SetupGet(a => a.ActiveWorkbook).Returns((Excel.Workbook)null!);

        Assert.False(new ExcelApplicationAdapter(application.Object).HasActiveWorkbook());
    }

    [Fact]
    public void HasActiveWorkbook_is_true_when_a_workbook_is_active()
    {
        var application = new Mock<Excel.Application>();
        application.SetupGet(a => a.ActiveWorkbook).Returns(new Mock<Excel.Workbook>().Object);

        Assert.True(new ExcelApplicationAdapter(application.Object).HasActiveWorkbook());
    }

    [Fact]
    public void HasActiveWorkbook_reports_not_determinable_when_the_probe_fails()
    {
        var application = new Mock<Excel.Application>();
        application.SetupGet(a => a.ActiveWorkbook).Throws(new InvalidOperationException("probe failed"));

        Assert.Null(new ExcelApplicationAdapter(application.Object).HasActiveWorkbook());
    }

    [Fact]
    public void SubscribeWorkbookStateChanged_rejects_a_null_handler()
    {
        // Positive test for the one argument guard on the adapter's port.
        var adapter = new ExcelApplicationAdapter(new Mock<Excel.Application>().Object);

        Assert.Throws<ArgumentNullException>(() => adapter.SubscribeWorkbookStateChanged(null!));
    }

    [Fact]
    public void SubscribeWorkbookStateChanged_without_an_application_returns_an_inert_subscription()
    {
        var subscription = new ExcelApplicationAdapter(null).SubscribeWorkbookStateChanged(() => { });

        Assert.Null(Record.Exception(subscription.Dispose));
        Assert.Null(Record.Exception(subscription.Dispose));
    }

    [Fact]
    public void Subscribe_attaches_exactly_one_handler_to_each_workbook_state_event()
    {
        var application = new Mock<Excel.Application>();
        var events = application.As<Excel.AppEvents_Event>();

        using var subscription = new ExcelApplicationAdapter(application.Object)
            .SubscribeWorkbookStateChanged(() => { });

        events.VerifyAdd(e => e.WorkbookActivate += It.IsAny<Excel.AppEvents_WorkbookActivateEventHandler>(), Times.Once);
        events.VerifyAdd(e => e.WorkbookDeactivate += It.IsAny<Excel.AppEvents_WorkbookDeactivateEventHandler>(), Times.Once);
        events.VerifyAdd(e => e.NewWorkbook += It.IsAny<Excel.AppEvents_NewWorkbookEventHandler>(), Times.Once);
    }

    [Fact]
    public void Dispose_detaches_exactly_the_handlers_it_attached()
    {
        var application = new Mock<Excel.Application>();
        var events = application.As<Excel.AppEvents_Event>();
        var subscription = new ExcelApplicationAdapter(application.Object)
            .SubscribeWorkbookStateChanged(() => { });

        subscription.Dispose();

        events.VerifyRemove(e => e.WorkbookActivate -= It.IsAny<Excel.AppEvents_WorkbookActivateEventHandler>(), Times.Once);
        events.VerifyRemove(e => e.WorkbookDeactivate -= It.IsAny<Excel.AppEvents_WorkbookDeactivateEventHandler>(), Times.Once);
        events.VerifyRemove(e => e.NewWorkbook -= It.IsAny<Excel.AppEvents_NewWorkbookEventHandler>(), Times.Once);
    }

    [Fact]
    public void Dispose_is_idempotent_and_detaches_once()
    {
        var application = new Mock<Excel.Application>();
        var events = application.As<Excel.AppEvents_Event>();
        var subscription = new ExcelApplicationAdapter(application.Object)
            .SubscribeWorkbookStateChanged(() => { });

        subscription.Dispose();

        Assert.Null(Record.Exception(subscription.Dispose));
        events.VerifyRemove(e => e.WorkbookActivate -= It.IsAny<Excel.AppEvents_WorkbookActivateEventHandler>(), Times.Once);
        events.VerifyRemove(e => e.WorkbookDeactivate -= It.IsAny<Excel.AppEvents_WorkbookDeactivateEventHandler>(), Times.Once);
        events.VerifyRemove(e => e.NewWorkbook -= It.IsAny<Excel.AppEvents_NewWorkbookEventHandler>(), Times.Once);
    }

    [Fact]
    public void A_workbook_state_change_invokes_the_handler_once_per_event()
    {
        var application = new Mock<Excel.Application>();
        var events = application.As<Excel.AppEvents_Event>();
        var calls = 0;

        using var subscription = new ExcelApplicationAdapter(application.Object)
            .SubscribeWorkbookStateChanged(() => calls++);

        events.Raise(e => e.WorkbookActivate += null, new Mock<Excel.Workbook>().Object);
        events.Raise(e => e.NewWorkbook += null, new Mock<Excel.Workbook>().Object);
        events.Raise(e => e.WorkbookDeactivate += null, new Mock<Excel.Workbook>().Object);

        Assert.Equal(3, calls);
    }

    [Fact]
    public void A_throwing_handler_never_propagates_into_the_event_raiser()
    {
        var application = new Mock<Excel.Application>();
        var events = application.As<Excel.AppEvents_Event>();

        using var subscription = new ExcelApplicationAdapter(application.Object)
            .SubscribeWorkbookStateChanged(() => throw new InvalidOperationException("handler failed"));

        var exception = Record.Exception(
            () => events.Raise(e => e.WorkbookDeactivate += null, new Mock<Excel.Workbook>().Object));

        Assert.Null(exception);
    }

    [Fact]
    public void A_disposed_subscription_stops_receiving_events()
    {
        var application = new Mock<Excel.Application>();
        var events = application.As<Excel.AppEvents_Event>();
        var calls = 0;
        var subscription = new ExcelApplicationAdapter(application.Object)
            .SubscribeWorkbookStateChanged(() => calls++);

        subscription.Dispose();
        events.Raise(e => e.WorkbookActivate += null, new Mock<Excel.Workbook>().Object);

        Assert.Equal(0, calls);
    }

    [Fact]
    public void An_attached_subscription_reports_each_handler_as_attached()
    {
        var application = new Mock<Excel.Application>();
        using var subscription = new ExcelApplicationAdapter(application.Object)
            .SubscribeWorkbookStateChanged(() => { });

        var status = Assert.IsAssignableFrom<IWorkbookStateSubscriptionStatus>(subscription);
        Assert.Equal(EventHandlerState.Attached, status.WorkbookActivateState);
        Assert.Equal(EventHandlerState.Attached, status.WorkbookDeactivateState);
        Assert.Equal(EventHandlerState.Attached, status.NewWorkbookState);
        Assert.False(status.AllDetached);
    }

    [Fact]
    public void A_successful_detach_reports_every_handler_as_detached()
    {
        var application = new Mock<Excel.Application>();
        var subscription = new ExcelApplicationAdapter(application.Object)
            .SubscribeWorkbookStateChanged(() => { });

        subscription.Dispose();

        var status = Assert.IsAssignableFrom<IWorkbookStateSubscriptionStatus>(subscription);
        Assert.Equal(EventHandlerState.Detached, status.WorkbookActivateState);
        Assert.Equal(EventHandlerState.Detached, status.WorkbookDeactivateState);
        Assert.Equal(EventHandlerState.Detached, status.NewWorkbookState);
        Assert.True(status.AllDetached);
    }

    [Fact]
    public void A_failed_removal_is_reported_as_detach_failed_and_never_as_detached()
    {
        // The audit finding: teardown previously swallowed a failed removal and
        // reported success, hiding a leaked event connection. The state must show
        // the failure, and the other removals must still be attempted.
        var application = new Mock<Excel.Application>();
        var events = application.As<Excel.AppEvents_Event>();
        _ = events
            .SetupRemove(e => e.WorkbookActivate -= It.IsAny<Excel.AppEvents_WorkbookActivateEventHandler>())
            .Throws(new InvalidOperationException("removal refused"));
        var subscription = new ExcelApplicationAdapter(application.Object)
            .SubscribeWorkbookStateChanged(() => { });

        Assert.Null(Record.Exception(subscription.Dispose));

        var status = Assert.IsAssignableFrom<IWorkbookStateSubscriptionStatus>(subscription);
        Assert.Equal(EventHandlerState.DetachFailed, status.WorkbookActivateState);
        Assert.False(status.AllDetached);

        // One failure must not skip the remaining removals.
        Assert.Equal(EventHandlerState.Detached, status.WorkbookDeactivateState);
        Assert.Equal(EventHandlerState.Detached, status.NewWorkbookState);
        events.VerifyRemove(
            e => e.WorkbookDeactivate -= It.IsAny<Excel.AppEvents_WorkbookDeactivateEventHandler>(),
            Times.Once);
        events.VerifyRemove(
            e => e.NewWorkbook -= It.IsAny<Excel.AppEvents_NewWorkbookEventHandler>(),
            Times.Once);
    }

    [Fact]
    public void A_later_dispose_retries_a_handler_whose_removal_failed()
    {
        // The failed removal must stay retryable: the delegate is retained, so a
        // second Dispose removes the handler and reports the clean teardown.
        var application = new Mock<Excel.Application>();
        var events = application.As<Excel.AppEvents_Event>();
        var attempts = 0;
        _ = events
            .SetupRemove(e => e.WorkbookActivate -= It.IsAny<Excel.AppEvents_WorkbookActivateEventHandler>())
            .Callback(() =>
            {
                attempts++;
                if (attempts == 1)
                {
                    throw new InvalidOperationException("removal refused");
                }
            });
        var subscription = new ExcelApplicationAdapter(application.Object)
            .SubscribeWorkbookStateChanged(() => { });

        subscription.Dispose();
        var status = Assert.IsAssignableFrom<IWorkbookStateSubscriptionStatus>(subscription);
        Assert.Equal(EventHandlerState.DetachFailed, status.WorkbookActivateState);

        Assert.Null(Record.Exception(subscription.Dispose));

        Assert.Equal(2, attempts);
        Assert.Equal(EventHandlerState.Detached, status.WorkbookActivateState);
        Assert.True(status.AllDetached);
    }

    [Fact]
    public void A_retried_dispose_still_does_not_invoke_the_handler()
    {
        // Callbacks are disabled by the first dispose and must stay disabled
        // across the retry, so a leaked event connection cannot reach the handler.
        var application = new Mock<Excel.Application>();
        var events = application.As<Excel.AppEvents_Event>();
        var calls = 0;
        _ = events
            .SetupRemove(e => e.WorkbookActivate -= It.IsAny<Excel.AppEvents_WorkbookActivateEventHandler>())
            .Throws(new InvalidOperationException("removal refused"));
        var subscription = new ExcelApplicationAdapter(application.Object)
            .SubscribeWorkbookStateChanged(() => calls++);

        subscription.Dispose();
        subscription.Dispose();
        events.Raise(e => e.WorkbookActivate += null, new Mock<Excel.Workbook>().Object);

        Assert.Equal(0, calls);
    }

    [Fact]
    public void A_retried_dispose_does_not_remove_an_already_detached_handler_again()
    {
        // Only the failed removal is retried; a clean handler is not removed twice.
        var application = new Mock<Excel.Application>();
        var events = application.As<Excel.AppEvents_Event>();
        _ = events
            .SetupRemove(e => e.WorkbookDeactivate -= It.IsAny<Excel.AppEvents_WorkbookDeactivateEventHandler>())
            .Throws(new InvalidOperationException("removal refused"));
        var subscription = new ExcelApplicationAdapter(application.Object)
            .SubscribeWorkbookStateChanged(() => { });

        subscription.Dispose();
        subscription.Dispose();

        events.VerifyRemove(
            e => e.WorkbookActivate -= It.IsAny<Excel.AppEvents_WorkbookActivateEventHandler>(),
            Times.Once);
        events.VerifyRemove(
            e => e.WorkbookDeactivate -= It.IsAny<Excel.AppEvents_WorkbookDeactivateEventHandler>(),
            Times.Exactly(2));
    }

    [Fact]
    public void An_inert_subscription_reports_every_handler_as_detached()
    {
        using var subscription = new ExcelApplicationAdapter(null).SubscribeWorkbookStateChanged(() => { });

        var status = Assert.IsAssignableFrom<IWorkbookStateSubscriptionStatus>(subscription);
        Assert.True(status.AllDetached);
    }
}