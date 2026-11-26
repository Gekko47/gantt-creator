using GanttCreator.Office;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The event-lifetime seam (finding 3): subscribing replaces rather than doubles, detaching
/// happens once, and a host that refuses a subscribe degrades to an inert connection instead of
/// failing the add-in load.
/// </summary>
public class RibbonEventWiringTests
{
    [Fact]
    public void Subscribe_attaches_every_source_and_detach_disposes_each_once()
    {
        FakeAdapter adapter = new();
        FakeSelection selection = new();
        RibbonEventWiring wiring = new(adapter, selection, () => { }, () => { }, () => { });

        wiring.Subscribe();

        Assert.Equal(1, adapter.SubscribeCount);
        Assert.Equal(1, adapter.SheetActivationSubscribeCount);
        Assert.Equal(1, selection.SubscribeCount);

        wiring.DetachAll();
        wiring.DetachAll();

        Assert.Equal(1, adapter.DisposeCount);
        Assert.Equal(1, adapter.SheetActivationDisposeCount);
        Assert.Equal(1, selection.DisposeCount);
    }

    [Fact]
    public void Re_subscribing_replaces_the_connection_rather_than_doubling_it()
    {
        FakeAdapter adapter = new();
        FakeSelection selection = new();
        RibbonEventWiring wiring = new(adapter, selection, () => { }, () => { }, () => { });

        wiring.Subscribe();
        wiring.Subscribe();

        Assert.Equal(2, adapter.SubscribeCount);
        Assert.Equal(1, adapter.DisposeCount);
        Assert.Equal(1, selection.DisposeCount);

        wiring.DetachAll();

        Assert.Equal(2, adapter.DisposeCount);
        Assert.Equal(2, selection.DisposeCount);
    }

    [Fact]
    public void A_refused_subscribe_is_inert_and_retried_on_the_next_attach()
    {
        FakeAdapter adapter = new() { RefuseSubscribes = true };
        RibbonEventWiring wiring = new(adapter, new FakeSelection(), () => { }, () => { }, () => { });

        wiring.Subscribe();

        Assert.Equal(1, adapter.SubscribeCount);
        Assert.Equal(0, adapter.DisposeCount);

        wiring.Subscribe();

        Assert.Equal(2, adapter.SubscribeCount);
        wiring.DetachAll();
        Assert.Equal(0, adapter.DisposeCount);
    }

    [Fact]
    public void Without_a_scheduler_the_refresh_requests_are_inert()
    {
        // No sources at all stands in for the test host, where no scheduler can be created: the
        // two request paths must stay no-ops rather than throw into a command or an event.
        RibbonEventWiring wiring = new(null, null, () => { }, () => { }, () => { });
        wiring.Subscribe();

        wiring.ScheduleCommandRefresh();
        wiring.CancelPendingRefresh();
        wiring.DetachAll();
    }

    [Fact]
    public void Null_handlers_are_rejected()
    {
        _ = Assert.Throws<ArgumentNullException>(
            () => _ = new RibbonEventWiring(null, null, null!, () => { }, () => { }));
        _ = Assert.Throws<ArgumentNullException>(
            () => _ = new RibbonEventWiring(null, null, () => { }, null!, () => { }));
        _ = Assert.Throws<ArgumentNullException>(
            () => _ = new RibbonEventWiring(null, null, () => { }, () => { }, null!));
    }

    private sealed class FakeAdapter : IExcelApplicationAdapter
    {
        public bool RefuseSubscribes { get; init; }

        public int SubscribeCount { get; private set; }

        public int DisposeCount { get; private set; }

        public int SheetActivationSubscribeCount { get; private set; }

        public int SheetActivationDisposeCount { get; private set; }

        public bool? HasActiveWorkbook() => true;

        public void ShowStatusBar(string message) => _ = message;

        public void ClearStatusBar()
        {
            // Nothing to clear.
        }

        public IDisposable SubscribeWorkbookStateChanged(Action handler)
        {
            SubscribeCount++;
            return RefuseSubscribes
                ? throw new InvalidOperationException("the host refused the connection")
                : new CountingSubscription(() => DisposeCount++);
        }

        public IDisposable SubscribeSheetActivated(Action handler)
        {
            SheetActivationSubscribeCount++;
            return RefuseSubscribes
                ? throw new InvalidOperationException("the host refused the connection")
                : new CountingSubscription(() => SheetActivationDisposeCount++);
        }
    }

    private sealed class FakeSelection : ISelectionIdentifier
    {
        public int SubscribeCount { get; private set; }

        public int DisposeCount { get; private set; }

        public SelectionFacts Identify() => new(
            HasSelection: true,
            IsSingleBodyRow: true,
            BodyIndex: 1,
            IsHeaderRow: false,
            SelectedRowCount: 1);

        public IDisposable SubscribeSelectionChanged(Action handler)
        {
            SubscribeCount++;
            return new CountingSubscription(() => DisposeCount++);
        }
    }

    private sealed class CountingSubscription(Action onDispose) : IDisposable
    {
        private Action? _onDispose = onDispose;

        public void Dispose()
        {
            Action? onDispose = _onDispose;
            _onDispose = null;
            onDispose?.Invoke();
        }
    }
}
