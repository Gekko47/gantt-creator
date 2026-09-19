using ExcelDna.Integration.CustomUI;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contract tests for <see cref="RibbonStateService"/>: the dynamic getters are
/// deterministic and side-effect-free, the invalidate mechanism drives the
/// RibbonUI handle, workbook-state events refresh, and every failure path
/// degrades instead of throwing into Excel.
/// </summary>
/// <remarks>
/// The class shares the <c>diagnostics-service</c> collection because
/// <see cref="AddInHost.AutoOpen"/>/<see cref="AddInHost.AutoClose"/> now arm and
/// reset the same session singleton; xUnit runs a collection's classes
/// sequentially, so the host lifecycle tests cannot race these.
/// </remarks>
[Collection("diagnostics-service")]
public sealed class RibbonStateServiceTests : IDisposable
{
    public RibbonStateServiceTests() => RibbonStateService.Reset();

    public void Dispose()
    {
        RibbonStateService.Reset();
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Hand-written fake for <see cref="IExcelApplicationAdapter"/>: counts
    /// captures and subscriptions so the tests can prove the getters never probe
    /// and the subscription is attached and detached exactly once.
    /// </summary>
    private sealed class FakeApplicationAdapter : IExcelApplicationAdapter
    {
        private Action? _handler;

        public int CaptureCount { get; private set; }

        public int SubscribeCount { get; private set; }

        public int DisposeCount { get; private set; }

        public bool? WorkbookFact { get; set; }

        public bool ThrowOnCapture { get; set; }

        public bool ThrowOnSubscribe { get; set; }

        public bool ThrowOnDispose { get; set; }

        public bool? HasActiveWorkbook()
        {
            CaptureCount++;
            return ThrowOnCapture ? throw new InvalidOperationException("capture failed") : WorkbookFact;
        }

        public IDisposable SubscribeWorkbookStateChanged(Action handler)
        {
            ArgumentNullException.ThrowIfNull(handler);
            SubscribeCount++;
            if (ThrowOnSubscribe)
            {
                throw new InvalidOperationException("subscribe failed");
            }

            _handler = handler;
            return new Subscription(this);
        }

        internal void RaiseWorkbookStateChanged() => _handler?.Invoke();

        private sealed class Subscription(FakeApplicationAdapter owner) : IDisposable
        {
            private bool _disposed;

            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }

                _disposed = true;
                owner.DisposeCount++;
                owner._handler = null;
                if (owner.ThrowOnDispose)
                {
                    throw new InvalidOperationException("dispose failed");
                }
            }
        }
    }

    private static RibbonStateService Arm(
        FakeApplicationAdapter adapter,
        Func<bool?>? logSource = null,
        IRibbonUI? ribbon = null)
    {
        var service = RibbonStateService.Instance;
        service.SetApplicationAdapter(adapter);
        service.SetLogAvailabilitySource(logSource);
        service.SetRibbon(ribbon);
        service.Activate();
        return service;
    }

    private static int CountInvalidations(Mock<IRibbonUI> ribbon)
        => ribbon.Invocations.Count(invocation => invocation.Method.Name == nameof(IRibbonUI.Invalidate));

    [Fact]
    public void GetEnabled_is_deterministic_and_never_probes_the_state_source()
    {
        var adapter = new FakeApplicationAdapter { WorkbookFact = true };
        var service = Arm(adapter, () => true);

        var capturesAfterActivate = adapter.CaptureCount;
        var diagnostics = new List<bool>();
        var openLog = new List<bool>();
        for (var i = 0; i < 10; i++)
        {
            diagnostics.Add(service.GetEnabled(RibbonControlIds.Diagnostics));
            openLog.Add(service.GetEnabled(RibbonControlIds.OpenLog));
        }

        Assert.All(diagnostics, value => Assert.True(value));
        Assert.All(openLog, value => Assert.True(value));
        Assert.Equal(capturesAfterActivate, adapter.CaptureCount);
    }

    [Fact]
    public void GetEnabled_never_touches_the_ribbon_handle()
    {
        var adapter = new FakeApplicationAdapter { WorkbookFact = true };
        var ribbon = new Mock<IRibbonUI>();
        var service = Arm(adapter, () => true, ribbon.Object);
        var invalidationsAfterActivate = CountInvalidations(ribbon);

        _ = service.GetEnabled(RibbonControlIds.Diagnostics);
        _ = service.GetEnabled(RibbonControlIds.OpenLog);
        _ = service.GetEnabled("btnUnknown");

        Assert.Equal(invalidationsAfterActivate, CountInvalidations(ribbon));
    }

    [Fact]
    public void GetEnabled_maps_the_diagnostics_control_to_the_workbook_fact()
    {
        var adapter = new FakeApplicationAdapter { WorkbookFact = true };
        var service = Arm(adapter);

        Assert.True(service.GetEnabled(RibbonControlIds.Diagnostics));

        adapter.WorkbookFact = false;
        service.Refresh();

        Assert.False(service.GetEnabled(RibbonControlIds.Diagnostics));
    }

    [Fact]
    public void GetEnabled_maps_the_open_log_control_to_the_log_fact()
    {
        var adapter = new FakeApplicationAdapter { WorkbookFact = true };
        var logAvailable = true;
        var service = Arm(adapter, () => logAvailable);

        Assert.True(service.GetEnabled(RibbonControlIds.OpenLog));

        logAvailable = false;
        service.Refresh();

        Assert.False(service.GetEnabled(RibbonControlIds.OpenLog));
    }

    [Fact]
    public void GetEnabled_maps_the_initialise_sheet_control_to_the_workbook_fact()
    {
        // R2.2: the Initialise sheet button is gated on the same workbook fact
        // as Diagnostics — it only makes sense to initialise a workbook that
        // exists. Driven from the cached snapshot, never from a live probe.
        var adapter = new FakeApplicationAdapter { WorkbookFact = true };
        var service = Arm(adapter);

        Assert.True(service.GetEnabled(RibbonControlIds.InitialiseSheet));

        adapter.WorkbookFact = false;
        service.Refresh();

        Assert.False(service.GetEnabled(RibbonControlIds.InitialiseSheet));
    }

    [Fact]
    public void GetEnabled_initialise_sheet_starts_disabled_before_the_first_capture()
    {
        // The initial snapshot knows no facts, so every gated control —
        // including the R2.2 Initialise sheet button — starts grey. A control
        // that is enabled before the host arms the service is a regression.
        var fresh = RibbonStateService.Instance;

        Assert.False(fresh.GetEnabled(RibbonControlIds.InitialiseSheet));
    }

    [Fact]
    public void GetEnabled_is_enabled_for_unknown_null_and_whitespace_control_ids()
    {
        var service = Arm(new FakeApplicationAdapter { WorkbookFact = false }, () => false);

        Assert.True(service.GetEnabled("btnNotRegistered"));
        Assert.True(service.GetEnabled(null));
        Assert.True(service.GetEnabled("   "));
    }

    [Fact]
    public void Activate_subscribes_workbook_state_changes_once()
    {
        var adapter = new FakeApplicationAdapter { WorkbookFact = true };
        _ = Arm(adapter);

        Assert.Equal(1, adapter.SubscribeCount);
        Assert.Equal(0, adapter.DisposeCount);
    }

    [Fact]
    public void Activate_twice_replaces_the_subscription_instead_of_double_subscribing()
    {
        var adapter = new FakeApplicationAdapter { WorkbookFact = true };
        var service = Arm(adapter);

        service.Activate();

        Assert.Equal(2, adapter.SubscribeCount);
        Assert.Equal(1, adapter.DisposeCount);
    }

    [Fact]
    public void A_workbook_state_change_captures_once_and_invalidates_once()
    {
        var adapter = new FakeApplicationAdapter { WorkbookFact = false };
        var ribbon = new Mock<IRibbonUI>();
        _ = Arm(adapter, ribbon: ribbon.Object);
        var capturesBefore = adapter.CaptureCount;
        var invalidationsBefore = CountInvalidations(ribbon);

        adapter.RaiseWorkbookStateChanged();

        Assert.Equal(capturesBefore + 1, adapter.CaptureCount);
        Assert.Equal(invalidationsBefore + 1, CountInvalidations(ribbon));
    }

    [Fact]
    public void Refresh_captures_once_and_invalidates_once()
    {
        var adapter = new FakeApplicationAdapter { WorkbookFact = true };
        var ribbon = new Mock<IRibbonUI>();
        var service = Arm(adapter, ribbon: ribbon.Object);
        var capturesBefore = adapter.CaptureCount;
        var invalidationsBefore = CountInvalidations(ribbon);

        service.Refresh();

        Assert.Equal(capturesBefore + 1, adapter.CaptureCount);
        Assert.Equal(invalidationsBefore + 1, CountInvalidations(ribbon));
    }

    [Fact]
    public void A_failed_capture_keeps_the_previous_snapshot()
    {
        // Snapshot stickiness (work item R1.5): a throwing capture must not
        // drive the Ribbon from a transient failure.
        var adapter = new FakeApplicationAdapter { WorkbookFact = true };
        var service = Arm(adapter, () => true);
        adapter.ThrowOnCapture = true;

        var exception = Record.Exception(service.Refresh);

        Assert.Null(exception);
        Assert.True(service.GetEnabled(RibbonControlIds.Diagnostics));
        Assert.True(service.GetEnabled(RibbonControlIds.OpenLog));
    }

    [Fact]
    public void A_not_determinable_source_keeps_the_previous_value()
    {
        var adapter = new FakeApplicationAdapter { WorkbookFact = true };
        bool? logAvailable = true;
        var service = Arm(adapter, () => logAvailable);

        adapter.WorkbookFact = null;
        logAvailable = null;
        service.Refresh();

        Assert.True(service.GetEnabled(RibbonControlIds.Diagnostics));
        Assert.True(service.GetEnabled(RibbonControlIds.OpenLog));
    }

    [Fact]
    public void Reset_restores_the_initial_snapshot_and_detaches_the_subscription_exactly_once()
    {
        var adapter = new FakeApplicationAdapter { WorkbookFact = true };
        _ = Arm(adapter, () => true);

        RibbonStateService.Reset();

        Assert.Equal(1, adapter.DisposeCount);
        var fresh = RibbonStateService.Instance;
        Assert.False(fresh.GetEnabled(RibbonControlIds.Diagnostics));
        Assert.False(fresh.GetEnabled(RibbonControlIds.OpenLog));
    }

    [Fact]
    public void InvalidateControl_invalidates_the_control_once_for_a_valid_id()
    {
        var adapter = new FakeApplicationAdapter { WorkbookFact = true };
        var ribbon = new Mock<IRibbonUI>();
        var service = Arm(adapter, ribbon: ribbon.Object);

        service.InvalidateControl(RibbonControlIds.Diagnostics);

        Assert.Equal(
            1,
            ribbon.Invocations.Count(invocation => invocation.Method.Name == nameof(IRibbonUI.InvalidateControl)));
    }

    [Fact]
    public void InvalidateControl_is_a_no_op_for_null_or_whitespace_ids()
    {
        var adapter = new FakeApplicationAdapter { WorkbookFact = true };
        var ribbon = new Mock<IRibbonUI>();
        var service = Arm(adapter, ribbon: ribbon.Object);

        service.InvalidateControl(null);
        service.InvalidateControl(string.Empty);
        service.InvalidateControl("   ");

        Assert.Equal(
            0,
            ribbon.Invocations.Count(invocation => invocation.Method.Name == nameof(IRibbonUI.InvalidateControl)));
    }

    [Fact]
    public void InvalidateControl_never_throws_when_the_ribbon_call_fails()
    {
        var adapter = new FakeApplicationAdapter { WorkbookFact = true };
        var ribbon = new Mock<IRibbonUI>();
        ribbon
            .Setup(r => r.InvalidateControl(It.IsAny<string>()))
            .Throws(new InvalidOperationException("ribbon gone"));
        var service = Arm(adapter, ribbon: ribbon.Object);

        var exception = Record.Exception(() => service.InvalidateControl(RibbonControlIds.Diagnostics));

        Assert.Null(exception);
    }
}
