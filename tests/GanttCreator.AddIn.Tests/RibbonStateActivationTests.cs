
using ExcelDna.Integration.CustomUI;
using GanttCreator.Office;
using Moq;
using static GanttCreator.AddIn.Tests.RibbonStateTestSupport;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Activation, capture and invalidation for <see cref="RibbonStateService"/>: one subscription per
/// activation, one capture and one invalidation per workbook-state event, the previous snapshot kept
/// when a capture fails, and the invalidate call that never throws into Excel.
/// </summary>
/// <remarks>
/// Shares the <c>diagnostics-service</c> collection: these facts arm the same session singleton the
/// host-lifecycle facts arm and reset, so xUnit runs the collection's classes sequentially. The
/// constructor and <see cref="Dispose"/> reset it around every fact.
/// </remarks>
[Collection("diagnostics-service")]
public sealed class RibbonStateActivationTests : IDisposable
{
    public RibbonStateActivationTests() => RibbonStateService.Reset();

    public void Dispose()
    {
        RibbonStateService.Reset();
        GC.SuppressFinalize(this);
    }
    [Fact]
    public void Activate_subscribes_workbook_state_changes_once()
    {
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
        _ = Arm(adapter);

        Assert.Equal(1, adapter.WorkbookStateSubscribeCount);
        Assert.Equal(0, adapter.WorkbookStateDisposeCount);
    }

    [Fact]
    public void Activate_twice_replaces_the_subscription_instead_of_double_subscribing()
    {
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
        var service = Arm(adapter);

        service.Activate();

        Assert.Equal(2, adapter.WorkbookStateSubscribeCount);
        Assert.Equal(1, adapter.WorkbookStateDisposeCount);
    }

    [Fact]
    public void Activate_keeps_an_injected_adapter_even_when_the_host_application_is_live()
    {
        // Regression (2026-11-02). Activate used to re-probe
        // ExcelDnaUtil.Application unconditionally and replace whatever adapter
        // had been injected. That probe resolves the RUNNING Excel instance, so
        // simply having Excel open on the machine replaced the fake with a real
        // adapter over the developer's workbook: WorkbookStateSubscribeCount read 0, the
        // captured facts came from that workbook, and 12 tests went red without
        // any source change. An injected adapter is authoritative.
        var service = new HostApplicationRibbonStateService(new object());
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
        service.SetApplicationAdapter(adapter);

        service.Activate();

        Assert.Same(adapter, service.GetApplicationAdapter());
        Assert.Equal(1, adapter.WorkbookStateSubscribeCount);
    }

    [Fact]
    public void Activate_adopts_the_host_application_when_no_adapter_was_injected()
    {
        // The other half of the same contract. AddInHost injects no adapter when
        // Excel-DNA has not exposed an application yet (AutoOpen can run first),
        // so Activate must supply one once the application becomes available;
        // the ribbon onLoad and the queued fallback both rely on that.
        var service = new HostApplicationRibbonStateService(new object());
        Assert.Null(service.GetApplicationAdapter());

        service.Activate();

        Assert.IsType<ExcelApplicationAdapter>(service.GetApplicationAdapter());
    }

    [Fact]
    public void A_workbook_state_change_captures_once_and_invalidates_once()
    {
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = false };
        var ribbon = new Mock<IRibbonUI>();
        _ = Arm(adapter, ribbon: ribbon.Object);
        var capturesBefore = adapter.ActiveWorkbookReads;
        var invalidationsBefore = CountInvalidations(ribbon);

        adapter.RaiseWorkbookStateChanged();

        Assert.Equal(capturesBefore + 1, adapter.ActiveWorkbookReads);
        Assert.Equal(invalidationsBefore + 1, CountInvalidations(ribbon));
    }

    [Fact]
    public void Refresh_captures_once_and_invalidates_once()
    {
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
        var ribbon = new Mock<IRibbonUI>();
        var service = Arm(adapter, ribbon: ribbon.Object);
        var capturesBefore = adapter.ActiveWorkbookReads;
        var invalidationsBefore = CountInvalidations(ribbon);

        service.Refresh();

        Assert.Equal(capturesBefore + 1, adapter.ActiveWorkbookReads);
        Assert.Equal(invalidationsBefore + 1, CountInvalidations(ribbon));
    }

    [Fact]
    public void A_failed_capture_keeps_the_previous_snapshot()
    {
        // Snapshot stickiness (work item R1.5): a throwing capture must not
        // drive the Ribbon from a transient failure.
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
        var service = Arm(adapter, () => true);
        adapter.ThrowOnHasActiveWorkbook = true;

        var exception = Record.Exception(service.Refresh);

        Assert.Null(exception);
        Assert.True(service.GetEnabled(RibbonControlIds.Diagnostics));
        Assert.True(service.GetEnabled(RibbonControlIds.OpenLog));
    }

    [Fact]
    public void A_not_determinable_source_keeps_the_previous_value()
    {
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
        bool? logAvailable = true;
        var service = Arm(adapter, () => logAvailable);

        adapter.ActiveWorkbook = null;
        logAvailable = null;
        service.Refresh();

        Assert.True(service.GetEnabled(RibbonControlIds.Diagnostics));
        Assert.True(service.GetEnabled(RibbonControlIds.OpenLog));
    }

    [Fact]
    public void Reset_restores_the_initial_snapshot_and_detaches_the_subscription_exactly_once()
    {
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
        _ = Arm(adapter, () => true);

        RibbonStateService.Reset();

        Assert.Equal(1, adapter.WorkbookStateDisposeCount);
        var fresh = RibbonStateService.Instance;
        Assert.False(fresh.GetEnabled(RibbonControlIds.Diagnostics));
        Assert.False(fresh.GetEnabled(RibbonControlIds.OpenLog));
    }

    [Fact]
    public void InvalidateControl_invalidates_the_control_once_for_a_valid_id()
    {
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
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
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
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
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
        var ribbon = new Mock<IRibbonUI>();
        ribbon
            .Setup(r => r.InvalidateControl(It.IsAny<string>()))
            .Throws(new InvalidOperationException("ribbon gone"));
        var service = Arm(adapter, ribbon: ribbon.Object);

        var exception = Record.Exception(() => service.InvalidateControl(RibbonControlIds.Diagnostics));

        Assert.Null(exception);
    }

    /// <summary>
    /// Test double that supplies a fixed host application object instead of querying
    /// <c>ExcelDnaUtil.Application</c>. The production probe resolves whichever Excel instance is
    /// running, so a test that used it would pass or fail according to the developer's machine
    /// rather than the code.
    /// </summary>
    private sealed class HostApplicationRibbonStateService(object? hostApplication) : RibbonStateService
    {
        internal override object? ResolveHostApplication() => hostApplication;
    }
}
