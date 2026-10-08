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
    public void GetEnabled_maps_the_validate_control_to_the_workbook_fact()
    {
        // R2.6: Validate reads the visible table, so it is gated on the same
        // workbook fact as Initialise sheet — driven from the cached snapshot.
        var adapter = new FakeApplicationAdapter { WorkbookFact = true };
        var service = Arm(adapter);

        Assert.True(service.GetEnabled(RibbonControlIds.ValidateSheet));

        adapter.WorkbookFact = false;
        service.Refresh();

        Assert.False(service.GetEnabled(RibbonControlIds.ValidateSheet));
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

    [Fact]
    public void TogglePlotStartAuto_from_explicit_to_auto_persists_and_preserves_the_start_date()
    {
        // Fix plan ruling 1: entering AUTO no longer clears the stored date —
        // the disabled edit box keeps showing it until the effective-date
        // derivation replaces it, and returning to explicit restores it. The
        // R5.1 selection-change invariant holds: no render; the persisted
        // payload is the assertion, not a ribbon refresh, which Set never
        // triggers. The finish end — still explicit — is preserved untouched.
        var service = RibbonStateService.Instance;
        service.SetPlotStartDate("15/03/2026");
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        _ = writer
            .Setup(w => w.WriteSettings(It.Is<IReadOnlyDictionary<string, string>>(
                settings =>
                    settings["PlotStartMode"] == "DataRange"
                    && settings["PlotFinishMode"] == "DataRange"
                    && settings["PlotStartDate"] == "15/03/2026"
                    && settings["PlotFinishDate"] == string.Empty)))
            .Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);

        service.TogglePlotStartAuto();

        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Once);
        Assert.True(service.IsPlotStartAuto());
        Assert.Equal("15/03/2026", service.GetPlotStartDate());
    }

    [Fact]
    public void TogglePlotStartAuto_from_auto_to_explicit_arms_the_edit_box_without_persisting()
    {
        // Unchecking AUTO must never persist Explicit with no date: an
        // Explicit mode without a date is the MissingExplicitDate refusal
        // and would break the next Refresh. The date commit owns that write,
        // so this toggle changes only the in-memory state that drives
        // getChecked/getEnabled.
        var service = RibbonStateService.Instance;
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        service.SetCatalogueWriter(writer.Object);

        service.TogglePlotStartAuto();

        Assert.False(service.IsPlotStartAuto());
        Assert.Equal(string.Empty, service.GetPlotStartDate());
        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Never);
    }

    [Fact]
    public void TogglePlotStartAuto_degrades_gracefully_without_a_writer()
    {
        // A null writer leaves the persist step as a no-op: the in-memory
        // state commit must not throw and no refresh is required. Start from
        // explicit so the toggle actually reaches the persist step.
        var service = RibbonStateService.Instance;
        service.SetCatalogueWriter(null);
        service.SetPlotStartDate("15/03/2026");

        var exception = Record.Exception(service.TogglePlotStartAuto);

        Assert.Null(exception);
        Assert.True(service.IsPlotStartAuto());
    }

    [Fact]
    public void TogglePlotFinishAuto_from_explicit_to_auto_persists_and_preserves_the_finish_date()
    {
        // Per-end analogue of the ruling-1 preserve: the start end (also
        // explicit) is preserved untouched in the payload.
        var service = RibbonStateService.Instance;
        service.SetPlotStartDate("01/03/2026");
        service.SetPlotFinishDate("15/06/2026");
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        _ = writer
            .Setup(w => w.WriteSettings(It.Is<IReadOnlyDictionary<string, string>>(
                settings =>
                    settings["PlotStartMode"] == "Explicit"
                    && settings["PlotFinishMode"] == "DataRange"
                    && settings["PlotStartDate"] == "01/03/2026"
                    && settings["PlotFinishDate"] == "15/06/2026")))
            .Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);

        service.TogglePlotFinishAuto();

        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Once);
        Assert.True(service.IsPlotFinishAuto());
        Assert.Equal("15/06/2026", service.GetPlotFinishDate());
        Assert.Equal("01/03/2026", service.GetPlotStartDate());
    }

    [Fact]
    public void TogglePlotFinishAuto_from_auto_to_explicit_arms_the_edit_box_without_persisting()
    {
        // Mirror of the start-side rule: no Explicit-without-date write.
        var service = RibbonStateService.Instance;
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        service.SetCatalogueWriter(writer.Object);

        service.TogglePlotFinishAuto();

        Assert.False(service.IsPlotFinishAuto());
        Assert.Equal(string.Empty, service.GetPlotFinishDate());
        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Never);
    }

    [Fact]
    public void TogglePlotFinishAuto_degrades_gracefully_without_a_writer()
    {
        var service = RibbonStateService.Instance;
        service.SetCatalogueWriter(null);
        service.SetPlotFinishDate("15/06/2026");

        var exception = Record.Exception(service.TogglePlotFinishAuto);

        Assert.Null(exception);
        Assert.True(service.IsPlotFinishAuto());
    }

    [Fact]
    public void SetPlotStartDate_normalises_ddMMyyyy_and_pivots_to_explicit()
    {
        // A valid typed date pivots the start end to Explicit and normalises
        // to the repository dd/MM/yyyy storage form, which the getText getter
        // then returns.
        var service = RibbonStateService.Instance;
        service.SetCatalogueWriter(null);

        service.SetPlotStartDate("15/03/2026");

        Assert.False(service.IsPlotStartAuto());
        Assert.Equal("15/03/2026", service.GetPlotStartDate());
    }

    [Fact]
    public void SetPlotStartDate_refuses_activity_auto_parser_equivalents()
    {
        // Positive for the exact-date rule: the plot boxes take the chart's
        // dd/MM/yyyy dates only, never the activity multi-format equivalents.
        // ISO, US month-first, and dd-MMM-yyyy entries are no-ops that leave
        // the stored value, mode, and writer untouched.
        var service = RibbonStateService.Instance;
        service.SetCatalogueWriter(null);
        service.SetPlotStartDate("15/03/2026");
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        service.SetCatalogueWriter(writer.Object);

        service.SetPlotStartDate("2026-03-15");
        service.SetPlotStartDate("03/15/2026");
        service.SetPlotStartDate("15-Mar-2026");

        Assert.False(service.IsPlotStartAuto());
        Assert.Equal("15/03/2026", service.GetPlotStartDate());
        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Never);
    }

    [Fact]
    public void SetPlotStartDate_reverts_on_an_unparseable_entry()
    {
        // Positive for the revert rule: a non-date leaves the last valid
        // stored value and mode untouched, and never reaches the writer, so
        // the edit box keeps showing the previously stored date on the next
        // getText query.
        var service = RibbonStateService.Instance;
        service.SetPlotStartDate("15/03/2026");
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        service.SetCatalogueWriter(writer.Object);

        service.SetPlotStartDate("not-a-date");

        Assert.False(service.IsPlotStartAuto());
        Assert.Equal("15/03/2026", service.GetPlotStartDate());
        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Never);
    }

    [Fact]
    public void SetPlotStartDate_ignores_blank_input()
    {
        // Positive: clearing the box is a no-op rather than a blank commit.
        var service = RibbonStateService.Instance;
        service.SetPlotStartDate("15/03/2026");
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        service.SetCatalogueWriter(writer.Object);

        service.SetPlotStartDate("   ");

        Assert.Equal("15/03/2026", service.GetPlotStartDate());
        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Never);
    }

    [Fact]
    public void SetPlotFinishDate_normalises_ddMMyyyy_and_pivots_to_explicit()
    {
        var service = RibbonStateService.Instance;
        service.SetCatalogueWriter(null);

        service.SetPlotFinishDate("30/06/2026");

        Assert.False(service.IsPlotFinishAuto());
        Assert.Equal("30/06/2026", service.GetPlotFinishDate());
    }

    [Fact]
    public void SetPlotFinishDate_reverts_on_an_unparseable_entry()
    {
        // Positive for the revert rule on the finish end.
        var service = RibbonStateService.Instance;
        service.SetPlotFinishDate("30/06/2026");
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        service.SetCatalogueWriter(writer.Object);

        service.SetPlotFinishDate("31/02/2026");

        Assert.False(service.IsPlotFinishAuto());
        Assert.Equal("30/06/2026", service.GetPlotFinishDate());
        writer.Verify(
            w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()),
            Times.Never);
    }
}
