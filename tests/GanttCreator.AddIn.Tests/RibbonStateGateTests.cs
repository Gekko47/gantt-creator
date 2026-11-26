
using ExcelDna.Integration.CustomUI;
using GanttCreator.Office;
using Moq;
using static GanttCreator.AddIn.Tests.RibbonStateTestSupport;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The gate half of <see cref="RibbonStateService"/>: which groups a mode hides, and which controls
/// <c>GetEnabled</c> reports — the workbook, log and initialised facts, the critical-interval rule,
/// and the fail-open answers for ids the service does not know.
/// </summary>
/// <remarks>
/// Shares the <c>diagnostics-service</c> collection: these facts arm the same session singleton the
/// host-lifecycle facts arm and reset, so xUnit runs the collection's classes sequentially. The
/// constructor and <see cref="Dispose"/> reset it around every fact.
/// </remarks>
[Collection("diagnostics-service")]
public sealed class RibbonStateGateTests : IDisposable
{
    public RibbonStateGateTests() => RibbonStateService.Reset();

    public void Dispose()
    {
        RibbonStateService.Reset();
        GC.SuppressFinalize(this);
    }
    [Fact]
    public void GetVisible_hides_a_group_it_does_not_know()
    {
        var service = Arm(new TestExcelApplicationAdapter { ActiveWorkbook = true });

        Assert.False(service.GetVisible("grpUnknown"));
        Assert.False(service.GetVisible(null));
    }
    [Fact]
    public void Critical_interval_is_enabled_when_the_table_has_an_eligible_parent()
    {
        var service = ArmedWithTableRead(
            new TestExcelApplicationAdapter { ActiveWorkbook = true },
            GanttTableReadOutcome.Ok([Row(1, "As-Planned Activity")]));

        Assert.True(service.GetEnabled(RibbonControlIds.CriticalIntervalMenuItem));
    }

    [Fact]
    public void Critical_interval_is_disabled_when_no_row_can_own_it()
    {
        // Positive test for the gate: a table of leaves and structural rows has
        // no parent, so the item is disabled rather than raising an error after
        // the click. This is the "refuse only when none exists" rule.
        var service = ArmedWithTableRead(
            new TestExcelApplicationAdapter { ActiveWorkbook = true },
            GanttTableReadOutcome.Ok(
            [
                Row(1, "As-Planned Milestone"),
                Row(2, "Spacer"),
                Row(3, "Splitter"),
            ]));

        Assert.False(service.GetEnabled(RibbonControlIds.CriticalIntervalMenuItem));
    }

    [Fact]
    public void Critical_interval_is_disabled_when_the_table_is_empty()
    {
        var service = ArmedWithTableRead(
            new TestExcelApplicationAdapter { ActiveWorkbook = true },
            GanttTableReadOutcome.Ok([]));

        Assert.False(service.GetEnabled(RibbonControlIds.CriticalIntervalMenuItem));
    }

    [Fact]
    public void GetEnabled_is_deterministic_and_never_probes_the_state_source()
    {
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
        var service = Arm(adapter, () => true);

        var capturesAfterActivate = adapter.ActiveWorkbookReads;
        var diagnostics = new List<bool>();
        var openLog = new List<bool>();
        for (var i = 0; i < 10; i++)
        {
            diagnostics.Add(service.GetEnabled(RibbonControlIds.Diagnostics));
            openLog.Add(service.GetEnabled(RibbonControlIds.OpenLog));
        }

        Assert.All(diagnostics, value => Assert.True(value));
        Assert.All(openLog, value => Assert.True(value));
        Assert.Equal(capturesAfterActivate, adapter.ActiveWorkbookReads);
    }

    [Fact]
    public void GetEnabled_never_touches_the_ribbon_handle()
    {
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
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
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
        var service = Arm(adapter);

        Assert.True(service.GetEnabled(RibbonControlIds.Diagnostics));

        adapter.ActiveWorkbook = false;
        service.Refresh();

        Assert.False(service.GetEnabled(RibbonControlIds.Diagnostics));
    }

    [Fact]
    public void GetEnabled_maps_the_open_log_control_to_the_log_fact()
    {
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
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
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
        var service = Arm(adapter);

        Assert.True(service.GetEnabled(RibbonControlIds.InitialiseSheet));

        adapter.ActiveWorkbook = false;
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
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
        var service = Arm(adapter);

        Assert.True(service.GetEnabled(RibbonControlIds.ValidateSheet));

        adapter.ActiveWorkbook = false;
        service.Refresh();

        Assert.False(service.GetEnabled(RibbonControlIds.ValidateSheet));
    }

    [Fact]
    public void GetEnabled_is_enabled_for_unknown_null_and_whitespace_control_ids()
    {
        var service = Arm(new TestExcelApplicationAdapter { ActiveWorkbook = false }, () => false);

        Assert.True(service.GetEnabled("btnNotRegistered"));
        Assert.True(service.GetEnabled(null));
        Assert.True(service.GetEnabled("   "));
    }
}
