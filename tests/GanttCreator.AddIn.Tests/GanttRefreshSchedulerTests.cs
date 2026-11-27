using System.Collections.Concurrent;
using ExcelDna.Integration;
using GanttCreator.AddIn;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

public sealed class GanttRefreshSchedulerTests
{
    [Fact]
    public void Excel_dna_queue_signature_is_excel_action_to_void()
    {
        System.Reflection.MethodInfo? method =
            typeof(ExcelAsyncUtil).GetMethod(nameof(ExcelAsyncUtil.QueueAsMacro), [typeof(ExcelAction)]);
        Assert.NotNull(method);
        Assert.Equal(typeof(void), method.ReturnType);
        Assert.Equal(typeof(ExcelAction), Assert.Single(method.GetParameters()).ParameterType);
    }

    [Fact]
    public void Constructor_rejects_null_dependencies_and_adapters()
    {
        IExcelApplicationAdapter adapter = Mock.Of<IExcelApplicationAdapter>();
        ISheetChangeAnalyzer analyzer = Mock.Of<ISheetChangeAnalyzer>();
        Action<ExcelAction> queue = _ => { };
        Action refresh = () => { };

        Assert.Throws<ArgumentNullException>(() => new GanttRefreshScheduler(null!, analyzer, queue, refresh));
        Assert.Throws<ArgumentNullException>(() => new GanttRefreshScheduler(adapter, null!, queue, refresh));
        Assert.Throws<ArgumentNullException>(() => new GanttRefreshScheduler(adapter, analyzer, null!, refresh));
        Assert.Throws<ArgumentNullException>(() => new GanttRefreshScheduler(adapter, analyzer, queue, null!));
    }

    [Fact]
    public void Multiple_schedules_coalesce_to_one_macro_refresh()
    {
        using ManualResetEventSlim queuedSignal = new(false);
        using ManualResetEventSlim secondQueueSignal = new(false);
        ConcurrentQueue<ExcelAction> macros = new();
        int queuedCount = 0;
        int refreshCount = 0;
        Action<ExcelAction> queue = macro =>
        {
            macros.Enqueue(macro);
            if (Interlocked.Increment(ref queuedCount) == 1)
            {
                queuedSignal.Set();
            }
            else
            {
                secondQueueSignal.Set();
            }
        };

        using GanttRefreshScheduler scheduler = CreateScheduler(queue, () => Interlocked.Increment(ref refreshCount));
        scheduler.ScheduleCommandRefresh();
        scheduler.ScheduleCommandRefresh();

        Assert.True(queuedSignal.Wait(TimeSpan.FromSeconds(5)), "The debounced macro was not queued before the deadline.");
        Assert.True(macros.TryDequeue(out ExcelAction? macro));
        macro!();
        Assert.Equal(1, Volatile.Read(ref refreshCount));

        // Observe the named macro-queue signal for longer than the debounce
        // window so a stale callback from the superseded timer cannot pass.
        Assert.False(secondQueueSignal.Wait(TimeSpan.FromMilliseconds(500)), "A superseded timer queued a second macro.");
        Assert.Equal(1, Volatile.Read(ref queuedCount));
    }

    [Fact]
    public void Cancelled_queued_macro_does_not_refresh()
    {
        using ManualResetEventSlim queuedSignal = new(false);
        ExcelAction? queuedMacro = null;
        int refreshCount = 0;
        Action<ExcelAction> queue = macro =>
        {
            queuedMacro = macro;
            queuedSignal.Set();
        };

        using GanttRefreshScheduler scheduler = CreateScheduler(queue, () => Interlocked.Increment(ref refreshCount));
        scheduler.ScheduleCommandRefresh();

        Assert.True(queuedSignal.Wait(TimeSpan.FromSeconds(5)), "The debounced macro was not queued before the deadline.");
        scheduler.CancelPendingRefresh();
        Assert.NotNull(queuedMacro);
        queuedMacro!();

        Assert.Equal(0, Volatile.Read(ref refreshCount));
    }

    /// <summary>
    /// The post-command window is shorter than the cell-edit window.
    /// </summary>
    /// <remarks>
    /// A command has already committed its writes, so the only events left to absorb are
    /// its own; waiting the full cell window was 300ms of dead time between the click and
    /// the plot. The relationship is pinned rather than the absolute figure, so tuning
    /// either window stays deliberate.
    /// </remarks>
    [Fact]
    public void The_command_window_is_shorter_than_the_cell_window()
    {
        Assert.True(GanttRefreshScheduler.CommandDebounce > TimeSpan.Zero);
        Assert.True(GanttRefreshScheduler.CommandDebounce < GanttRefreshScheduler.CellDebounce);
    }

    /// <summary>
    /// A command's own cell writes do not restart the longer window, so one click
    /// produces exactly one refresh.
    /// </summary>
    [Fact]
    public void A_command_schedule_absorbs_the_cells_its_own_write_emitted()
    {
        using ManualResetEventSlim queuedSignal = new(false);
        using ManualResetEventSlim secondQueueSignal = new(false);
        ConcurrentQueue<ExcelAction> macros = new();
        Action<string?>? onChanged = null;
        int queuedCount = 0;
        int refreshCount = 0;

        var adapter = new Mock<IExcelApplicationAdapter>();
        _ = adapter
            .Setup(a => a.SubscribeSheetChanged(It.IsAny<Action<string?>>()))
            .Callback<Action<string?>>(handler => onChanged = handler)
            .Returns(Mock.Of<IDisposable>());
        var analyzer = new Mock<ISheetChangeAnalyzer>();
        _ = analyzer.Setup(a => a.ShouldAutoRefresh(It.IsAny<string?>())).Returns(true);

        using GanttRefreshScheduler scheduler = new(
            adapter.Object,
            analyzer.Object,
            macro =>
            {
                macros.Enqueue(macro);
                if (Interlocked.Increment(ref queuedCount) == 1)
                {
                    queuedSignal.Set();
                }
                else
                {
                    secondQueueSignal.Set();
                }
            },
            () => Interlocked.Increment(ref refreshCount));
        scheduler.Activate();

        scheduler.ScheduleCommandRefresh();
        Assert.NotNull(onChanged);
        onChanged!("\"book|Sheet\"!D5");

        Assert.True(queuedSignal.Wait(TimeSpan.FromSeconds(5)), "The command refresh was not queued before the deadline.");
        Assert.True(macros.TryDequeue(out ExcelAction? macro));
        macro!();

        Assert.Equal(1, Volatile.Read(ref refreshCount));
        Assert.Equal(1, Volatile.Read(ref queuedCount));

        // The flag is cleared for the next genuine edit: a cell event after the pass
        // schedules a new refresh rather than being swallowed. Observed through the
        // second signal, because the first signal is already set and would make the wait
        // above pass vacuously.
        onChanged!("\"book|Sheet\"!D5");
        Assert.True(secondQueueSignal.Wait(TimeSpan.FromSeconds(5)), "A later edit did not schedule a refresh.");
        Assert.Equal(2, Volatile.Read(ref queuedCount));
    }

    /// <summary>
    /// ADR-0055 clause 7: while the CHART toggle is off, the RENDER is suspended for both a cell
    /// edit and a command, because either would rebuild the chart over a manual drag. The ribbon
    /// projection is a read of the table, not a rebuild, and is never suspended (owner ruling
    /// 2026-11-27), so the same edit still queues one projection. Unlocking the sheet's CELLS is
    /// deliberately not a suspension: a cell edit is the input an automatic refresh exists to
    /// consume.
    /// </summary>
    [Fact]
    public void An_unlocked_chart_suspends_the_render_but_not_the_projection()
    {
        var preferences = new ProtectionPreferences();
        int queued = 0;
        int rendered = 0;
        int projected = 0;
        ConcurrentQueue<ExcelAction> macros = new();
        Mock<IExcelApplicationAdapter> adapter = new();
        Action<string?>? onChanged = null;
        _ = adapter
            .Setup(a => a.SubscribeSheetChanged(It.IsAny<Action<string?>>()))
            .Callback<Action<string?>>(handler => onChanged = handler)
            .Returns(Mock.Of<IDisposable>());
        Mock<ISheetChangeAnalyzer> analyzer = new();
        _ = analyzer.Setup(a => a.ShouldAutoRefresh(It.IsAny<string?>())).Returns(true);
        using GanttRefreshScheduler scheduler = new(
            adapter.Object,
            analyzer.Object,
            macro =>
            {
                macros.Enqueue(macro);
                _ = Interlocked.Increment(ref queued);
            },
            () => Interlocked.Increment(ref rendered),
            preferences,
            stateRefresh: () => Interlocked.Increment(ref projected));
        scheduler.Activate();

        // Positive control first: with the chart locked both paths render, so the assertions below
        // cannot pass on a harness that simply never queues anything.
        scheduler.ScheduleCommandRefresh();
        Assert.True(
            SpinWait.SpinUntil(() => Volatile.Read(ref queued) == 1, TimeSpan.FromSeconds(5)),
            "With the chart locked, a command did not schedule a refresh.");
        Assert.True(macros.TryDequeue(out ExcelAction? first));
        first!();
        Assert.Equal(1, Volatile.Read(ref rendered));

        // Unlocked: a command schedules nothing, and a cell edit projects without rendering.
        preferences.ChartLocked = false;
        int before = Volatile.Read(ref queued);
        scheduler.ScheduleCommandRefresh();
        onChanged!("\"book|Sheet\"!D5");

        Assert.True(
            SpinWait.SpinUntil(() => Volatile.Read(ref queued) > before, TimeSpan.FromSeconds(5)),
            "An unlocked chart queued nothing for the edit; the projection must still run.");
        Assert.True(macros.TryDequeue(out ExcelAction? projection));
        projection!();
        Assert.Equal(1, Volatile.Read(ref rendered));
        Assert.Equal(1, Volatile.Read(ref projected));

        // Locking the chart again restores the render.
        preferences.ChartLocked = true;
        int beforeRelock = Volatile.Read(ref queued);
        scheduler.ScheduleCommandRefresh();
        Assert.True(
            SpinWait.SpinUntil(() => Volatile.Read(ref queued) > beforeRelock, TimeSpan.FromSeconds(5)),
            "Locking the chart again did not re-enable scheduling.");
        Assert.True(macros.TryDequeue(out ExcelAction? relocked));
        relocked!();
        Assert.Equal(2, Volatile.Read(ref rendered));

        // The cell toggle is not a suspension: the user unlocking cells is choosing to type, and the
        // automatic refresh is exactly what should consume that edit.
        int beforeCells = Volatile.Read(ref queued);
        preferences.CellsUnlocked = true;
        scheduler.ScheduleCommandRefresh();
        Assert.True(
            SpinWait.SpinUntil(() => Volatile.Read(ref queued) > beforeCells, TimeSpan.FromSeconds(5)),
            "Unlocking the cells suspended the automatic refresh; it must not.");
    }

    /// <summary>
    /// R5.20: the persisted switch suspends the RENDER for both scheduling paths, and turning it
    /// back on re-enables them. The ribbon projection is never suspended, so an edit while the
    /// switch is off still queues one projection (owner ruling 2026-11-27). Mirrors the chart-unlock
    /// fact beside it, with the same positive control first.
    /// </summary>
    [Fact]
    public void An_off_switch_suspends_the_render_but_not_the_projection()
    {
        var autoRefresh = new AutoRefreshPreferences();
        int queued = 0;
        int rendered = 0;
        int projected = 0;
        ConcurrentQueue<ExcelAction> macros = new();
        Mock<IExcelApplicationAdapter> adapter = new();
        Action<string?>? onChanged = null;
        _ = adapter
            .Setup(a => a.SubscribeSheetChanged(It.IsAny<Action<string?>>()))
            .Callback<Action<string?>>(handler => onChanged = handler)
            .Returns(Mock.Of<IDisposable>());
        Mock<ISheetChangeAnalyzer> analyzer = new();
        _ = analyzer.Setup(a => a.ShouldAutoRefresh(It.IsAny<string?>())).Returns(true);
        using GanttRefreshScheduler scheduler = new(
            adapter.Object,
            analyzer.Object,
            macro =>
            {
                macros.Enqueue(macro);
                _ = Interlocked.Increment(ref queued);
            },
            () => Interlocked.Increment(ref rendered),
            preferences: null,
            autoRefresh,
            stateRefresh: () => Interlocked.Increment(ref projected));
        scheduler.Activate();

        scheduler.ScheduleCommandRefresh();
        Assert.True(
            SpinWait.SpinUntil(() => Volatile.Read(ref queued) == 1, TimeSpan.FromSeconds(5)),
            "With the switch on, a command did not schedule a refresh.");
        Assert.True(macros.TryDequeue(out ExcelAction? first));
        first!();
        Assert.Equal(1, Volatile.Read(ref rendered));

        autoRefresh.AutoRefreshEnabled = false;
        int before = Volatile.Read(ref queued);
        scheduler.ScheduleCommandRefresh();
        onChanged!("\"book|Sheet\"!D5");

        Assert.True(
            SpinWait.SpinUntil(() => Volatile.Read(ref queued) > before, TimeSpan.FromSeconds(5)),
            "The switch-off edit queued nothing; the projection must still run.");
        Assert.True(macros.TryDequeue(out ExcelAction? projection));
        projection!();
        Assert.Equal(1, Volatile.Read(ref rendered));
        Assert.Equal(1, Volatile.Read(ref projected));

        autoRefresh.AutoRefreshEnabled = true;
        int beforeOn = Volatile.Read(ref queued);
        scheduler.ScheduleCommandRefresh();
        Assert.True(
            SpinWait.SpinUntil(() => Volatile.Read(ref queued) > beforeOn, TimeSpan.FromSeconds(5)),
            "Turning the switch back on did not re-enable scheduling.");
        Assert.True(macros.TryDequeue(out ExcelAction? backOn));
        backOn!();
        Assert.Equal(2, Volatile.Read(ref rendered));
    }

    /// <summary>
    /// R5.20: the two suspensions compose. An off switch keeps scheduling quiet even when the
    /// chart is locked, and unlocking the chart does not resurrect it.
    /// </summary>
    [Fact]
    public void The_switch_and_the_chart_lock_compose_independently()
    {
        var preferences = new ProtectionPreferences();
        var autoRefresh = new AutoRefreshPreferences { AutoRefreshEnabled = false };
        int queued = 0;
        using GanttRefreshScheduler scheduler = new(
            Mock.Of<IExcelApplicationAdapter>(),
            Mock.Of<ISheetChangeAnalyzer>(),
            _ => Interlocked.Increment(ref queued),
            () => { },
            preferences,
            autoRefresh);

        scheduler.ScheduleCommandRefresh();
        Assert.False(
            SpinWait.SpinUntil(() => Volatile.Read(ref queued) > 0, TimeSpan.FromMilliseconds(400)),
            "An off switch queued a refresh with the chart locked.");

        preferences.ChartLocked = false;
        scheduler.ScheduleCommandRefresh();
        Assert.Equal(0, Volatile.Read(ref queued));

        autoRefresh.AutoRefreshEnabled = true;
        scheduler.ScheduleCommandRefresh();
        Assert.False(
            SpinWait.SpinUntil(() => Volatile.Read(ref queued) > 0, TimeSpan.FromMilliseconds(400)),
            "An unlocked chart scheduled a refresh once the switch was on; the lock must still suspend.");

        preferences.ChartLocked = true;
        scheduler.ScheduleCommandRefresh();
        Assert.True(
            SpinWait.SpinUntil(() => Volatile.Read(ref queued) > 0, TimeSpan.FromSeconds(5)),
            "Both suspensions lifted and the command still did not schedule.");
    }

    /// <summary>
    /// R5.23: a burst that touched the Type column heals before the projection, and the heal
    /// runs while the render is suspended. A burst that did not touch Type projects only.
    /// </summary>
    [Fact]
    public void A_type_change_heals_before_the_projection_while_suspended()
    {
        var autoRefresh = new AutoRefreshPreferences();
        var order = new List<string>();
        ConcurrentQueue<ExcelAction> macros = new();
        Mock<IExcelApplicationAdapter> adapter = new();
        Action<string?>? onChanged = null;
        _ = adapter
            .Setup(a => a.SubscribeSheetChanged(It.IsAny<Action<string?>>()))
            .Callback<Action<string?>>(handler => onChanged = handler)
            .Returns(Mock.Of<IDisposable>());
        Mock<ISheetChangeAnalyzer> analyzer = new();
        _ = analyzer.Setup(a => a.ShouldAutoRefresh(It.IsAny<string?>())).Returns(true);
        _ = analyzer
            .Setup(a => a.IsTypeColumnChange(It.IsAny<string?>()))
            .Returns((string? address) =>
                address is not null && address.EndsWith("C5", StringComparison.Ordinal));
        using GanttRefreshScheduler scheduler = new(
            adapter.Object,
            analyzer.Object,
            macro => macros.Enqueue(macro),
            () => order.Add("render"),
            preferences: null,
            autoRefresh,
            stateRefresh: () => order.Add("project"),
            typeHeal: () => order.Add("heal"));
        scheduler.Activate();
        autoRefresh.AutoRefreshEnabled = false;

        onChanged!("\"book|Sheet\"!C5");
        Assert.True(
            SpinWait.SpinUntil(() => !macros.IsEmpty, TimeSpan.FromSeconds(5)),
            "The Type burst queued nothing.");
        Assert.True(macros.TryDequeue(out ExcelAction? typeMacro));
        typeMacro!();

        Assert.Equal(["heal", "project"], order);

        onChanged!("\"book|Sheet\"!D5");
        Assert.True(
            SpinWait.SpinUntil(() => !macros.IsEmpty, TimeSpan.FromSeconds(5)),
            "The non-Type burst queued nothing.");
        Assert.True(macros.TryDequeue(out ExcelAction? plainMacro));
        plainMacro!();

        Assert.Equal(["heal", "project", "project"], order);
    }

    /// <summary>
    /// R5.23: a Type burst still heals when no projection is wired. The guard that skips a
    /// suspended burst exists for the projection; the heal is not a projection.
    /// </summary>
    [Fact]
    public void A_type_change_heals_when_no_projection_is_wired()
    {
        var autoRefresh = new AutoRefreshPreferences { AutoRefreshEnabled = false };
        int healed = 0;
        ConcurrentQueue<ExcelAction> macros = new();
        Mock<IExcelApplicationAdapter> adapter = new();
        Action<string?>? onChanged = null;
        _ = adapter
            .Setup(a => a.SubscribeSheetChanged(It.IsAny<Action<string?>>()))
            .Callback<Action<string?>>(handler => onChanged = handler)
            .Returns(Mock.Of<IDisposable>());
        Mock<ISheetChangeAnalyzer> analyzer = new();
        _ = analyzer.Setup(a => a.ShouldAutoRefresh(It.IsAny<string?>())).Returns(true);
        _ = analyzer.Setup(a => a.IsTypeColumnChange(It.IsAny<string?>())).Returns(true);
        using GanttRefreshScheduler scheduler = new(
            adapter.Object,
            analyzer.Object,
            macros.Enqueue,
            () => { },
            preferences: null,
            autoRefresh,
            trace: null,
            stateRefresh: null,
            typeHeal: () => healed++);
        scheduler.Activate();

        Assert.NotNull(onChanged);
        onChanged!("\"book|Sheet\"!C5");
        Assert.True(
            SpinWait.SpinUntil(() => !macros.IsEmpty, TimeSpan.FromSeconds(5)),
            "The Type burst queued nothing.");
        Assert.True(macros.TryDequeue(out ExcelAction? macro));
        macro!();

        Assert.Equal(1, healed);
    }

    /// <summary>
    /// R5.20: the production constructor must hand the process-wide switch to the pipeline, or the
    /// ribbon's toggle writes a value the scheduler never reads and every other test still passes.
    /// </summary>
    [Fact]
    public void The_production_constructor_tracks_the_shared_switch()
    {
        using GanttRefreshScheduler scheduler = new(
            Mock.Of<IExcelApplicationAdapter>(),
            Mock.Of<ISheetChangeAnalyzer>());

        // Identity only: the singleton's VALUE is session state other test classes mutate, so
        // asserting it here would make this fact depend on collection ordering.
        Assert.Same(AutoRefreshPreferences.Shared, scheduler.GetAutoRefreshSource());
    }

    /// <summary>
    /// The first-change window is shorter than the burst window, and shorter than nothing else's
    /// ordering violates. Pinned as a relationship rather than an absolute figure so tuning either
    /// window stays deliberate.
    /// </summary>
    [Fact]
    public void The_single_cell_window_is_shorter_than_the_burst_window()
    {
        Assert.True(GanttRefreshScheduler.SingleCellDebounce > TimeSpan.Zero);
        Assert.True(GanttRefreshScheduler.SingleCellDebounce < GanttRefreshScheduler.CellDebounce);
        Assert.True(GanttRefreshScheduler.CommandDebounce < GanttRefreshScheduler.SingleCellDebounce);
    }

    /// <summary>
    /// A lone cell edit asks for the short window.
    /// </summary>
    /// <remarks>
    /// The assertion is on the due time the scheduler REQUESTED, not on when a callback arrived. The
    /// first version polled wall-clock arrival against a 250 ms deadline and failed once in
    /// `verify-quick` while the machine was also running the analyzer and script gates: a correct
    /// 100 ms timer arrived at 279 ms under load. A test of a timer policy must not depend on the
    /// scheduler thread being prompt.
    /// </remarks>
    [Fact]
    public void A_single_cell_change_is_queued_on_the_short_window()
    {
        Action<string?>? onChanged = null;
        List<int> requested = [];
        Mock<IExcelApplicationAdapter> adapter = new();
        _ = adapter
            .Setup(a => a.SubscribeSheetChanged(It.IsAny<Action<string?>>()))
            .Callback<Action<string?>>(handler => onChanged = handler)
            .Returns(Mock.Of<IDisposable>());
        Mock<ISheetChangeAnalyzer> analyzer = new();
        _ = analyzer.Setup(a => a.ShouldAutoRefresh(It.IsAny<string?>())).Returns(true);
        using GanttRefreshScheduler scheduler = new(
            adapter.Object,
            analyzer.Object,
            _ => { },
            () => { },
            preferences: null,
            autoRefresh: null,
            trace: null,
            timerFactory: (milliseconds, _) =>
            {
                requested.Add(milliseconds);
                return new System.Threading.Timer(_ => { }, null, Timeout.Infinite, Timeout.Infinite);
            });
        scheduler.Activate();

        Assert.NotNull(onChanged);
        onChanged!("\"book|Sheet\"!D5");

        int only = Assert.Single(requested);
        Assert.Equal((int)GanttRefreshScheduler.SingleCellDebounce.TotalMilliseconds, only);
    }

    /// <summary>
    /// The second change in the same burst escalates to the burst window.
    /// </summary>
    /// <remarks>
    /// This is the half that keeps a paste coalescing into one render. Both changes are raised
    /// back to back, so the short timer cannot have fired between them.
    /// </remarks>
    [Fact]
    public void A_second_change_in_the_same_burst_escalates_to_the_burst_window()
    {
        Action<string?>? onChanged = null;
        List<int> requested = [];
        Mock<IExcelApplicationAdapter> adapter = new();
        _ = adapter
            .Setup(a => a.SubscribeSheetChanged(It.IsAny<Action<string?>>()))
            .Callback<Action<string?>>(handler => onChanged = handler)
            .Returns(Mock.Of<IDisposable>());
        Mock<ISheetChangeAnalyzer> analyzer = new();
        _ = analyzer.Setup(a => a.ShouldAutoRefresh(It.IsAny<string?>())).Returns(true);
        using GanttRefreshScheduler scheduler = new(
            adapter.Object,
            analyzer.Object,
            _ => { },
            () => { },
            preferences: null,
            autoRefresh: null,
            trace: null,
            timerFactory: (milliseconds, _) =>
            {
                requested.Add(milliseconds);
                return new System.Threading.Timer(_ => { }, null, Timeout.Infinite, Timeout.Infinite);
            });
        scheduler.Activate();

        Assert.NotNull(onChanged);
        onChanged!("\"book|Sheet\"!D5");
        onChanged!("\"book|Sheet\"!D6");

        // Two requests, in order: the short window first, then the burst window that replaces it.
        // This is the load-bearing half - a scheduler that ignored the escalation would ask for the
        // short window twice and a paste would render several times.
        Assert.Equal(2, requested.Count);
        Assert.Equal((int)GanttRefreshScheduler.SingleCellDebounce.TotalMilliseconds, requested[0]);
        Assert.Equal((int)GanttRefreshScheduler.CellDebounce.TotalMilliseconds, requested[1]);
    }

    /// <summary>
    /// A cell burst reports the analyzer cost it paid before the debounce and the target that was
    /// scheduled, because neither is visible in the pipeline's own timings.
    /// </summary>
    [Fact]
    public void A_cell_burst_traces_detect_cost_and_the_scheduled_target()
    {
        Action<string?>? onChanged = null;
        List<string> lines = [];
        ConcurrentQueue<ExcelAction> macros = new();
        int analyzerCalls = 0;
        Mock<IExcelApplicationAdapter> adapter = new();
        _ = adapter
            .Setup(a => a.SubscribeSheetChanged(It.IsAny<Action<string?>>()))
            .Callback<Action<string?>>(handler => onChanged = handler)
            .Returns(Mock.Of<IDisposable>());
        Mock<ISheetChangeAnalyzer> analyzer = new();
        _ = analyzer
            .Setup(a => a.ShouldAutoRefresh(It.IsAny<string?>()))
            .Callback(() => analyzerCalls++)
            .Returns(true);
        using GanttRefreshScheduler scheduler = new(
            adapter.Object,
            analyzer.Object,
            macros.Enqueue,
            () => { },
            preferences: null,
            autoRefresh: null,
            trace: line => lines.Add(line));
        scheduler.Activate();

        Assert.NotNull(onChanged);
        onChanged!("\"Book|Sheet1\"!D5");
        onChanged!("\"Book|Sheet1\"!D6");

        Assert.True(
            SpinWait.SpinUntil(() => macros.Count == 1, TimeSpan.FromSeconds(5)),
            "The burst did not queue a pass.");
        Assert.True(macros.TryDequeue(out ExcelAction? macro));
        macro!();

        string line = Assert.Single(lines);
        Assert.Contains("gantt.refresh detect source=cell", line, StringComparison.Ordinal);
        Assert.Contains("events=2", line, StringComparison.Ordinal);
        Assert.Contains("detectms=", line, StringComparison.Ordinal);
        Assert.Contains("queuedms=", line, StringComparison.Ordinal);
        Assert.Contains("scheduled=\"Book|Sheet1\"", line, StringComparison.Ordinal);
        Assert.Equal(2, analyzerCalls);
    }

    /// <summary>
    /// A command pass reports zero analysed events and no cell target, so a reader can tell the two
    /// scheduling paths apart in the log.
    /// </summary>
    [Fact]
    public void A_command_refresh_traces_zero_detect_events()
    {
        List<string> lines = [];
        ConcurrentQueue<ExcelAction> macros = new();
        using GanttRefreshScheduler scheduler = new(
            Mock.Of<IExcelApplicationAdapter>(),
            Mock.Of<ISheetChangeAnalyzer>(),
            macros.Enqueue,
            () => { },
            preferences: null,
            autoRefresh: null,
            trace: line => lines.Add(line));

        scheduler.ScheduleCommandRefresh();
        Assert.True(
            SpinWait.SpinUntil(() => macros.Count == 1, TimeSpan.FromSeconds(5)),
            "The command did not queue a pass.");
        Assert.True(macros.TryDequeue(out ExcelAction? macro));
        macro!();

        string line = Assert.Single(lines);
        Assert.Contains("gantt.refresh detect source=command events=0", line, StringComparison.Ordinal);
        Assert.Contains("detectms=0 ", line, StringComparison.Ordinal);
        Assert.Contains("scheduled=\"\"", line, StringComparison.Ordinal);
    }

    private static GanttRefreshScheduler CreateScheduler(
        Action<ExcelAction> queue,
        Action refresh,
        Action? stateRefresh = null) =>
        new(
            Mock.Of<IExcelApplicationAdapter>(),
            Mock.Of<ISheetChangeAnalyzer>(),
            queue,
            refresh,
            stateRefresh: stateRefresh);
}
