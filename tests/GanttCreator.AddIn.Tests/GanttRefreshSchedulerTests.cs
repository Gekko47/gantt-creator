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
    /// ADR-0055 clause 6: while a shape is unlocked for manual editing, neither a cell edit nor a
    /// command may schedule a refresh, because either would rebuild the chart over the edit. Locking
    /// the shapes again re-enables both.
    /// </summary>
    [Fact]
    public void An_unlocked_shape_suspends_both_scheduling_paths()
    {
        var unlockState = new ShapeUnlockState();
        int queued = 0;
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
            _ => Interlocked.Increment(ref queued),
            () => { },
            unlockState);
        scheduler.Activate();

        // Positive control first: with nothing unlocked both paths queue, so the negative assertion
        // below cannot pass on a harness that simply never queues anything.
        scheduler.ScheduleCommandRefresh();
        Assert.True(
            SpinWait.SpinUntil(() => Volatile.Read(ref queued) == 1, TimeSpan.FromSeconds(5)),
            "With nothing unlocked, a command did not schedule a refresh.");

        _ = unlockState.Add("G-0123456789abcdef0123456789abcdef:bar");
        int before = Volatile.Read(ref queued);
        scheduler.ScheduleCommandRefresh();
        onChanged!("\"book|Sheet\"!D5");

        // A negative assertion needs a bound. The command window is 50 ms, so spinning for 400 ms is
        // longer than any timer could hide behind, and SpinUntil does not sleep the thread.
        Assert.False(
            SpinWait.SpinUntil(
                () => Volatile.Read(ref queued) > before,
                TimeSpan.FromMilliseconds(400)),
            "A suspended scheduler queued a refresh while a shape was unlocked.");

        _ = unlockState.Clear();
        scheduler.ScheduleCommandRefresh();
        Assert.True(
            SpinWait.SpinUntil(() => Volatile.Read(ref queued) > before, TimeSpan.FromSeconds(5)),
            "Locking the shapes again did not re-enable scheduling.");
    }

    private static GanttRefreshScheduler CreateScheduler(Action<ExcelAction> queue, Action refresh) =>
        new(
            Mock.Of<IExcelApplicationAdapter>(),
            Mock.Of<ISheetChangeAnalyzer>(),
            queue,
            refresh);
}
