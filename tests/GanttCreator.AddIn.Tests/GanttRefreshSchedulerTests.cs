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
        scheduler.ScheduleRefresh();
        scheduler.ScheduleRefresh();

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
        scheduler.ScheduleRefresh();

        Assert.True(queuedSignal.Wait(TimeSpan.FromSeconds(5)), "The debounced macro was not queued before the deadline.");
        scheduler.CancelPendingRefresh();
        Assert.NotNull(queuedMacro);
        queuedMacro!();

        Assert.Equal(0, Volatile.Read(ref refreshCount));
    }

    private static GanttRefreshScheduler CreateScheduler(Action<ExcelAction> queue, Action refresh) =>
        new(
            Mock.Of<IExcelApplicationAdapter>(),
            Mock.Of<ISheetChangeAnalyzer>(),
            queue,
            refresh);
}
