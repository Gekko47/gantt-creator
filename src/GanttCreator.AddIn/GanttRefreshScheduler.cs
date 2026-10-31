#pragma warning disable IDE0048, IDE0047 // Timer constructor expression: analyzer rules conflict on parentheses in this file.
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Schedules a debounced automatic Gantt refresh when the user edits
/// chart-affecting cells in <c>tblGanttData</c>.
/// </summary>
/// <remarks>
/// <para>
/// Excel's <c>SheetChange</c> event fires on every cell edit. The scheduler
/// consults <see cref="ISheetChangeAnalyzer.ShouldAutoRefresh"/> to decide
/// whether the change warrants an automatic refresh: chart-affecting data
/// and presentation changes do; <c>Duration</c> is engine-driven. Changes are debounced (~300 ms) so
/// rapid multi-cell edits produce a single refresh rather than one per edit.
/// </para>
/// <para>
/// The automatic refresh is silent: it writes progress to the status bar and
/// never shows a modal dialog, so it never interrupts the user's workflow.
/// Blocking errors are still reported through the status bar.
/// </para>
/// <para>
/// <b>Timer safety.</b> The Dispose method on System.Threading.Timer does
/// not cancel already-queued callbacks, so rapid edits can cause a stale
/// timer callback to fire after a newer timer has been created. A
/// <see cref="_timerGeneration"/> counter is incremented on every
/// <see cref="ResetTimer"/> and checked in <see cref="OnTimerElapsed"/>;
/// callbacks from a superseded generation exit without queuing a refresh,
/// preventing double rendering.
/// </para>
/// <para>
/// <b>Change tracking.</b> The safe refresh validates and builds the complete
/// table before reconciliation, so events schedule a debounced full refresh.
/// </para>
/// <para>
/// All Excel COM calls run on the Excel main STA thread via
/// <see cref="ExcelDna.Integration.ExcelAsyncUtil"/>. The timer callback
/// itself runs on a thread-pool thread and is marshalled accordingly.
/// </para>
/// </remarks>
internal sealed class GanttRefreshScheduler : IDisposable
{
    private const int _debounceMilliseconds = 300;

    private readonly IExcelApplicationAdapter _adapter;
    private readonly ISheetChangeAnalyzer _analyzer;
    private readonly Action<ExcelDna.Integration.ExcelAction> _queueAsMacro;
    private readonly Action _refresh;
    private System.Threading.Timer? _timer;
    private int _timerGeneration;
    private IDisposable? _sheetChangeSubscription;
    private bool _disposed;

    internal GanttRefreshScheduler(IExcelApplicationAdapter adapter, ISheetChangeAnalyzer analyzer)
        : this(adapter, analyzer, QueueMacro, RefreshSheetCommand.RunSilentForExcel)
    {
    }

    internal GanttRefreshScheduler(
        IExcelApplicationAdapter adapter,
        ISheetChangeAnalyzer analyzer,
        Action<ExcelDna.Integration.ExcelAction> queueAsMacro,
        Action refresh)
    {
        _adapter = adapter ?? throw new ArgumentNullException(nameof(adapter));
        _analyzer = analyzer ?? throw new ArgumentNullException(nameof(analyzer));
        _queueAsMacro = queueAsMacro ?? throw new ArgumentNullException(nameof(queueAsMacro));
        _refresh = refresh ?? throw new ArgumentNullException(nameof(refresh));
    }

    private static void QueueMacro(ExcelDna.Integration.ExcelAction macro) =>
        ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(macro);

    /// <summary>
    /// Activates the scheduler: subscribes to worksheet cell changes.
    /// Safe to call multiple times; each call replaces the previous
    /// subscription.
    /// </summary>
    public void Activate()
    {
        if (_disposed)
        {
            return;
        }

        Deactivate();

        _sheetChangeSubscription = _adapter.SubscribeSheetChanged(OnSheetChanged);
    }

    /// <summary>
    /// Deactivates the scheduler: detaches the sheet-change subscription and
    /// disposes any pending timer.
    /// </summary>
    public void Deactivate()
    {
        _ = System.Threading.Interlocked.Increment(ref _timerGeneration);
        DisposeTimer();
        _sheetChangeSubscription?.Dispose();
        _sheetChangeSubscription = null;
    }

    internal void ScheduleRefresh()
    {
        if (!_disposed)
        {
            ResetTimer();
        }
    }

    internal void CancelPendingRefresh()
    {
        _ = System.Threading.Interlocked.Increment(ref _timerGeneration);
        DisposeTimer();
    }

    /// <summary>
    /// Handles a worksheet cell change: if the changed address warrants an
    /// automatic refresh and resets the debounce timer.
    /// </summary>
    /// <param name="changedAddress">The Excel-style address of the changed range.</param>
    private void OnSheetChanged(string? changedAddress)
    {
        if (!_analyzer.ShouldAutoRefresh(changedAddress))
        {
            return;
        }

        ResetTimer();
    }

    /// <summary>
    /// Resets (or starts) the debounce timer. Each change restarts the
    /// ~300 ms window, so rapid multi-cell edits coalesce into one refresh.
    /// The generation counter is incremented so stale timer callbacks from
    /// superseded timers can be detected and discarded.
    /// </summary>
    private void ResetTimer()
    {
        DisposeTimer();
        int generation = System.Threading.Interlocked.Increment(ref _timerGeneration);
        _timer = new System.Threading.Timer(
            _ => OnTimerElapsed(generation),
            null,
            _debounceMilliseconds,
            Timeout.Infinite);
    }

    /// <summary>
    /// Timer elapsed: runs the silent refresh on the Excel main thread.
    /// Callbacks from a superseded timer generation are discarded to prevent
    /// double rendering when <see cref="ResetTimer"/> is called before the
    /// previous timer's callback has fired.
    /// </summary>
    /// <param name="generation">The timer generation that created this callback.</param>
    private void OnTimerElapsed(int generation)
    {
        // Discard stale callbacks from superseded timer generations.
        if (generation != _timerGeneration)
        {
            return;
        }

        // Marshal to the Excel main STA thread before touching COM.
        _queueAsMacro(() =>
        {
            if (!_disposed && generation == System.Threading.Volatile.Read(ref _timerGeneration))
            {
                _refresh();
            }
        });
    }

    /// <summary>
    /// Disposes the debounce timer if one is pending.
    /// </summary>
    private void DisposeTimer()
    {
        _timer?.Dispose();
        _timer = null;
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        Deactivate();
    }
}
