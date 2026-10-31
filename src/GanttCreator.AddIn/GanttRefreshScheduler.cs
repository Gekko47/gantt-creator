#pragma warning disable IDE0048, IDE0047 // Timer constructor expression: analyzer rules conflict on parentheses in this file.
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Schedules a debounced automatic Gantt refresh when the user edits
/// <c>Id</c> or <c>Type</c> cells in <c>tblGanttData</c>.
/// </summary>
/// <remarks>
/// <para>
/// Excel's <c>SheetChange</c> event fires on every cell edit. The scheduler
/// consults <see cref="ISheetChangeAnalyzer.ShouldAutoRefresh"/> to decide
/// whether the change warrants an automatic refresh: <c>Id</c>/<c>Type</c>
/// changes do; <c>Start</c>/<c>Finish</c> require a manual Refresh; and
/// <c>Duration</c> is engine-driven. Changes are debounced (~300 ms) so
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
/// <b>Change tracking.</b> Each <c>SheetChange</c> that warrants auto-refresh
/// records the changed row on <see cref="RefreshSheetCommand._sharedTracker"/>
/// so the next auto-refresh can target only affected rows. The cell value is
/// not read here (old value is unavailable because <c>SheetChange</c> fires
/// after the edit); only the row number is recorded, which is sufficient for
/// change-aware validation and highlighting gating.
/// </para>
/// <para>
/// All Excel COM calls run on the Excel main STA thread via
/// <see cref="ExcelDna.Integration.ExcelAsyncUtil"/>. The timer callback
/// itself runs on a thread-pool thread and is marshalled accordingly.
/// </para>
/// </remarks>
internal sealed class GanttRefreshScheduler(
    IExcelApplicationAdapter adapter,
    ISheetChangeAnalyzer analyzer) : IDisposable
{
    private const int _debounceMilliseconds = 300;

    private readonly IExcelApplicationAdapter _adapter = adapter;
    private readonly ISheetChangeAnalyzer _analyzer = analyzer;
    private System.Threading.Timer? _timer;
    private int _timerGeneration;
    private IDisposable? _sheetChangeSubscription;
    private bool _disposed;

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
        DisposeTimer();
        _sheetChangeSubscription?.Dispose();
        _sheetChangeSubscription = null;
    }

    /// <summary>
    /// Handles a worksheet cell change: if the changed address warrants an
    /// automatic refresh, records the changed row on the shared tracker and
    /// resets the debounce timer.
    /// </summary>
    /// <param name="changedAddress">The Excel-style address of the changed range.</param>
    private void OnSheetChanged(string? changedAddress)
    {
        if (!_analyzer.ShouldAutoRefresh(changedAddress))
        {
            return;
        }

#pragma warning disable CA1031 // Best-effort tracking: catch all exceptions to never block a refresh on tracking failure.
        try
        {
            RecordChange(changedAddress);
        }
        catch (Exception)
        {
            // Best-effort tracking; never block a refresh on tracking failure.
        }
#pragma warning restore CA1031

        ResetTimer();
    }

    /// <summary>
    /// Parses an Excel-style address (e.g. <c>"$K$5"</c>, <c>"K$5"</c>,
    /// <c>"Sheet1!$K$5"</c>) and records the changed row on the shared
    /// tracker so the next auto-refresh can target only affected rows.
    /// </summary>
    /// <param name="changedAddress">The Excel-style address of the changed range.</param>
    private static void RecordChange(string? changedAddress)
    {
        if (string.IsNullOrEmpty(changedAddress))
        {
            return;
        }

        // Strip sheet name prefix if present (everything before '!').
        string address = changedAddress;
        int bang = address.IndexOf('!', StringComparison.Ordinal);
        if (bang >= 0)
        {
            address = address[(bang + 1)..];
        }

        // Remove '$' absolute-reference markers.
        address = address.Replace('$', '\0');

        // Split into column letters and row number.
        int i = 0;
        while (i < address.Length && char.IsLetter(address[i]))
        {
            i++;
        }

        if (i == 0 || i == address.Length)
        {
            // Not a recognizable cell address.
            return;
        }

        string columnLetters = address[..i];
        string rowPart = address[i..];

        if (!int.TryParse(rowPart, out int rowNumber))
        {
            return;
        }

        // Convert column letters to 1-based index: A=1, B=2, ..., Z=26, AA=27, ...
        int column = 0;
        foreach (char c in columnLetters)
        {
            column = column * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
        }

        // Old value is null because SheetChange fires after the edit.
        RefreshSheetCommand.RecordChange(rowNumber, column, null);
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
        ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(
            RefreshSheetCommand.RunSilentForExcel);
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
