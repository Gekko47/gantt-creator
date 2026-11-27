using System.Diagnostics;
using System.Globalization;
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
/// and presentation changes do; <c>Duration</c> is engine-driven. Changes are debounced so
/// rapid multi-cell edits produce a single refresh rather than one per edit: the first
/// accepted change of a burst waits 100 ms and the second escalates to 300 ms.
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
    /// <summary>The coalescing window for a burst of worksheet cell edits.</summary>
    /// <remarks>
    /// A user pasting, or dragging a fill, raises its events together (the owner's trace shows
    /// nine inside 3.4 ms), so the window is long enough to collapse a burst into one render.
    /// </remarks>
    private const int _cellDebounceMilliseconds = 300;

    /// <summary>The coalescing window for the FIRST accepted cell change of a burst.</summary>
    /// <remarks>
    /// A lone cell edit pays the whole burst window for no coalescing benefit, and the 2026-11-20
    /// trace made that window the largest single term in click-to-plot (300 ms of debounce against
    /// a 10-115 ms macro queue and a 164-1072 ms pipeline). The escalation below is what keeps
    /// paste coalescing: it only has to win the gap between a paste's events, measured at ~0.4 ms,
    /// so 100 ms is ample margin for a slow host delivering an event late.
    /// </remarks>
    private const int _singleCellDebounceMilliseconds = 100;

    /// <summary>The coalescing window for a successful command.</summary>
    /// <remarks>
    /// A command has already finished its own writes before it schedules, so the only
    /// events left to absorb are the ones its write emitted — not a user burst. Waiting
    /// the full cell window added 300ms of dead time between a button click and the
    /// plot, which is exactly the latency the owner reported.
    /// </remarks>
    private const int _commandDebounceMilliseconds = 50;

    /// <summary>The cell-edit coalescing window, exposed for the policy pin.</summary>
    internal static readonly TimeSpan CellDebounce = TimeSpan.FromMilliseconds(_cellDebounceMilliseconds);

    /// <summary>The first-change coalescing window, exposed for the policy pin.</summary>
    internal static readonly TimeSpan SingleCellDebounce =
        TimeSpan.FromMilliseconds(_singleCellDebounceMilliseconds);

    /// <summary>The post-command coalescing window, exposed for the policy pin.</summary>
    internal static readonly TimeSpan CommandDebounce = TimeSpan.FromMilliseconds(_commandDebounceMilliseconds);

    private readonly IExcelApplicationAdapter _adapter;
    private readonly ISheetChangeAnalyzer _analyzer;
    private readonly Action<ExcelDna.Integration.ExcelAction> _queueAsMacro;
    private readonly Action _refresh;
    private System.Threading.Timer? _timer;
    private int _timerGeneration;

    /// <summary>
    /// How many accepted cell changes the current burst has seen, so the first one can take the
    /// short window and every later one the long window. Reset when a burst's timer fires.
    /// </summary>
    private int _cellBurstLength;

    /// <summary>Where a burst's <c>gantt.refresh detect</c> line goes, or null to discard it.</summary>
    private readonly Action<string>? _trace;

    /// <summary>
    /// Creates the debounce timer, or <see langword="null"/> for the real host timer.
    /// </summary>
    /// <remarks>
    /// A test seam, and the reason the two window facts are deterministic: they assert the due time
    /// the scheduler ASKED for, so a loaded machine cannot make a 100 ms window look like a 300 ms
    /// one. An earlier version measured wall-clock arrival against a 250 ms deadline and failed in
    /// `verify-quick` while the box was also running the analyzer and script gates.
    /// </remarks>
    private readonly Func<int, Action, System.Threading.Timer>? _timerFactory;

    /// <summary>Accepted cell changes in the current burst, reported by the detect trace.</summary>
    private int _burstEvents;

    /// <summary>Analyzer <see cref="Stopwatch"/> ticks spent in the current burst.</summary>
    private long _burstAnalyzerTicks;

    /// <summary>The <c>book|sheet</c> target of the burst's last accepted change.</summary>
    private string _burstTarget = string.Empty;

    /// <summary>Whether the pending pass was scheduled by a cell edit or by a command.</summary>
    private string _pendingSource = "cell";
    private IDisposable? _sheetChangeSubscription;
    private bool _disposed;

    /// <summary>
    /// Whether a successful command has a refresh queued, so the cell events that
    /// command's own writes emit are folded into it instead of restarting the longer
    /// cell window.
    /// </summary>
    private volatile bool _commandRefreshPending;

    /// <summary>
    /// Republishes the ribbon snapshot when a burst's render is suspended, or
    /// <see langword="null"/> when nothing tracks it.
    /// </summary>
    /// <remarks>
    /// The ribbon is a projection of the table, not of the render, so it is never suspended
    /// (owner ruling 2026-11-27). This is the action the debounce runs when the chart must stay
    /// still; the pass republishes through its own postlude when it is allowed to run.
    /// </remarks>
    private readonly Action? _stateRefresh;

    internal GanttRefreshScheduler(IExcelApplicationAdapter adapter, ISheetChangeAnalyzer analyzer)
        : this(
            adapter,
            analyzer,
            QueueMacro,
            RefreshSheetCommand.RunSilentForExcel,
            ProtectionPreferences.Shared,
            AutoRefreshPreferences.Shared,
            CommandBoundary.Instance.WriteTechnicalRecord,
            stateRefresh: () => RibbonStateService.Instance.RefreshAfterTableChange())
    {
    }

    internal GanttRefreshScheduler(
        IExcelApplicationAdapter adapter,
        ISheetChangeAnalyzer analyzer,
        Action<ExcelDna.Integration.ExcelAction> queueAsMacro,
        Action refresh,
        IProtectionPreferenceSource? preferences = null,
        IAutoRefreshPreferenceSource? autoRefresh = null,
        Action<string>? trace = null,
        Func<int, Action, System.Threading.Timer>? timerFactory = null,
        Action? stateRefresh = null)
    {
        _adapter = adapter ?? throw new ArgumentNullException(nameof(adapter));
        _analyzer = analyzer ?? throw new ArgumentNullException(nameof(analyzer));
        _queueAsMacro = queueAsMacro ?? throw new ArgumentNullException(nameof(queueAsMacro));
        _refresh = refresh ?? throw new ArgumentNullException(nameof(refresh));
        _preferences = preferences;
        _autoRefresh = autoRefresh;
        _trace = trace;
        _timerFactory = timerFactory;
        _stateRefresh = stateRefresh;
    }

    /// <summary>
    /// The CHART toggle, or <see langword="null"/> when nothing tracks it (ADR-0055 clause 7).
    /// </summary>
    /// <remarks>
    /// While the chart is unlocked an automatic refresh would rebuild the chart underneath a manual
    /// drag, so the RENDER stays quiet. The ribbon projection is not a rebuild and is never
    /// suspended. An explicit Refresh is unaffected: that is the user asking for the rebuild.
    /// Unlocking the sheet's CELLS is deliberately not a suspension — a cell edit is the input an
    /// automatic refresh exists to consume, and suppressing it would leave the chart stale for as
    /// long as the user keeps cells unlocked.
    /// </remarks>
    private readonly IProtectionPreferenceSource? _preferences;

    /// <summary>
    /// The user's persisted automatic-refresh switch, or <see langword="null"/> when nothing tracks
    /// it (R5.20). Null means enabled, so every existing test and every caller that does not care
    /// about the switch keeps the behaviour it had.
    /// </summary>
    private readonly IAutoRefreshPreferenceSource? _autoRefresh;

    /// <summary>Gets whether automatic refresh is suspended by an unlocked chart.</summary>
    private bool IsSuspendedByUnlock => _preferences?.ChartLocked == false;

    /// <summary>
    /// Gets whether an automatic render may be scheduled at all.
    /// </summary>
    /// <remarks>
    /// Two independent reasons suspend the render, and they compose: an unlocked chart is a manual
    /// edit in progress, and the persisted switch is the user asking for the explicit Refresh to be
    /// the only rebuild. Neither suspends the ribbon projection, which is a read of the table and a
    /// UI invalidation. An explicit Refresh is never affected by this predicate — it does not come
    /// through this class.
    /// </remarks>
    private bool IsSuspended => IsSuspendedByUnlock || _autoRefresh?.AutoRefreshEnabled == false;

    /// <summary>
    /// Returns the switch this scheduler reads, for test verification only.
    /// </summary>
    /// <remarks>
    /// The production constructor must hand the process-wide
    /// <see cref="AutoRefreshPreferences.Shared"/> to the pipeline, or the ribbon's toggle writes
    /// a value nothing consults and every test still passes. That wiring fact is asserted through
    /// this accessor rather than by scheduling a real pass into Excel.
    /// </remarks>
    internal IAutoRefreshPreferenceSource? GetAutoRefreshSource() => _autoRefresh;

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
        _ = Interlocked.Increment(ref _timerGeneration);
        DisposeTimer();
        _sheetChangeSubscription?.Dispose();
        _sheetChangeSubscription = null;
    }

    internal void ScheduleCommandRefresh()
    {
        if (_disposed || IsSuspended)
        {
            return;
        }

        // The command's own cell writes raise SheetChange synchronously; without this
        // flag they would reset the timer to the longer cell window and delay (or
        // duplicate) the pass the command just asked for. Cleared when the queued macro
        // runs, so a genuine edit made while the refresh is in flight still schedules the
        // follow-up pass.
        _commandRefreshPending = true;
        Volatile.Write(ref _pendingSource, "command");
        ResetTimer(_commandDebounceMilliseconds);
    }

    internal void CancelPendingRefresh()
    {
        _ = Interlocked.Increment(ref _timerGeneration);
        _commandRefreshPending = false;
        DisposeTimer();
    }

    /// <summary>
    /// Handles a worksheet cell change: if the changed address warrants an
    /// automatic refresh and resets the debounce timer.
    /// </summary>
    /// <param name="changedAddress">The Excel-style address of the changed range.</param>
    private void OnSheetChanged(string? changedAddress)
    {
        if (_commandRefreshPending)
        {
            // This is the command's own write, already covered by its pending refresh.
            return;
        }

        // The render suspension is NOT consulted here. The ribbon snapshot is a projection of
        // the table and is never suspended, so an accepted change still has to reach it while
        // the chart stays put. Which action the burst runs is chosen when the debounce fires,
        // because the switch or the lock can change inside the window.

        // The analyzer cost is paid per changed cell, before the debounce, so it is invisible in the
        // pipeline's own timings. Accumulate it for the burst and report it with the pass.
        long analyzerStarted = Stopwatch.GetTimestamp();
        bool chartAffecting = _analyzer.ShouldAutoRefresh(changedAddress);
        _ = Interlocked.Add(ref _burstAnalyzerTicks, Stopwatch.GetTimestamp() - analyzerStarted);
        if (!chartAffecting)
        {
            return;
        }

        _ = Interlocked.Increment(ref _burstEvents);
        Volatile.Write(ref _pendingSource, "cell");
        string target = TargetOf(changedAddress);
        if (target.Length > 0)
        {
            Volatile.Write(ref _burstTarget, target);
        }

        // The first accepted change of a burst gets the short window; a second change proves the
        // burst is a paste or a fill and takes the long one.
        int burstLength = Interlocked.Increment(ref _cellBurstLength);
        ResetTimer(burstLength == 1 ? _singleCellDebounceMilliseconds : _cellDebounceMilliseconds);
    }

    /// <summary>
    /// Returns the <c>book|sheet</c> part of a changed address, or an empty string when the address
    /// carries none.
    /// </summary>
    /// <param name="changedAddress">The address the sheet-change event reported.</param>
    /// <returns>The owner prefix, trimmed of the quoting Excel adds, or an empty string.</returns>
    internal static string TargetOf(string? changedAddress)
    {
        if (string.IsNullOrEmpty(changedAddress))
        {
            return string.Empty;
        }

        int bang = changedAddress.LastIndexOf('!');
        return bang <= 0 ? string.Empty : changedAddress[..bang].Trim('"', '\'');
    }

    /// <summary>
    /// Records the debounce segment of one pass: how many cell changes were analysed, what they
    /// cost, how long the macro wait was, and which workbook/sheet the schedule targeted.
    /// </summary>
    /// <param name="source">Whether a cell edit or a command scheduled the pass.</param>
    /// <param name="events">Accepted cell changes in the burst.</param>
    /// <param name="detectMilliseconds">Analyzer time the burst paid before the debounce.</param>
    /// <param name="queuedMilliseconds">Timer fire to macro entry.</param>
    /// <param name="target">The <c>book|sheet</c> captured at schedule time, or an empty string.</param>
    private void TraceDetect(
        string source,
        int events,
        double detectMilliseconds,
        double queuedMilliseconds,
        string target) =>
        _trace?.Invoke(string.Create(
            CultureInfo.InvariantCulture,
            $"gantt.refresh detect source={source} events={events} detectms={detectMilliseconds:0.#} queuedms={queuedMilliseconds:0.#} scheduled=\"{target}\""));

    /// <summary>
    /// Resets (or starts) the debounce timer. Each change restarts the window, so rapid
    /// multi-cell edits coalesce into one refresh; the window is the caller's, because a
    /// burst's later changes wait longer than its first.
    /// The generation counter is incremented so stale timer callbacks from
    /// superseded timers can be detected and discarded.
    /// </summary>
    /// <param name="debounceMilliseconds">The window this change waits for.</param>
    private void ResetTimer(int debounceMilliseconds)
    {
        DisposeTimer();
        int generation = Interlocked.Increment(ref _timerGeneration);
        if (_timerFactory is not null)
        {
            _timer = _timerFactory(debounceMilliseconds, () => OnTimerElapsed(generation));
            return;
        }

        // A local function keeps the callback shape unambiguous, so the two
        // parenthesis style rules never disagree about this expression.
        void HandleDebounceElapsed(object? state)
        {
            _ = state;
            OnTimerElapsed(generation);
        }

        _timer = new System.Threading.Timer(
            HandleDebounceElapsed,
            null,
            debounceMilliseconds,
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

        // This burst is over, so the next accepted change starts a new one and takes the short
        // window again. Reset before the guards below, so a pass suppressed while suspended does
        // not leave the next genuine edit using the burst window.
        Volatile.Write(ref _cellBurstLength, 0);
        int burstEvents = Interlocked.Exchange(ref _burstEvents, 0);
        long analyzerTicks = Interlocked.Exchange(ref _burstAnalyzerTicks, 0);
        string source = Volatile.Read(ref _pendingSource);
        string target = Volatile.Read(ref _burstTarget);

        // Re-check at the last moment: the chart can be unlocked, or the switch turned off,
        // between the debounce and this callback. That suspends the RENDER only. The ribbon is a
        // projection of the table and is never suspended, so the burst still reaches it.
        bool suspended = IsSuspended;
        if (suspended && _stateRefresh is null)
        {
            // Nothing tracks the projection in this composition (a harness, or a host that never
            // wired the state service), so a suspended burst keeps the old no-op behaviour.
            return;
        }

        long firedTicks = Stopwatch.GetTimestamp();

        // A command path never consults the analyzer, so it reports zero. Carrying a stale
        // accumulator here made a command line claim analyzer cost it never paid (observed
        // 2026-11-20: `source=command detectms=0.4`), which a reader would attribute to the pass.
        double analyzerMilliseconds = string.Equals(source, "command", StringComparison.Ordinal)
            ? 0d
            : analyzerTicks * 1000d / Stopwatch.Frequency;

        // The detect line names the action this burst produced: "state" when the render was
        // suspended, the original source otherwise.
        string traceSource = suspended ? "state" : source;

        // Marshal to the Excel main STA thread before touching COM.
        _queueAsMacro(() =>
        {
            if (!_disposed && generation == Volatile.Read(ref _timerGeneration))
            {
                // Emitted before the pipeline's own lines so a reader can attribute the wait that
                // preceded them: the debounce, the analyzer cost and the macro-queue turn.
                TraceDetect(
                    traceSource,
                    burstEvents,
                    analyzerMilliseconds,
                    Stopwatch.GetElapsedTime(firedTicks).TotalMilliseconds,
                    target);

                // Clear before the action so an edit made during it still schedules the next
                // burst rather than being swallowed as the command's own write.
                _commandRefreshPending = false;

                if (suspended)
                {
                    _stateRefresh?.Invoke();
                    return;
                }

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
