using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Owns one session's Excel event connections and the coalescing refresh scheduler.
/// </summary>
/// <remarks>
/// <para>
/// The handlers stay on <see cref="RibbonStateService"/>, because each one transitions state that
/// the service owns (the snapshot, the workbook-scoped facts, the persisted-switch seed). What
/// lives here is the lifetime: activation replaces every connection, so a second load cannot
/// double-subscribe, and every detach is guarded, so teardown never throws into Excel.
/// </para>
/// <para>
/// The scheduler is created here rather than at the call site because it is one of the
/// connections: it subscribes to sheet changes, it is cancelled by the same events, and it must
/// die with them. <c>ExcelDnaUtil.Application</c> is read only while subscribing, exactly as it
/// was before, so a host without Excel keeps an inert set of connections.
/// </para>
/// </remarks>
internal sealed class RibbonEventWiring
{
    private readonly IExcelApplicationAdapter? _applicationAdapter;
    private readonly ISelectionIdentifier? _selectionIdentifier;
    private readonly Action _onSelectionChanged;
    private readonly Action _onSheetActivated;
    private readonly Action _onWorkbookStateChanged;

    /// <summary>The selection-change subscription, detached exactly once.</summary>
    private IDisposable? _selectionSubscription;

    /// <summary>The workbook-state subscription, detached exactly once.</summary>
    private IDisposable? _workbookStateSubscription;

    /// <summary>The sheet-activation subscription (R5.16b), detached exactly once.</summary>
    private IDisposable? _sheetActivationSubscription;

    private GanttRefreshScheduler? _refreshScheduler;

    /// <summary>Creates the wiring over one activation's sources and handlers.</summary>
    /// <param name="applicationAdapter">The workbook-state and sheet-activation source, or null.</param>
    /// <param name="selectionIdentifier">The selection source (R5.5 ruling 6), or null.</param>
    /// <param name="onSelectionChanged">The service's selection handler.</param>
    /// <param name="onSheetActivated">The service's sheet-activation handler.</param>
    /// <param name="onWorkbookStateChanged">The service's workbook-state handler.</param>
    /// <exception cref="ArgumentNullException">A handler is <see langword="null"/>.</exception>
    internal RibbonEventWiring(
        IExcelApplicationAdapter? applicationAdapter,
        ISelectionIdentifier? selectionIdentifier,
        Action onSelectionChanged,
        Action onSheetActivated,
        Action onWorkbookStateChanged)
    {
        _onSelectionChanged = onSelectionChanged ?? throw new ArgumentNullException(nameof(onSelectionChanged));
        _onSheetActivated = onSheetActivated ?? throw new ArgumentNullException(nameof(onSheetActivated));
        _onWorkbookStateChanged = onWorkbookStateChanged ?? throw new ArgumentNullException(nameof(onWorkbookStateChanged));
        _applicationAdapter = applicationAdapter;
        _selectionIdentifier = selectionIdentifier;
    }

    /// <summary>Schedules the coalesced automatic refresh, when the scheduler was created.</summary>
    internal void ScheduleCommandRefresh() => _refreshScheduler?.ScheduleCommandRefresh();

    /// <summary>Cancels a pending automatic refresh, when the scheduler was created.</summary>
    internal void CancelPendingRefresh() => _refreshScheduler?.CancelPendingRefresh();

    /// <summary>
    /// Attaches every connection, replacing any previous one so a re-activation cannot
    /// double-subscribe. Never throws.
    /// </summary>
    internal void Subscribe()
    {
        // Detach first, in any order: each attach below replaces the field, and the original
        // detach-immediately-before-attach sequence differs only in which connection is dead
        // for a few instructions inside one activation.
        DetachWorkbookState();
        DisposeRefreshScheduler();
        DetachSheetActivation();
        DetachSelection();

        AttachWorkbookState();
        CreateScheduler();
        AttachSheetActivation();
        AttachSelection();
    }

    /// <summary>
    /// Detaches every connection and disposes the scheduler, each exactly once. Never throws.
    /// </summary>
    /// <remarks>
    /// Named rather than implementing <see cref="IDisposable"/>: the object's lifetime is the
    /// session, and its single owner is <see cref="RibbonStateService.Teardown"/>. A disposable
    /// field there would make the singleton disposable for no caller's benefit.
    /// </remarks>
    internal void DetachAll()
    {
        DisposeRefreshScheduler();
        DetachSheetActivation();
        DetachSelection();
        DetachWorkbookState();
    }

    private void AttachWorkbookState()
    {
        if (_applicationAdapter is not { } adapter)
        {
            return;
        }

        _workbookStateSubscription = TrySubscribe(
            () => adapter.SubscribeWorkbookStateChanged(_onWorkbookStateChanged));
    }

    private void AttachSheetActivation()
    {
        if (_applicationAdapter is not { } adapter)
        {
            return;
        }

        _sheetActivationSubscription = TrySubscribe(
            () => adapter.SubscribeSheetActivated(_onSheetActivated));
    }

    private void AttachSelection()
    {
        if (_selectionIdentifier is not { } identifier)
        {
            return;
        }

        _selectionSubscription = TrySubscribe(
            () => identifier.SubscribeSelectionChanged(_onSelectionChanged));
    }

    /// <summary>Creates and activates the sheet-change scheduler. Never throws.</summary>
    /// <remarks>
    /// The application object is read here because the analyzer needs it; a host without Excel
    /// makes the construction fail, and the wiring then simply has no scheduler — the same
    /// degradation the service had when this code lived inline.
    /// </remarks>
    private void CreateScheduler()
    {
        if (_applicationAdapter is not { } adapter)
        {
            return;
        }

        // CA1031: activation runs inside the never-throwing AutoOpen; a failure to create the
        // scheduler degrades to command-driven refresh only.
#pragma warning disable CA1031
        try
        {
            object? application = ExcelDna.Integration.ExcelDnaUtil.Application;
            GanttRefreshScheduler scheduler = new(adapter, new ExcelSheetChangeAnalyzer(application));
            scheduler.Activate();
            _refreshScheduler = scheduler;
        }
        catch
        {
            _refreshScheduler = null;
        }
#pragma warning restore CA1031
    }

    /// <summary>Attaches one subscription, reporting an inert connection instead of throwing.</summary>
    /// <param name="subscribe">The subscribe call.</param>
    /// <returns>The subscription, or <see langword="null"/> when subscribing failed.</returns>
    /// <remarks>
    /// CA1031: activation runs inside the never-throwing AutoOpen; a failed subscribe degrades to
    /// command-driven invalidation only, which is a working add-in minus one live update.
    /// </remarks>
    private static IDisposable? TrySubscribe(Func<IDisposable> subscribe)
    {
        // CA1031: see the remarks; the caller cannot act on a failure and must not fail the load.
#pragma warning disable CA1031
        try
        {
            return subscribe();
        }
        catch
        {
            return null;
        }
#pragma warning restore CA1031
    }

    private void DetachSelection() => Detach(ref _selectionSubscription);

    private void DetachWorkbookState() => Detach(ref _workbookStateSubscription);

    private void DetachSheetActivation() => Detach(ref _sheetActivationSubscription);

    /// <summary>Detaches one subscription exactly once. Never throws.</summary>
    /// <param name="subscription">The field holding the subscription; cleared first.</param>
    private static void Detach(ref IDisposable? subscription)
    {
        IDisposable? current = subscription;
        subscription = null;
        if (current is null)
        {
            return;
        }

        // CA1031: teardown must never throw into Excel; a failed detach degrades to a retained
        // event connection.
#pragma warning disable CA1031
        try
        {
            current.Dispose();
        }
        catch
        {
            // Intentionally empty: see the teardown rationale above.
        }
#pragma warning restore CA1031
    }

    private void DisposeRefreshScheduler()
    {
        GanttRefreshScheduler? scheduler = _refreshScheduler;
        _refreshScheduler = null;
        if (scheduler is null)
        {
            return;
        }

        // CA1031: AutoClose must not surface a teardown failure into Excel.
#pragma warning disable CA1031
        try
        {
            scheduler.Dispose();
        }
        catch
        {
            // Intentionally empty: see the teardown rationale above.
        }
#pragma warning restore CA1031
    }
}
