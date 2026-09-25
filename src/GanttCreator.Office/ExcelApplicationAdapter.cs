// Namespace alias rather than a wholesale using: Microsoft.Office.Interop.Excel
// declares its own Action type, which would collide with System.Action under a
// plain `using Microsoft.Office.Interop.Excel;`.
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="IExcelApplicationAdapter"/> over the Excel application object
/// supplied by the host at add-in load.
/// </summary>
/// <param name="application">
/// The Excel application object (for example <c>ExcelDnaUtil.Application</c>), or
/// <see langword="null"/> when the host supplied none (unit tests, non-Excel
/// host). A foreign object fails the interface casts and degrades to "not
/// determinable" with no event connection.
/// </param>
/// <remarks>
/// <para>
/// COM ownership: the Excel <c>Application</c> object and the <c>Workbook</c>
/// objects reachable from <c>ActiveWorkbook</c> are Excel-owned shared roots.
/// This adapter takes no ownership of them, never calls
/// <c>FinalReleaseComObject</c>, and force-releases nothing. The one resource it
/// owns is its event connection, which is detached exactly once.
/// </para>
/// <para>
/// Every member degrades instead of throwing: the adapter is reached from
/// Ribbon state callbacks and from Excel application events, where a
/// propagating exception surfaces as a host error dialog.
/// </para>
/// </remarks>
public sealed class ExcelApplicationAdapter(object? application) : IExcelApplicationAdapter
{
    private readonly Excel.Application? _application = application as Excel.Application;

    // Events are bound through the AppEvents_Event interface: `NewWorkbook`
    // would otherwise be ambiguous with _Application.NewWorkbook (the method
    // that creates a workbook).
    private readonly Excel.AppEvents_Event? _events = application as Excel.AppEvents_Event;

    /// <inheritdoc />
    public bool? HasActiveWorkbook()
    {
        Excel.Application? application = _application;
        if (application is null)
        {
            return null;
        }

        // CA1031: this read runs on the Ribbon state path and inside Excel
        // application-event handlers, where a propagating COM failure surfaces
        // as a host error. A failed probe reports "not determinable" so the
        // state service keeps its previous snapshot.
#pragma warning disable CA1031
        try
        {
            Excel.Workbook? workbook = application.ActiveWorkbook;
            return workbook is not null;
        }
        catch
        {
            return null;
        }
#pragma warning restore CA1031
    }

    /// <inheritdoc />
    public IDisposable SubscribeWorkbookStateChanged(Action handler)
    {
        ArgumentNullException.ThrowIfNull(handler);

        Excel.AppEvents_Event? events = _events;
        if (events is null)
        {
            return InertSubscription.Instance;
        }

        var subscription = new WorkbookStateSubscription(events, handler);
        subscription.Attach();
        return subscription;
    }

    /// <summary>
    /// The one owned resource: the event connection. It adds exactly one handler
    /// to each of the three workbook-state events and removes exactly those
    /// instances. Callbacks are disabled by the first dispose; a removal that
    /// failed keeps its delegate so a later <see cref="Dispose"/> can retry it.
    /// </summary>
    /// <param name="events">The Excel application event interface (a shared root; not owned).</param>
    /// <param name="handler">The subscriber's change handler.</param>
    private sealed class WorkbookStateSubscription(Excel.AppEvents_Event events, Action handler)
        : IDisposable, IWorkbookStateSubscriptionStatus
    {
        private Excel.AppEvents_WorkbookActivateEventHandler? _onWorkbookActivate;
        private Excel.AppEvents_WorkbookDeactivateEventHandler? _onWorkbookDeactivate;
        private Excel.AppEvents_NewWorkbookEventHandler? _onNewWorkbook;
        private bool _disposed;
        private EventHandlerState _workbookActivateState = EventHandlerState.Detached;
        private EventHandlerState _workbookDeactivateState = EventHandlerState.Detached;
        private EventHandlerState _newWorkbookState = EventHandlerState.Detached;

        /// <inheritdoc />
        public EventHandlerState WorkbookActivateState => _workbookActivateState;

        /// <inheritdoc />
        public EventHandlerState WorkbookDeactivateState => _workbookDeactivateState;

        /// <inheritdoc />
        public EventHandlerState NewWorkbookState => _newWorkbookState;

        /// <inheritdoc />
        public bool AllDetached =>
            _workbookActivateState == EventHandlerState.Detached
            && _workbookDeactivateState == EventHandlerState.Detached
            && _newWorkbookState == EventHandlerState.Detached;

        /// <summary>
        /// Attaches the three handlers. A failure degrades to "no automatic
        /// invalidation" and detaches whatever was already attached, so the
        /// subscription stays inert rather than half-connected.
        /// </summary>
        internal void Attach()
        {
            // CA1031: attaching to a live Excel object can fail (busy host,
            // object in teardown, refused COM connection). The Ribbon state
            // service still works with command-driven invalidation, so this
            // degrades instead of throwing into AutoOpen.
#pragma warning disable CA1031
            try
            {
                _onWorkbookActivate = _ => Invoke();
                _onWorkbookDeactivate = _ => Invoke();
                _onNewWorkbook = _ => Invoke();

                events.WorkbookActivate += _onWorkbookActivate;
                _workbookActivateState = EventHandlerState.Attached;
                events.WorkbookDeactivate += _onWorkbookDeactivate;
                _workbookDeactivateState = EventHandlerState.Attached;
                events.NewWorkbook += _onNewWorkbook;
                _newWorkbookState = EventHandlerState.Attached;
            }
            catch
            {
                Dispose();
            }
#pragma warning restore CA1031
        }

        /// <summary>
        /// Invokes the subscriber's handler. The production handler is
        /// non-throwing by contract; this guard is defence in depth so a
        /// misbehaving handler cannot escape into Excel's event dispatch.
        /// </summary>
        private void Invoke()
        {
            if (_disposed)
            {
                return;
            }

            // CA1031: see the attach rationale above — the notification degrades
            // to none rather than surfacing as a host error dialog.
#pragma warning disable CA1031
            try
            {
                handler();
            }
            catch
            {
                // Intentionally empty: the change notification is skipped.
            }
#pragma warning restore CA1031
        }

        /// <inheritdoc />
        public void Dispose()
        {
            // Callbacks stay disabled after the first disposal (the handler
            // guard reads _disposed), but teardown itself is re-enterable so a
            // later Dispose retries a removal that previously failed.
            _disposed = true;

            // CA1031: teardown runs from AutoClose and must never throw into
            // Excel. Each removal is attempted independently so a failure on one
            // event does not prevent the remaining removals, and a failed removal
            // is recorded as DetachFailed rather than reported as a success.
            Detach(
                ref _workbookActivateState,
                () =>
                {
                    if (_onWorkbookActivate is not null)
                    {
                        events.WorkbookActivate -= _onWorkbookActivate;
                    }
                });

            Detach(
                ref _workbookDeactivateState,
                () =>
                {
                    if (_onWorkbookDeactivate is not null)
                    {
                        events.WorkbookDeactivate -= _onWorkbookDeactivate;
                    }
                });

            Detach(
                ref _newWorkbookState,
                () =>
                {
                    if (_onNewWorkbook is not null)
                    {
                        events.NewWorkbook -= _onNewWorkbook;
                    }
                });

            // Drop only the delegates whose handler is now detached. A delegate
            // whose removal failed is retained so a later Dispose can retry the
            // removal; the recorded DetachFailed state stays the honest signal
            // that the event source may still hold the handler.
            Release(ref _onWorkbookActivate, _workbookActivateState);
            Release(ref _onWorkbookDeactivate, _workbookDeactivateState);
            Release(ref _onNewWorkbook, _newWorkbookState);
        }

        /// <summary>
        /// Clears a delegate reference once its handler is detached, so the
        /// handler is not kept alive after teardown. A delegate that is still
        /// attached or whose removal failed is retained for a retry.
        /// </summary>
        /// <typeparam name="TDelegate">The COM event delegate type.</typeparam>
        /// <param name="handler">The delegate reference, cleared in place.</param>
        /// <param name="state">The handler's recorded state.</param>
        private static void Release<TDelegate>(ref TDelegate? handler, EventHandlerState state)
            where TDelegate : Delegate
        {
            if (state == EventHandlerState.Detached)
            {
                handler = null;
            }
        }

        /// <summary>
        /// Removes one handler and records the outcome. A handler that was never
        /// attached stays <see cref="EventHandlerState.Detached"/>; a removal that
        /// throws becomes <see cref="EventHandlerState.DetachFailed"/> and never
        /// throws out of <see cref="Dispose"/>.
        /// </summary>
        /// <param name="state">The handler's state, updated in place.</param>
        /// <param name="remove">The removal action.</param>
        private static void Detach(ref EventHandlerState state, Action remove)
        {
            if (state == EventHandlerState.Detached)
            {
                return;
            }

            state = EventHandlerState.Detaching;

            // CA1031: see the teardown rationale in Dispose.
#pragma warning disable CA1031
            try
            {
                remove();
                state = EventHandlerState.Detached;
            }
            catch
            {
                // The removal genuinely failed; reporting Detached here would make
                // a leaked event connection look like a clean teardown.
                state = EventHandlerState.DetachFailed;
            }
#pragma warning restore CA1031
        }
    }

    /// <summary>
    /// The subscription returned when no application object is available. It is
    /// already inert, so callers need no special case.
    /// </summary>
    private sealed class InertSubscription : IDisposable, IWorkbookStateSubscriptionStatus
    {
        private InertSubscription()
        {
        }

        /// <summary>Gets the shared instance; the type holds no state.</summary>
        internal static InertSubscription Instance { get; } = new();

        /// <inheritdoc />
        public EventHandlerState WorkbookActivateState => EventHandlerState.Detached;

        /// <inheritdoc />
        public EventHandlerState WorkbookDeactivateState => EventHandlerState.Detached;

        /// <inheritdoc />
        public EventHandlerState NewWorkbookState => EventHandlerState.Detached;

        /// <inheritdoc />
        public bool AllDetached => true;

        /// <inheritdoc />
        public void Dispose()
        {
            // Intentionally empty: nothing was attached.
        }
    }
}
