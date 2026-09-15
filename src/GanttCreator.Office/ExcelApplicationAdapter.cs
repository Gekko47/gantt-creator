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
    /// instances on the first dispose.
    /// </summary>
    /// <param name="events">The Excel application event interface (a shared root; not owned).</param>
    /// <param name="handler">The subscriber's change handler.</param>
    private sealed class WorkbookStateSubscription(Excel.AppEvents_Event events, Action handler) : IDisposable
    {
        private Excel.AppEvents_WorkbookActivateEventHandler? _onWorkbookActivate;
        private Excel.AppEvents_WorkbookDeactivateEventHandler? _onWorkbookDeactivate;
        private Excel.AppEvents_NewWorkbookEventHandler? _onNewWorkbook;
        private bool _disposed;

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
                events.WorkbookDeactivate += _onWorkbookDeactivate;
                events.NewWorkbook += _onNewWorkbook;
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
            if (_disposed)
            {
                return;
            }

            _disposed = true;

            // CA1031: teardown runs from AutoClose and must never throw into
            // Excel; a failed detach degrades to a retained event connection.
#pragma warning disable CA1031
            try
            {
                if (_onWorkbookActivate is not null)
                {
                    events.WorkbookActivate -= _onWorkbookActivate;
                }

                if (_onWorkbookDeactivate is not null)
                {
                    events.WorkbookDeactivate -= _onWorkbookDeactivate;
                }

                if (_onNewWorkbook is not null)
                {
                    events.NewWorkbook -= _onNewWorkbook;
                }
            }
            catch
            {
                // Intentionally empty: see the teardown rationale above.
            }
#pragma warning restore CA1031
            finally
            {
                // Drop the delegate references unconditionally: the connection
                // is not retried, and the delegates must not keep the handler
                // alive after teardown.
                _onWorkbookActivate = null;
                _onWorkbookDeactivate = null;
                _onNewWorkbook = null;
            }
        }
    }

    /// <summary>
    /// The subscription returned when no application object is available. It is
    /// already inert, so callers need no special case.
    /// </summary>
    private sealed class InertSubscription : IDisposable
    {
        private InertSubscription()
        {
        }

        /// <summary>Gets the shared instance; the type holds no state.</summary>
        internal static InertSubscription Instance { get; } = new();

        /// <inheritdoc />
        public void Dispose()
        {
            // Intentionally empty: nothing was attached.
        }
    }
}
