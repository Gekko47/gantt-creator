namespace GanttCreator.Office;

/// <summary>
/// The state of one Excel event handler within a subscription. Tracked per
/// handler so a failed removal is never reported as a successful detach.
/// </summary>
public enum EventHandlerState
{
    /// <summary>The handler is registered with the Excel event source.</summary>
    Attached = 0,

    /// <summary>Removal is in progress.</summary>
    Detaching = 1,

    /// <summary>The handler is no longer registered.</summary>
    Detached = 2,

    /// <summary>Removal was attempted and failed; the handler may still be registered.</summary>
    DetachFailed = 3,
}

/// <summary>
/// Observes the per-handler detach state of a workbook-state subscription.
/// The port still returns <see cref="IDisposable"/>; callers that need to
/// confirm teardown cast to this interface.
/// </summary>
public interface IWorkbookStateSubscriptionStatus
{
    /// <summary>Gets the state of the WorkbookActivate handler.</summary>
    EventHandlerState WorkbookActivateState { get; }

    /// <summary>Gets the state of the WorkbookDeactivate handler.</summary>
    EventHandlerState WorkbookDeactivateState { get; }

    /// <summary>Gets the state of the NewWorkbook handler.</summary>
    EventHandlerState NewWorkbookState { get; }

    /// <summary>Gets whether every handler reached <see cref="EventHandlerState.Detached"/>.</summary>
    bool AllDetached { get; }
}

/// <summary>
/// Narrow port over the Excel application facts the Ribbon state service needs:
/// whether a workbook is active, and notification when application-level
/// workbook state changes.
/// </summary>
/// <remarks>
/// The surface is deliberately primitive — no interop type crosses the port —
/// so the AddIn compilation never names <c>Microsoft.Office.Interop.Excel</c>.
/// Both <c>ExcelDna.Integration</c> and the Office PIA declare a public
/// <c>Application</c> type, which is a CS0433 (duplicate type) hazard in any
/// compilation referencing both, and the AddIn references both.
/// </remarks>
public interface IExcelApplicationAdapter
{
    /// <summary>
    /// Gets whether Excel currently has an active workbook, or <see langword="null"/>
    /// when the fact cannot be determined.
    /// </summary>
    /// <returns>
    /// <see langword="true"/> when an active workbook exists;
    /// <see langword="false"/> when Excel is running with no active workbook;
    /// <see langword="null"/> when the fact is not determinable (no application
    /// object was supplied, the object was not an Excel application, or the
    /// probe failed). The caller keeps its previous value for
    /// <see langword="null"/>, so a transient failure cannot disable dynamic
    /// Ribbon controls. Never throws.
    /// </returns>
    bool? HasActiveWorkbook();

    /// <summary>
    /// Subscribes <paramref name="handler"/> to application-level workbook state
    /// changes: a workbook becoming active, a workbook being deactivated, and a
    /// new workbook being created.
    /// </summary>
    /// <param name="handler">
    /// Invoked for each change, on the thread that raised the Excel event (the
    /// Excel main thread). The handler must not throw; the adapter guards the
    /// invocation so a throw cannot escape into Excel's event dispatch.
    /// </param>
    /// <returns>
    /// A subscription that detaches exactly once when disposed; disposing it
    /// again is a no-op. When no application object is available the returned
    /// subscription is already inert, so callers never special-case it.
    /// </returns>
    /// <exception cref="ArgumentNullException"><paramref name="handler"/> is <see langword="null"/>.</exception>
    IDisposable SubscribeWorkbookStateChanged(Action handler);
}
