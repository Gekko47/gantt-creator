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
    /// Subscribes to worksheet cell changes and supplies an address qualified by
    /// workbook and worksheet name. Implementations without worksheet events return
    /// an inert subscription.
    /// </summary>
    /// <param name="handler">Receives a qualified address on Excel's main thread.</param>
    /// <returns>A subscription that detaches when disposed.</returns>
    IDisposable SubscribeSheetChanged(Action<string?> handler)
    {
        ArgumentNullException.ThrowIfNull(handler);
        return InertSheetChangeSubscription.Instance;
    }

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
    /// Writes a message to Excel's status bar, non-modally.
    /// </summary>
    /// <param name="message">The message to show.</param>
    /// <remarks>
    /// The status bar is where Excel itself reports a rejected edit, and --
    /// unlike a message box -- it needs no dismissal. That makes it the right
    /// home for an advisory warning raised from a Ribbon callback: a modal
    /// dialog there would hold the Excel UI thread until the user dismissed
    /// it, and would block a headless test run forever.
    /// <para>
    /// Never throws: a failure to write loses the message and nothing more.
    /// </para>
    /// </remarks>
    void ShowStatusBar(string message);

    /// <summary>Clears any add-in status-bar text so Excel resumes its automatic status.</summary>
    /// <remarks>Never throws: a failure to clear loses nothing.</remarks>
    void ClearStatusBar();

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

    /// <summary>
    /// Subscribes <paramref name="handler"/> to the application-level
    /// sheet-activation event, so entering the Gantt sheet can refresh the chart
    /// and restore the add-in tab (R5.16b, ADR-0058).
    /// </summary>
    /// <param name="handler">
    /// Invoked on Excel's main thread after a sheet becomes active, with no
    /// arguments: the subscriber re-reads the workbook, because only the read
    /// knows which worksheet is the Gantt sheet. The handler must not throw; the
    /// adapter guards the invocation so a throw cannot escape into Excel's
    /// dispatch.
    /// </param>
    /// <returns>
    /// A subscription that detaches exactly once when disposed; disposing it
    /// again is a no-op. Implementations without application events return an
    /// inert subscription, so callers never special-case them.
    /// </returns>
    /// <exception cref="ArgumentNullException"><paramref name="handler"/> is <see langword="null"/>.</exception>
    IDisposable SubscribeSheetActivated(Action handler)
    {
        ArgumentNullException.ThrowIfNull(handler);
        return InertSheetChangeSubscription.Instance;
    }
}

/// <summary>An inert sheet subscription for adapters without Excel events.</summary>
internal sealed class InertSheetChangeSubscription : IDisposable
{
    private InertSheetChangeSubscription()
    {
    }

    /// <summary>The shared stateless instance.</summary>
    internal static InertSheetChangeSubscription Instance { get; } = new();

    /// <inheritdoc />
    public void Dispose()
    {
        // No event handler was attached.
    }
}
