namespace GanttCreator.AddIn;

/// <summary>
/// Writes a message to Excel's status bar, non-modally.
/// </summary>
/// <remarks>
/// <para>
/// The status bar is where Excel itself reports a rejected edit, and -- unlike
/// a message box -- it needs no dismissal. That makes it the right home for
/// an ADVISORY warning raised from a ribbon callback: the user sees the
/// message and the Excel UI thread is never blocked waiting for an OK.
/// </para>
/// <para>
/// This matters for the headless test run as much as for the user. A modal
/// dialog on this path would block the test host forever, because there is
/// nobody to dismiss it, and the test would never reach its own assertion.
/// </para>
/// </remarks>
internal static class RibbonStatusBar
{
    /// <summary>
    /// The status-bar message shown when a warning is raised.
    /// </summary>
    /// <param name="message">The message to show.</param>
    internal static void Warn(string message) => RibbonStateService.WriteStatusBar(message);
}
