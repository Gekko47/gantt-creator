namespace GanttCreator.AddIn;

/// <summary>
/// Managed wrapper for the Windows Task Dialog, used to display diagnostic
/// information with a clickable hyperlink to the active log file.
/// </summary>
/// <remarks>
/// <para>
/// Implemented with the WinForms <see cref="TaskDialog"/> API (.NET 5+) rather
/// than hand-rolled comctl32 P/Invoke. Root cause (2026-09-13): the previous
/// P/Invoke declared <c>GetDesktopWindow</c> against comctl32.dll (it is
/// exported by user32.dll), and comctl32 v6 exports <c>TaskDialogIndirect</c>
/// only by ordinal (121) — never by name — so every ribbon click threw
/// <see cref="System.EntryPointNotFoundException"/> before any dialog could
/// show. The managed API activates comctl32 v6 itself and removes both
/// entry-point resolution problems.
/// </para>
/// <para>
/// Callers must invoke <see cref="ShowWithHyperlink"/> on the Excel main STA
/// thread (the RibbonX onAction thread). The wrapper never opens the dialog
/// itself in tests; <see cref="CreatePage"/> is the testable seam.
/// </para>
/// </remarks>
internal static class TaskDialogApi
{
    /// <summary>
    /// Builds the diagnostics task-dialog page: title, main instruction,
    /// hyperlink-enabled content, and a Close button. The <paramref
    /// name="openLinkAction"/> receives the clicked link's <c>href</c> value.
    /// </summary>
    /// <param name="title">The dialog caption.</param>
    /// <param name="mainInstruction">The main instruction heading.</param>
    /// <param name="content">
    /// The dialog content; may contain <c>&lt;a href="..."&gt;text&lt;/a&gt;</c>
    /// link markup, which requires <see cref="TaskDialogPage.EnableLinks"/>.
    /// </param>
    /// <param name="openLinkAction">
    /// Action invoked with the clicked link's href target.
    /// </param>
    /// <returns>The configured, unbound page.</returns>
    internal static TaskDialogPage CreatePage(
        string title,
        string mainInstruction,
        string content,
        Action<string> openLinkAction)
    {
        ArgumentNullException.ThrowIfNull(title);
        ArgumentNullException.ThrowIfNull(mainInstruction);
        ArgumentNullException.ThrowIfNull(content);
        ArgumentNullException.ThrowIfNull(openLinkAction);

        var page = new TaskDialogPage
        {
            Caption = title,
            Heading = mainInstruction,
            Text = content,
            Icon = TaskDialogIcon.Information,
            EnableLinks = true,
        };
        page.Buttons.Add(TaskDialogButton.Close);

        // The Task Dialog never executes links itself; execution must be
        // handled in the LinkClicked event. OpenLogFile is failure-proof by
        // contract (it catches everything), so no additional guard is needed.
        page.LinkClicked += (_, e) => HandleLinkClicked(e, openLinkAction);

        return page;
    }

    /// <summary>
    /// Shows the diagnostics TaskDialog and returns a simple button result:
    /// 1 when the user closed the dialog via the Close button, 0 otherwise.
    /// </summary>
    /// <param name="title">The dialog caption.</param>
    /// <param name="mainInstruction">The main instruction heading.</param>
    /// <param name="content">The dialog content; may contain link markup.</param>
    /// <param name="openFileAction">
    /// Action invoked with the clicked hyperlink target (the log file path).
    /// </param>
    /// <returns>1 when closed via the Close button; otherwise 0.</returns>
    internal static int ShowWithHyperlink(
        string title,
        string mainInstruction,
        string content,
        Action<string> openFileAction)
    {
        TaskDialogPage page = CreatePage(title, mainInstruction, content, openFileAction);

        TaskDialogButton button = TaskDialog.ShowDialog(page);
        return ReferenceEquals(button, TaskDialogButton.Close) ? 1 : 0;
    }

    /// <summary>
    /// Delivers a link-click event to the open action. Split out from
    /// <see cref="CreatePage"/> so the delivery contract can be tested
    /// without showing a dialog.
    /// </summary>
    /// <param name="e">The link-click event data.</param>
    /// <param name="openLinkAction">Action invoked with the link href.</param>
    internal static void HandleLinkClicked(TaskDialogLinkClickedEventArgs e, Action<string> openLinkAction)
    {
        ArgumentNullException.ThrowIfNull(e);
        ArgumentNullException.ThrowIfNull(openLinkAction);
        openLinkAction(e.LinkHref);
    }
}
