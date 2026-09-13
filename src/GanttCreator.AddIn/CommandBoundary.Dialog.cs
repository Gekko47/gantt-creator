namespace GanttCreator.AddIn;

/// <summary>
/// The production error dialog for the command boundary: a managed WinForms
/// TaskDialog (Error icon, one OK button) with a MessageBox fallback.
/// Follows the ADR-0006 managed TaskDialog pattern — no hand-rolled comctl32
/// P/Invoke. The dialog has no hyperlink, so no access-key escaping is
/// needed. Callers must invoke on the Excel main STA thread (the RibbonX
/// onAction thread).
/// </summary>
internal static class CommandErrorDialog
{
    /// <summary>The dialog caption for every command failure.</summary>
    private const string _title = "Gantt Creator";

    /// <summary>The main instruction heading for every command failure.</summary>
    private const string _mainInstruction = "Gantt Creator could not complete the last action";

    /// <summary>
    /// Shows the error dialog. Never throws by contract: if the TaskDialog
    /// cannot be shown, a plain MessageBox fallback is attempted; the
    /// boundary additionally guards this method, so a double failure degrades
    /// to no dialog instead of propagating into Excel.
    /// </summary>
    /// <param name="message">The translated user-safe message.</param>
    internal static void Show(string message)
    {
        ArgumentNullException.ThrowIfNull(message);

        // CA1031: the dialog must always degrade — first to the MessageBox
        // fallback, then (via the boundary's own guard) to no dialog. A
        // failure here must never propagate into Excel.
#pragma warning disable CA1031
        try
        {
            _ = TaskDialog.ShowDialog(CreatePage(message));
        }
        catch
        {
            // The managed TaskDialog can fail outside the Excel main STA
            // thread or on a malformed page configuration. Fall back to a
            // plain MessageBox carrying the same information.
            _ = MessageBox.Show(
                message,
                _title,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Builds the error task-dialog page. Internal (the testable seam, as in
    /// <see cref="TaskDialogApi.CreatePage"/>) so contract tests can inspect
    /// the page without showing a dialog.
    /// </summary>
    /// <param name="message">The translated user-safe message.</param>
    /// <returns>The configured, unbound page.</returns>
    internal static TaskDialogPage CreatePage(string message)
    {
        ArgumentNullException.ThrowIfNull(message);

        var page = new TaskDialogPage
        {
            Caption = _title,
            Heading = _mainInstruction,
            Text = message,
            Icon = TaskDialogIcon.Error,
        };
        page.Buttons.Add(TaskDialogButton.OK);
        return page;
    }
}
