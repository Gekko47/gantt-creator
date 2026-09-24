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

#pragma warning disable CA1031
        try
        {
            _ = TaskDialog.ShowDialog(CreatePage(message));
        }
        catch
        {
            _ = MessageBox.Show(message, _title, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
#pragma warning restore CA1031
    }

    /// <summary>Shows a yes/no confirmation and returns true only for Yes.</summary>
    /// <param name="message">The count-only confirmation message.</param>
    /// <returns>Whether the user selected Yes.</returns>
    internal static bool Confirm(string message)
    {
        ArgumentNullException.ThrowIfNull(message);

#pragma warning disable CA1031
        try
        {
            TaskDialogButton result = TaskDialog.ShowDialog(CreateConfirmationPage(message));
            return result == TaskDialogButton.Yes;
        }
        catch
        {
            return MessageBox.Show(
                message,
                _title,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) == DialogResult.Yes;
        }
#pragma warning restore CA1031
    }

    /// <summary>Builds the confirmation task-dialog page.</summary>
    /// <param name="message">The count-only confirmation message.</param>
    /// <returns>The configured page.</returns>
    internal static TaskDialogPage CreateConfirmationPage(string message)
    {
        ArgumentNullException.ThrowIfNull(message);
        var page = new TaskDialogPage
        {
            Caption = _title,
            Heading = "Repair Gantt Creator configuration?",
            Text = message,
            Icon = TaskDialogIcon.Information,
        };
        page.Buttons.Add(TaskDialogButton.Yes);
        page.Buttons.Add(TaskDialogButton.No);
        return page;
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
