using GanttCreator.Core;

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
    /// The label on the destructive-command confirmation's
    /// verification checkbox (ADR-0042).
    /// </summary>
    private const string _dontAskAgainText = "Don't ask again this session";

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

    /// <summary>
    /// Shows a yes/no confirmation and returns true only for Yes,
    /// with no opt-out checkbox. The heading is the shared
    /// destructive-command main instruction
    /// (<see cref="DestructiveCommandPolicy.ConfirmationMainInstruction"/>).
    /// </summary>
    /// <param name="message">The confirmation body.</param>
    /// <returns>Whether the user selected Yes.</returns>
    internal static bool Confirm(string message) =>
        Confirm(message, session: null);

    /// <summary>
    /// Shows a yes/no confirmation and returns true only for Yes. The
    /// heading is the shared destructive-command main instruction
    /// (<see cref="DestructiveCommandPolicy.ConfirmationMainInstruction"/>),
    /// constant across every destructive command class. When
    /// <paramref name="session"/> is supplied, the dialog also offers
    /// a "don't ask again this session" verification checkbox
    /// (ADR-0042); confirming with it checked records the opt-out on
    /// the session so later delete-row commands in the same Excel
    /// session proceed without the dialog.
    /// </summary>
    /// <param name="message">The confirmation body.</param>
    /// <param name="session">
    /// The session whose opt-out the checkbox drives, or null for a
    /// plain confirmation with no opt-out checkbox.
    /// </param>
    /// <returns>Whether the user selected Yes.</returns>
    internal static bool Confirm(string message, ConfirmationSession? session)
    {
        ArgumentNullException.ThrowIfNull(message);

#pragma warning disable CA1031
        try
        {
            TaskDialogPage page = CreateConfirmationPage(message, session);
            TaskDialogButton result = TaskDialog.ShowDialog(page);
            var confirmed = result == TaskDialogButton.Yes;
            // The checkbox state lives on the page; record it on the
            // session only for a confirmed change, so a declined
            // dialog cannot opt the user out.
            session?.RecordChoice(confirmed, page.Verification?.Checked ?? false);
            return confirmed;
        }
        catch
        {
            // MessageBox fallback: a plain MessageBox has no
            // verification checkbox, so no opt-out is recorded on
            // this path. The confirmation still gates the delete.
            return MessageBox.Show(
                message,
                _title,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) == DialogResult.Yes;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Builds the confirmation task-dialog page with no
    /// opt-out checkbox.
    /// </summary>
    /// <param name="message">The confirmation body.</param>
    /// <returns>The configured page.</returns>
    internal static TaskDialogPage CreateConfirmationPage(string message) =>
        CreateConfirmationPage(message, session: null);

    /// <summary>
    /// Builds the confirmation task-dialog page. When
    /// <paramref name="session"/> is supplied, the page
    /// carries the "don't ask again this session"
    /// verification checkbox (ADR-0042); null builds the
    /// plain two-button page.
    /// </summary>
    /// <param name="message">The confirmation body.</param>
    /// <param name="session">
    /// The session whose opt-out the checkbox drives, or null
    /// for a plain page with no opt-out checkbox.
    /// </param>
    /// <returns>The configured page.</returns>
    internal static TaskDialogPage CreateConfirmationPage(string message, ConfirmationSession? session)
    {
        ArgumentNullException.ThrowIfNull(message);

        var page = new TaskDialogPage
        {
            Caption = _title,
            Heading = DestructiveCommandPolicy.ConfirmationMainInstruction,
            Text = message,
            Icon = TaskDialogIcon.Information,
        };
        page.Buttons.Add(TaskDialogButton.Yes);
        page.Buttons.Add(TaskDialogButton.No);
        if (session is not null)
        {
            page.Verification = new TaskDialogVerificationCheckBox
            {
                Text = _dontAskAgainText,
            };
        }
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
