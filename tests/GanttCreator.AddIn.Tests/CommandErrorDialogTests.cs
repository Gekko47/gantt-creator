using GanttCreator.Core;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The command boundary's error and confirmation dialogs, asserted through their page-building
/// seams so no dialog is ever shown.
/// </summary>
/// <remarks>
/// <para>
/// <c>Show</c> and <c>Confirm</c> block on a real TaskDialog, so what is covered here is everything
/// they hand to it: the caption, heading, body, icon, button set and the ADR-0042 opt-out text that
/// only a confirmation session supplies.
/// </para>
/// <para>
/// Two WinForms behaviours shape the assertions. A page <em>clones</em> the shared
/// <c>TaskDialogButton</c> instances into its own collection, so a button is identified by its text
/// rather than by reference. And <c>TaskDialogPage.Verification</c> materialises a control whose
/// text is <see langword="null"/> until a caller sets it, so "no opt-out offered" means the text is
/// unset, not that the object is absent.
/// </para>
/// <para>
/// The null checks are the boundary's first line of defence — a null message must be refused before
/// any UI is created — and each one is asserted here rather than left to the caller.
/// </para>
/// </remarks>
public class CommandErrorDialogTests
{
    [Fact]
    public void The_error_page_states_the_shared_caption_heading_and_message()
    {
        TaskDialogPage page = CommandErrorDialog.CreatePage("Refresh refused: the table is missing.");

        Assert.Equal("Gantt Creator", page.Caption);
        Assert.Equal("Gantt Creator could not complete the last action", page.Heading);
        Assert.Equal("Refresh refused: the table is missing.", page.Text);
        Assert.Same(TaskDialogIcon.Error, page.Icon);
        TaskDialogButton only = Assert.Single(page.Buttons);
        Assert.Equal("OK", only.Text);
        TaskDialogVerificationCheckBox? verification = page.Verification;
        Assert.NotNull(verification);
        Assert.Null(verification.Text);
    }

    [Fact]
    public void The_error_page_carries_no_opt_out_text()
    {
        // The opt-out belongs to the destructive-command confirmation only. An error page that grew
        // that text would offer a second, unrelated way to suppress a message.
        TaskDialogPage page = CommandErrorDialog.CreatePage("Resolve the configuration sheet, then retry.");

        TaskDialogVerificationCheckBox? verification = page.Verification;
        Assert.NotNull(verification);
        Assert.Null(verification.Text);
    }

    [Fact]
    public void The_plain_confirmation_page_offers_yes_and_no_without_an_opt_out()
    {
        TaskDialogPage page = CommandErrorDialog.CreateConfirmationPage("Remove row 4?");

        Assert.Equal("Gantt Creator", page.Caption);
        Assert.Equal(DestructiveCommandPolicy.ConfirmationMainInstruction, page.Heading);
        Assert.Equal("Remove row 4?", page.Text);
        Assert.Same(TaskDialogIcon.Information, page.Icon);
        Assert.Collection(
            page.Buttons,
            button => Assert.Equal("Yes", button.Text),
            button => Assert.Equal("No", button.Text));
        TaskDialogVerificationCheckBox? verification = page.Verification;
        Assert.NotNull(verification);
        Assert.Null(verification.Text);
    }

    [Fact]
    public void A_session_adds_the_dont_ask_again_checkbox_with_the_pinned_wording()
    {
        ConfirmationSession session = new();

        TaskDialogPage page = CommandErrorDialog.CreateConfirmationPage("Remove row 4?", session);

        Assert.NotNull(page.Verification);
        Assert.Equal("Don't ask again this session", page.Verification.Text);
    }

    [Fact]
    public void The_returned_page_is_unbound_so_a_caller_cannot_show_it_twice()
    {
        // A page carries its controls; handing the same page to a second dialog throws. The seam
        // returns a fresh page per call, which is what keeps the fallback path safe.
        TaskDialogPage first = CommandErrorDialog.CreatePage("One message.");
        TaskDialogPage second = CommandErrorDialog.CreatePage("One message.");

        Assert.NotSame(first, second);
        Assert.NotSame(first.Buttons[0], second.Buttons[0]);
    }

    [Fact]
    public void The_dialog_seams_refuse_a_null_message_before_any_ui_exists()
    {
        Assert.Throws<ArgumentNullException>(() => CommandErrorDialog.CreatePage(null!));
        Assert.Throws<ArgumentNullException>(() => CommandErrorDialog.CreateConfirmationPage(null!));
        Assert.Throws<ArgumentNullException>(
            () => CommandErrorDialog.CreateConfirmationPage(null!, new ConfirmationSession()));
        Assert.Throws<ArgumentNullException>(() => CommandErrorDialog.Confirm(null!));
        Assert.Throws<ArgumentNullException>(() => CommandErrorDialog.Confirm(null!, new ConfirmationSession()));
        Assert.Throws<ArgumentNullException>(() => CommandErrorDialog.Show(null!));
    }
}
