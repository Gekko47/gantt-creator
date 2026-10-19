using GanttCreator.Core;
using GanttCreator.Office;
using Moq;
using Xunit;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contract tests for the R5.4 remove command (D6): the
/// ADR-0008 confirmation gates the mutator, No is the safe
/// default that mutates nothing, and refusals surface exactly
/// one actionable message.
/// </summary>
public class RemoveCommandTests
{
    private static Mock<IParentChildMutator> MutatorReturning(ParentChildMutationOutcome outcome)
    {
        var mutator = new Mock<IParentChildMutator>();
        _ = mutator.Setup(m => m.Remove()).Returns(outcome);
        return mutator;
    }

    [Fact]
    public void Remove_throws_for_a_null_mutator()
    {
        Assert.Throws<ArgumentNullException>(
            () => RemoveCommand.Run(null!, (_, _) => true, _ => { }));
    }

    [Fact]
    public void Remove_throws_for_a_null_confirm()
    {
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(2, 4));
        Assert.Throws<ArgumentNullException>(
            () => RemoveCommand.Run(mutator.Object, null!, _ => { }));
    }

    [Fact]
    public void Remove_throws_for_a_null_presenter()
    {
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(2, 4));
        Assert.Throws<ArgumentNullException>(
            () => RemoveCommand.Run(mutator.Object, (_, _) => true, null!));
    }

    [Fact]
    public void Declining_the_confirmation_never_calls_the_mutator()
    {
        // ADR-0008 D3: No is the safe default.
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(1, 1));

        RemoveCommand.Run(mutator.Object, (_, _) => false, messages.Add);

        Assert.Empty(messages);
        mutator.Verify(m => m.Remove(), Times.Never);
    }

    [Fact]
    public void Remove_confirmation_names_the_command_and_the_undo_cost()
    {
        // ADR-0008 D2: the confirmation body names the command
        // and states that it cannot be undone.
        string? capturedBody = null;
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(1, 1));

        RemoveCommand.Run(mutator.Object, (body, _) => { capturedBody = body; return true; }, _ => { });

        ConfirmationText expected = DestructiveCommandPolicy.BuildConfirmationText(
            DestructiveCommandPolicy.CommandClass.RemoveRow);
        Assert.Equal(expected.Body, capturedBody);
        Assert.Contains("Remove", capturedBody, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("cannot be undone", capturedBody, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void The_confirmation_dialog_renders_the_shared_destructive_heading()
    {
        TaskDialogPage page = CommandErrorDialog.CreateConfirmationPage("body");

        Assert.Equal(
            DestructiveCommandPolicy.ConfirmationMainInstruction,
            page.Heading);
    }

    [Fact]
    public void Remove_success_is_silent_and_calls_remove_exactly_once()
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(6, 4));

        RemoveCommand.Run(mutator.Object, (_, _) => true, messages.Add);

        Assert.Empty(messages);
        mutator.Verify(m => m.Remove(), Times.Once);
    }

    [Theory]
    [InlineData(ParentChildMutationRefusalReason.NoActiveWorkbook, "open workbook")]
    [InlineData(ParentChildMutationRefusalReason.TableMissing, "tblGanttData")]
    [InlineData(ParentChildMutationRefusalReason.TargetProtected, "protected")]
    [InlineData(ParentChildMutationRefusalReason.SelectionOutsideTable, "Click a cell inside")]
    [InlineData(ParentChildMutationRefusalReason.TableInvalid, "Run Validate")]
    [InlineData(ParentChildMutationRefusalReason.WriteFailed, "Refresh the chart")]
    public void Remove_infrastructure_refusal_surfaces_exactly_one_actionable_message(
        ParentChildMutationRefusalReason refusal,
        string expectedText)
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Refused(refusal));

        RemoveCommand.Run(mutator.Object, (_, _) => true, messages.Add);

        string message = Assert.Single(messages);
        Assert.Contains(expectedText, message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Remove_success_invokes_the_post_mutation_refresh_once()
    {
        // R5.4 D11-amended: the delete removed a worksheet row and
        // its anchored shapes, so the chart is rebuilt immediately.
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(6, 4));
        var refreshed = 0;

        RemoveCommand.Run(mutator.Object, (_, _) => true, _ => { }, () => refreshed++);

        Assert.Equal(1, refreshed);
    }

    [Fact]
    public void Remove_of_a_lone_row_with_zero_cells_written_still_refreshes()
    {
        // A lone-row delete resequences nothing (zero cells) but
        // still changed the sheet, so the refresh must run.
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(0, 1));
        var refreshed = 0;

        RemoveCommand.Run(mutator.Object, (_, _) => true, _ => { }, () => refreshed++);

        Assert.Equal(1, refreshed);
    }

    [Fact]
    public void Declining_the_confirmation_never_refreshes()
    {
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(6, 4));
        var refreshed = 0;

        RemoveCommand.Run(mutator.Object, (_, _) => false, _ => { }, () => refreshed++);

        Assert.Equal(0, refreshed);
    }

    [Fact]
    public void A_refused_remove_never_refreshes()
    {
        // The mutation refused, so nothing changed and the chart
        // must not be rebuilt.
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.WriteFailed));
        var refreshed = 0;

        RemoveCommand.Run(mutator.Object, (_, _) => true, _ => { }, () => refreshed++);

        Assert.Equal(0, refreshed);
    }

    [Fact]
    public void ConfirmationSession_RecordChoice_arms_the_opt_out_only_when_confirmed()
    {
        // ADR-0042 D2: the opt-out is recorded only by an
        // explicit confirmation with the checkbox checked.
        var session = new ConfirmationSession();

        session.RecordChoice(confirmed: true, dontAskAgain: true);

        Assert.True(session.DontAskAgain);
    }

    [Fact]
    public void ConfirmationSession_RecordChoice_leaves_the_session_untouched_when_declined()
    {
        // ADR-0042 D2: a declined dialog cannot opt the user
        // out, so the ADR-0008 D3 guard stays in force.
        var session = new ConfirmationSession();

        session.RecordChoice(confirmed: false, dontAskAgain: true);

        Assert.False(session.DontAskAgain);
    }

    [Fact]
    public void A_fresh_session_still_shows_the_confirmation()
    {
        // ADR-0042 D1: the guard is armed by default; the
        // opt-out is never a silent default.
        var session = new ConfirmationSession();
        var mutator = MutatorReturning(ParentChildMutationOutcome.Ok(1, 1));
        var confirmCalls = 0;

        RemoveCommand.Run(
            mutator.Object,
            (_, _) => { confirmCalls++; return true; },
            _ => { },
            onMutated: null,
            session: session);

        Assert.Equal(1, confirmCalls);
        mutator.Verify(m => m.Remove(), Times.Once);
    }

    [Fact]
    public void An_opted_out_session_skips_the_confirmation_and_still_deletes()
    {
        // ADR-0042 D4: once opted out, the delete proceeds
        // without the confirmation dialog, but the row is
        // still deleted.
        var session = new ConfirmationSession { DontAskAgain = true };
        var mutator = MutatorReturning(ParentChildMutationOutcome.Ok(1, 1));
        var confirmCalls = 0;

        RemoveCommand.Run(
            mutator.Object,
            (_, _) => { confirmCalls++; return true; },
            _ => { },
            onMutated: null,
            session: session);

        Assert.Equal(0, confirmCalls);
        mutator.Verify(m => m.Remove(), Times.Once);
    }

    [Fact]
    public void Confirming_with_dont_ask_again_arms_the_session_for_the_next_delete()
    {
        // ADR-0042 D2/D4: the first delete's confirmation,
        // confirmed with the checkbox checked, arms the
        // session so the second delete skips the dialog.
        var session = new ConfirmationSession();
        var mutator = MutatorReturning(ParentChildMutationOutcome.Ok(1, 1));
        var confirmCalls = 0;

        RemoveCommand.Run(
            mutator.Object,
            (body, s) =>
            {
                confirmCalls++;
                s?.RecordChoice(confirmed: true, dontAskAgain: true);
                return true;
            },
            _ => { },
            onMutated: null,
            session: session);
        Assert.Equal(1, confirmCalls);

        RemoveCommand.Run(
            mutator.Object,
            (_, _) => { confirmCalls++; return true; },
            _ => { },
            onMutated: null,
            session: session);

        Assert.Equal(1, confirmCalls);
        mutator.Verify(m => m.Remove(), Times.Exactly(2));
    }

    [Fact]
    public void Declining_the_confirmation_does_not_arm_the_opt_out()
    {
        // ADR-0042 D2: declining leaves the session
        // un-opted-out, so a later delete still confirms.
        var session = new ConfirmationSession();
        var mutator = MutatorReturning(ParentChildMutationOutcome.Ok(1, 1));

        RemoveCommand.Run(
            mutator.Object,
            (body, s) =>
            {
                s?.RecordChoice(confirmed: false, dontAskAgain: true);
                return false;
            },
            _ => { },
            onMutated: null,
            session: session);

        Assert.False(session.DontAskAgain);
        mutator.Verify(m => m.Remove(), Times.Never);
    }

    [Fact]
    public void Confirming_without_the_checkbox_does_not_arm_the_opt_out()
    {
        // ADR-0042 D2: confirming without checking the box
        // keeps the guard armed for the next delete.
        var session = new ConfirmationSession();
        var mutator = MutatorReturning(ParentChildMutationOutcome.Ok(1, 1));

        RemoveCommand.Run(
            mutator.Object,
            (body, s) =>
            {
                s?.RecordChoice(confirmed: true, dontAskAgain: false);
                return true;
            },
            _ => { },
            onMutated: null,
            session: session);

        Assert.False(session.DontAskAgain);
        mutator.Verify(m => m.Remove(), Times.Once);
    }

    [Fact]
    public void The_confirmation_page_carries_the_opt_out_checkbox_when_a_session_is_supplied()
    {
        // ADR-0042 D2: the dialog offers the checkbox only
        // when a session drives the opt-out.
        var session = new ConfirmationSession();

        TaskDialogPage page = CommandErrorDialog.CreateConfirmationPage("body", session);

        Assert.NotNull(page.Verification);
        Assert.Equal("Don't ask again this session", page.Verification!.Text);
        Assert.False(page.Verification.Checked);
    }

    [Fact]
    public void The_confirmation_page_has_no_checkbox_when_no_session_is_supplied()
    {
        // ADR-0042 D6: other destructive commands keep the
        // plain two-button confirmation. TaskDialogPage always
        // carries a verification checkbox instance, but WinForms
        // does not render one whose Text is null or empty, so
        // leaving the label unset keeps the opt-out checkbox
        // hidden.
        TaskDialogPage page = CommandErrorDialog.CreateConfirmationPage("body");

        Assert.True(string.IsNullOrEmpty(page.Verification?.Text));
    }
}
