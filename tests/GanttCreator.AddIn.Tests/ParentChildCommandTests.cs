using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>Contract tests for the R5.3 parent/child commands (D6).</summary>
public class ParentChildCommandTests
{
    private static Mock<IParentChildMutator> MutatorReturning(ParentChildMutationOutcome outcome)
    {
        var mutator = new Mock<IParentChildMutator>();
        _ = mutator.Setup(m => m.Promote()).Returns(outcome);
        _ = mutator.Setup(m => m.Demote()).Returns(outcome);
        _ = mutator.Setup(m => m.MoveUp()).Returns(outcome);
        _ = mutator.Setup(m => m.MoveDown()).Returns(outcome);
        return mutator;
    }

    private static Mock<IInsertedRowSelector> Selector() => new();

    [Fact]
    public void Promote_throws_for_a_null_mutator()
    {
        Assert.Throws<ArgumentNullException>(
            () => PromoteToParentCommand.Run(null!, _ => { }));
    }

    [Fact]
    public void Promote_throws_for_a_null_presenter()
    {
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(2, 4));
        Assert.Throws<ArgumentNullException>(
            () => PromoteToParentCommand.Run(mutator.Object, null!));
    }

    [Fact]
    public void Demote_throws_for_a_null_mutator()
    {
        Assert.Throws<ArgumentNullException>(
            () => DemoteToChildCommand.Run(null!, _ => { }));
    }

    [Fact]
    public void Demote_throws_for_a_null_presenter()
    {
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(2, 4));
        Assert.Throws<ArgumentNullException>(
            () => DemoteToChildCommand.Run(mutator.Object, null!));
    }

    [Fact]
    public void MoveUp_throws_for_a_null_mutator()
    {
        Assert.Throws<ArgumentNullException>(
            () => MoveUpCommand.Run(null!, Selector().Object, _ => { }));
    }

    [Fact]
    public void MoveUp_throws_for_a_null_presenter()
    {
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(2, 4));
        Assert.Throws<ArgumentNullException>(
            () => MoveUpCommand.Run(mutator.Object, Selector().Object, null!));
    }

    [Fact]
    public void MoveDown_throws_for_a_null_mutator()
    {
        Assert.Throws<ArgumentNullException>(
            () => MoveDownCommand.Run(null!, Selector().Object, _ => { }));
    }

    [Fact]
    public void MoveDown_throws_for_a_null_presenter()
    {
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(2, 4));
        Assert.Throws<ArgumentNullException>(
            () => MoveDownCommand.Run(mutator.Object, Selector().Object, null!));
    }

    [Fact]
    public void Promote_success_is_silent_and_calls_promote_exactly_once()
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(6, 4));

        PromoteToParentCommand.Run(mutator.Object, messages.Add);

        Assert.Empty(messages);
        mutator.Verify(m => m.Promote(), Times.Once);
        mutator.Verify(m => m.Demote(), Times.Never);
    }

    [Fact]
    public void Demote_success_is_silent_and_calls_demote_exactly_once()
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(3, 4));

        DemoteToChildCommand.Run(mutator.Object, messages.Add);

        Assert.Empty(messages);
        mutator.Verify(m => m.Demote(), Times.Once);
        mutator.Verify(m => m.Promote(), Times.Never);
    }

    [Fact]
    public void MoveUp_success_is_silent_and_calls_move_up_exactly_once()
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(5, 3));

        MoveUpCommand.Run(mutator.Object, Selector().Object, messages.Add);

        Assert.Empty(messages);
        mutator.Verify(m => m.MoveUp(), Times.Once);
        mutator.Verify(m => m.MoveDown(), Times.Never);
    }

    [Fact]
    public void MoveDown_success_is_silent_and_calls_move_down_exactly_once()
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(5, 3));

        MoveDownCommand.Run(mutator.Object, Selector().Object, messages.Add);

        Assert.Empty(messages);
        mutator.Verify(m => m.MoveDown(), Times.Once);
        mutator.Verify(m => m.MoveUp(), Times.Never);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void A_move_at_the_boundary_of_the_set_is_a_silent_no_op(bool moveUp)
    {
        // R5.4 D4: the top/bottom of the selection's set is a success with
        // zero cells written. The command reports nothing — the user asked
        // for a move that cannot happen.
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(cellsWritten: 0, selectedRowNumber: 1));

        if (moveUp)
        {
            MoveUpCommand.Run(mutator.Object, Selector().Object, messages.Add);
        }
        else
        {
            MoveDownCommand.Run(mutator.Object, Selector().Object, messages.Add);
        }

        Assert.Empty(messages);
    }

    [Theory]
    [InlineData(ParentChildMutationRefusalReason.NoActiveWorkbook, "open workbook")]
    [InlineData(ParentChildMutationRefusalReason.TableMissing, "tblGanttData")]
    [InlineData(ParentChildMutationRefusalReason.TargetProtected, "protected")]
    [InlineData(ParentChildMutationRefusalReason.SelectionOutsideTable, "Click a cell inside")]
    [InlineData(ParentChildMutationRefusalReason.TableInvalid, "Run Validate")]
    [InlineData(ParentChildMutationRefusalReason.WriteFailed, "Refresh the chart")]
    public void Promote_infrastructure_refusal_surfaces_exactly_one_actionable_message(
        ParentChildMutationRefusalReason refusal,
        string expectedText)
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Refused(refusal));

        PromoteToParentCommand.Run(mutator.Object, messages.Add);

        string message = Assert.Single(messages);
        Assert.Contains(expectedText, message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(ParentChildMutationRefusalReason.NoActiveWorkbook, "open workbook")]
    [InlineData(ParentChildMutationRefusalReason.TableMissing, "tblGanttData")]
    [InlineData(ParentChildMutationRefusalReason.TargetProtected, "protected")]
    [InlineData(ParentChildMutationRefusalReason.SelectionOutsideTable, "Click a cell inside")]
    [InlineData(ParentChildMutationRefusalReason.TableInvalid, "Run Validate")]
    [InlineData(ParentChildMutationRefusalReason.WriteFailed, "Refresh the chart")]
    public void Demote_infrastructure_refusal_surfaces_exactly_one_actionable_message(
        ParentChildMutationRefusalReason refusal,
        string expectedText)
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Refused(refusal));

        DemoteToChildCommand.Run(mutator.Object, messages.Add);

        string message = Assert.Single(messages);
        Assert.Contains(expectedText, message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(ParentChildMutationRefusalReason.NoActiveWorkbook, "open workbook")]
    [InlineData(ParentChildMutationRefusalReason.TableMissing, "tblGanttData")]
    [InlineData(ParentChildMutationRefusalReason.TargetProtected, "protected")]
    [InlineData(ParentChildMutationRefusalReason.SelectionOutsideTable, "Click a cell inside")]
    [InlineData(ParentChildMutationRefusalReason.TableInvalid, "Run Validate")]
    [InlineData(ParentChildMutationRefusalReason.WriteFailed, "Refresh the chart")]
    public void MoveUp_infrastructure_refusal_surfaces_exactly_one_actionable_message(
        ParentChildMutationRefusalReason refusal,
        string expectedText)
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Refused(refusal));

        MoveUpCommand.Run(mutator.Object, Selector().Object, messages.Add);

        string message = Assert.Single(messages);
        Assert.Contains(expectedText, message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(ParentChildMutationRefusalReason.NoActiveWorkbook, "open workbook")]
    [InlineData(ParentChildMutationRefusalReason.TableMissing, "tblGanttData")]
    [InlineData(ParentChildMutationRefusalReason.TargetProtected, "protected")]
    [InlineData(ParentChildMutationRefusalReason.SelectionOutsideTable, "Click a cell inside")]
    [InlineData(ParentChildMutationRefusalReason.TableInvalid, "Run Validate")]
    [InlineData(ParentChildMutationRefusalReason.WriteFailed, "Refresh the chart")]
    public void MoveDown_infrastructure_refusal_surfaces_exactly_one_actionable_message(
        ParentChildMutationRefusalReason refusal,
        string expectedText)
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Refused(refusal));

        MoveDownCommand.Run(mutator.Object, Selector().Object, messages.Add);

        string message = Assert.Single(messages);
        Assert.Contains(expectedText, message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Move_planner_refusal_reports_the_selection_failure()
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.PlanRefused(ParentChildRefusalReason.SelectionUnknown, selectedRowNumber: 2));

        MoveUpCommand.Run(mutator.Object, Selector().Object, messages.Add);

        // SelectionUnknown carries no row context: the translator reports it
        // as a selection-resolution failure, which is what the refusal is.
        string message = Assert.Single(messages);
        Assert.Contains("could not find the selected", message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(ParentChildRefusalReason.NotAChild, "nothing to promote")]
    [InlineData(ParentChildRefusalReason.PromotionForbiddenForType, "critical interval")]
    [InlineData(ParentChildRefusalReason.HasChildren, "children of its own")]
    public void Promote_planner_refusal_names_the_row_and_the_rule(
        ParentChildRefusalReason planRefusal,
        string expectedText)
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.PlanRefused(planRefusal, selectedRowNumber: 4));

        PromoteToParentCommand.Run(mutator.Object, messages.Add);

        string message = Assert.Single(messages);
        Assert.Contains("Row 4", message, StringComparison.Ordinal);
        Assert.Contains(expectedText, message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(ParentChildRefusalReason.AlreadyAChild, "nothing to demote")]
    [InlineData(ParentChildRefusalReason.TypeCannotBeChild, "stands on its own")]
    [InlineData(ParentChildRefusalReason.NoEligibleParentAbove, "no activity above it")]
    [InlineData(ParentChildRefusalReason.ParentUnavailable, "seven children each")]
    [InlineData(ParentChildRefusalReason.TooDeep, "grandchild")]
    [InlineData(ParentChildRefusalReason.Cycle, "loop in the parent links")]
    [InlineData(ParentChildRefusalReason.ParentNotSpan, "planned span")]
    public void Demote_planner_refusal_names_the_row_and_the_rule(
        ParentChildRefusalReason planRefusal,
        string expectedText)
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.PlanRefused(planRefusal, selectedRowNumber: 7));

        DemoteToChildCommand.Run(mutator.Object, messages.Add);

        string message = Assert.Single(messages);
        Assert.Contains("Row 7", message, StringComparison.Ordinal);
        Assert.Contains(expectedText, message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void An_unknown_infrastructure_enum_value_falls_back_to_the_generic_message()
    {
        // Positive test for the fallback arm: cast an out-of-range int so the
        // switch's default branch must fire rather than returning nothing.
        var messages = new List<string>();
        var outcome = new ParentChildMutationOutcome(
            Succeeded: false,
            Refusal: (ParentChildMutationRefusalReason)99,
            PlanRefusal: null,
            SelectedRowNumber: null,
            CellsWritten: 0);
        Mock<IParentChildMutator> mutator = MutatorReturning(outcome);

        PromoteToParentCommand.Run(mutator.Object, messages.Add);

        string message = Assert.Single(messages);
        Assert.Contains("Diagnostics", message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void An_unknown_planner_refusal_falls_back_to_the_verb_specific_message()
    {
        // Positive test for the planner fallback arm: an out-of-range plan
        // refusal on a resolved row must still produce one row-named message.
        var messages = new List<string>();
        var outcome = new ParentChildMutationOutcome(
            Succeeded: false,
            Refusal: null,
            PlanRefusal: (ParentChildRefusalReason)99,
            SelectedRowNumber: 3,
            CellsWritten: 0);
        Mock<IParentChildMutator> mutator = MutatorReturning(outcome);

        DemoteToChildCommand.Run(mutator.Object, messages.Add);

        string message = Assert.Single(messages);
        Assert.Contains("Row 3", message, StringComparison.Ordinal);
        Assert.Contains("demoted", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Promote_presenter_failure_does_not_escape_the_command()
    {
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.TargetProtected));

        Exception? exception = Record.Exception(() => PromoteToParentCommand.Run(
            mutator.Object,
            _ => throw new InvalidOperationException("dialog down")));

        Assert.Null(exception);
    }

    [Fact]
    public void Demote_presenter_failure_does_not_escape_the_command()
    {
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.TargetProtected));

        Exception? exception = Record.Exception(() => DemoteToChildCommand.Run(
            mutator.Object,
            _ => throw new InvalidOperationException("dialog down")));

        Assert.Null(exception);
    }

    [Fact]
    public void The_translator_rejects_a_null_outcome()
    {
        Assert.Throws<ArgumentNullException>(
            () => ParentChildRefusalTranslator.TranslatePromote(null!));
        Assert.Throws<ArgumentNullException>(
            () => ParentChildRefusalTranslator.TranslateDemote(null!));
    }
}
