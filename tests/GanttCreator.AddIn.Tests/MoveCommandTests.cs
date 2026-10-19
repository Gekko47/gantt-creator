using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>Contract tests for the R5.4 move commands (D5, D6, D9).</summary>
public class MoveCommandTests
{
    private static Mock<IParentChildMutator> MutatorReturning(ParentChildMutationOutcome outcome)
    {
        var mutator = new Mock<IParentChildMutator>();
        _ = mutator.Setup(m => m.MoveUp()).Returns(outcome);
        _ = mutator.Setup(m => m.MoveDown()).Returns(outcome);
        return mutator;
    }

    private static Mock<IInsertedRowSelector> Selector() => new();

    [Fact]
    public void MoveUp_throws_for_a_null_mutator()
    {
        Assert.Throws<ArgumentNullException>(
            () => MoveUpCommand.Run(null!, Selector().Object, _ => { }));
    }

    [Fact]
    public void MoveUp_throws_for_a_null_selector()
    {
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(2, 4));
        Assert.Throws<ArgumentNullException>(
            () => MoveUpCommand.Run(mutator.Object, null!, _ => { }));
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
    public void MoveDown_throws_for_a_null_selector()
    {
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(2, 4));
        Assert.Throws<ArgumentNullException>(
            () => MoveDownCommand.Run(mutator.Object, null!, _ => { }));
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
    public void MoveUp_success_is_silent_and_calls_move_up_exactly_once()
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(6, 4));
        Mock<IInsertedRowSelector> selector = Selector();

        MoveUpCommand.Run(mutator.Object, selector.Object, messages.Add);

        Assert.Empty(messages);
        mutator.Verify(m => m.MoveUp(), Times.Once);
        mutator.Verify(m => m.MoveDown(), Times.Never);
        // A boundary no-op moved nothing: the selection stays put.
        selector.Verify(s => s.SelectBodyRow(It.IsAny<int>()), Times.Never);
    }

    [Fact]
    public void MoveDown_success_is_silent_and_calls_move_down_exactly_once()
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(3, 4));
        Mock<IInsertedRowSelector> selector = Selector();

        MoveDownCommand.Run(mutator.Object, selector.Object, messages.Add);

        Assert.Empty(messages);
        mutator.Verify(m => m.MoveDown(), Times.Once);
        mutator.Verify(m => m.MoveUp(), Times.Never);
        selector.Verify(s => s.SelectBodyRow(It.IsAny<int>()), Times.Never);
    }

    [Fact]
    public void MoveUp_success_moves_the_selection_onto_the_moved_row()
    {
        // R5.4 D9: the outcome names the row's new one-based body
        // index and the command puts the selection on it, so the
        // next move repeats on the same row.
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(6, 4, 2));
        Mock<IInsertedRowSelector> selector = Selector();

        MoveUpCommand.Run(mutator.Object, selector.Object, messages.Add);

        Assert.Empty(messages);
        selector.Verify(s => s.SelectBodyRow(2), Times.Once);
    }

    [Fact]
    public void MoveDown_success_moves_the_selection_onto_the_moved_row()
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(3, 4, 5));
        Mock<IInsertedRowSelector> selector = Selector();

        MoveDownCommand.Run(mutator.Object, selector.Object, messages.Add);

        Assert.Empty(messages);
        selector.Verify(s => s.SelectBodyRow(5), Times.Once);
    }

    [Fact]
    public void MoveUp_selection_failure_degrades_to_silence()
    {
        // The row moved; only the cursor stays put. A host selection
        // failure must neither escape into Excel nor surface as an
        // error dialog.
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(6, 4, 2));
        Mock<IInsertedRowSelector> selector = Selector();
        selector.Setup(s => s.SelectBodyRow(It.IsAny<int>()))
            .Throws(new InvalidOperationException("host busy"));

        Exception? exception = Record.Exception(() =>
            MoveUpCommand.Run(mutator.Object, selector.Object, messages.Add));

        Assert.Null(exception);
        Assert.Empty(messages);
    }

    [Fact]
    public void MoveDown_selection_failure_degrades_to_silence()
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(3, 4, 5));
        Mock<IInsertedRowSelector> selector = Selector();
        selector.Setup(s => s.SelectBodyRow(It.IsAny<int>()))
            .Throws(new InvalidOperationException("host busy"));

        Exception? exception = Record.Exception(() =>
            MoveDownCommand.Run(mutator.Object, selector.Object, messages.Add));

        Assert.Null(exception);
        Assert.Empty(messages);
    }

    [Fact]
    public void MoveUp_refusal_leaves_the_selection_alone()
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.TableMissing));
        Mock<IInsertedRowSelector> selector = Selector();

        MoveUpCommand.Run(mutator.Object, selector.Object, messages.Add);

        Assert.Single(messages);
        selector.Verify(s => s.SelectBodyRow(It.IsAny<int>()), Times.Never);
    }

    [Fact]
    public void MoveDown_refusal_leaves_the_selection_alone()
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.TableMissing));
        Mock<IInsertedRowSelector> selector = Selector();

        MoveDownCommand.Run(mutator.Object, selector.Object, messages.Add);

        Assert.Single(messages);
        selector.Verify(s => s.SelectBodyRow(It.IsAny<int>()), Times.Never);
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
        Mock<IInsertedRowSelector> selector = Selector();

        MoveUpCommand.Run(mutator.Object, selector.Object, messages.Add);

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
        Mock<IInsertedRowSelector> selector = Selector();

        MoveDownCommand.Run(mutator.Object, selector.Object, messages.Add);

        string message = Assert.Single(messages);
        Assert.Contains(expectedText, message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MoveUp_planner_refusal_names_the_row_and_the_rule()
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.PlanRefused(ParentChildRefusalReason.SelectionUnknown, selectedRowNumber: 4));
        Mock<IInsertedRowSelector> selector = Selector();

        MoveUpCommand.Run(mutator.Object, selector.Object, messages.Add);

        string message = Assert.Single(messages);
        Assert.Contains("could not find the selected row", message, StringComparison.Ordinal);
    }

    [Fact]
    public void MoveDown_planner_refusal_names_the_verb()
    {
        var messages = new List<string>();
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.PlanRefused(ParentChildRefusalReason.SelectionUnknown, selectedRowNumber: 4));
        Mock<IInsertedRowSelector> selector = Selector();

        MoveDownCommand.Run(mutator.Object, selector.Object, messages.Add);

        string message = Assert.Single(messages);
        Assert.Contains("could not find the selected row", message, StringComparison.Ordinal);
    }

    [Fact]
    public void MoveUp_success_that_wrote_cells_invokes_the_post_mutation_refresh_once()
    {
        // R5.4 D11-amended: the move rotates content, so the
        // chart is rebuilt immediately for the shapes to follow.
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(6, 4, movedToRowNumber: 1));
        Mock<IInsertedRowSelector> selector = Selector();
        var refreshed = 0;

        MoveUpCommand.Run(mutator.Object, selector.Object, _ => { }, () => refreshed++);

        Assert.Equal(1, refreshed);
    }

    [Fact]
    public void MoveDown_success_that_wrote_cells_invokes_the_post_mutation_refresh_once()
    {
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(6, 4, movedToRowNumber: 3));
        Mock<IInsertedRowSelector> selector = Selector();
        var refreshed = 0;

        MoveDownCommand.Run(mutator.Object, selector.Object, _ => { }, () => refreshed++);

        Assert.Equal(1, refreshed);
    }

    [Fact]
    public void MoveUp_boundary_no_op_skips_the_post_mutation_refresh()
    {
        // R5.4 D4: a boundary no-op wrote zero cells and changed
        // nothing, so there is no stale chart to rebuild.
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(0, 4));
        Mock<IInsertedRowSelector> selector = Selector();
        var refreshed = 0;

        MoveUpCommand.Run(mutator.Object, selector.Object, _ => { }, () => refreshed++);

        Assert.Equal(0, refreshed);
    }

    [Fact]
    public void MoveDown_boundary_no_op_skips_the_post_mutation_refresh()
    {
        Mock<IParentChildMutator> mutator = MutatorReturning(
            ParentChildMutationOutcome.Ok(0, 4));
        Mock<IInsertedRowSelector> selector = Selector();
        var refreshed = 0;

        MoveDownCommand.Run(mutator.Object, selector.Object, _ => { }, () => refreshed++);

        Assert.Equal(0, refreshed);
    }

    [Fact]
    public void A_refused_move_never_invokes_the_post_mutation_refresh()
    {
        // Both verbs: a refusal mutated nothing, so the chart is
        // left alone.
        Mock<IParentChildMutator> refused = MutatorReturning(
            ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.WriteFailed));
        Mock<IInsertedRowSelector> selector = Selector();
        var refreshed = 0;

        MoveUpCommand.Run(refused.Object, selector.Object, _ => { }, () => refreshed++);
        MoveDownCommand.Run(refused.Object, selector.Object, _ => { }, () => refreshed++);

        Assert.Equal(0, refreshed);
    }
}
