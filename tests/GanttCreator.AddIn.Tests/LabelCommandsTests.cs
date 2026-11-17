using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contract tests for the three label commands. Each re-reads the selection at
/// execution (ruling 1), refuses a non-single-row selection without touching the
/// mutator, stays silent on success, presents the translated refusal otherwise,
/// and never lets a dialog failure escape into Excel.
/// </summary>
public sealed class SetDescriptionLabelCommandTests
{
    [Fact]
    public void A_null_mutator_is_a_programming_error() => _ = Assert.Throws<ArgumentNullException>(
        () => SetDescriptionLabelCommand.Run(null!, Labels.Selection(), _ => { }, GanttLabelPosition.Left));

    [Fact]
    public void A_null_selection_is_a_programming_error() => _ = Assert.Throws<ArgumentNullException>(
        () => SetDescriptionLabelCommand.Run(
            Labels.LabelMutator(LabelPositionMutationOutcome.Ok(1, 3)).Object,
            null!,
            _ => { },
            GanttLabelPosition.Left));

    [Fact]
    public void A_null_presenter_is_a_programming_error() => _ = Assert.Throws<ArgumentNullException>(
        () => SetDescriptionLabelCommand.Run(
            Labels.LabelMutator(LabelPositionMutationOutcome.Ok(1, 3)).Object,
            Labels.Selection(),
            null!,
            GanttLabelPosition.Left));

    [Fact]
    public void A_non_single_row_selection_refuses_without_running_the_mutator()
    {
        Mock<ILabelPositionMutator> mutator = Labels.LabelMutator(LabelPositionMutationOutcome.Ok(1, 3));
        List<string> messages = [];

        SetDescriptionLabelCommand.Run(
            mutator.Object,
            Labels.Selection(new SelectionFacts(true, false, null, true, 0)),
            messages.Add,
            GanttLabelPosition.Left);

        Assert.Contains("not the table header", Assert.Single(messages), StringComparison.Ordinal);
        mutator.Verify(m => m.SetLabelPosition(It.IsAny<GanttLabelPosition?>()), Times.Never);
    }

    [Fact]
    public void A_successful_write_is_silent_and_passes_the_position()
    {
        Mock<ILabelPositionMutator> mutator = Labels.LabelMutator(LabelPositionMutationOutcome.Ok(1, 3));
        List<string> messages = [];

        SetDescriptionLabelCommand.Run(mutator.Object, Labels.Selection(), messages.Add, GanttLabelPosition.Right);

        Assert.Empty(messages);
        mutator.Verify(m => m.SetLabelPosition(GanttLabelPosition.Right), Times.Once);
    }

    [Fact]
    public void A_refusal_is_presented_as_the_translated_message()
    {
        Mock<ILabelPositionMutator> mutator = Labels.LabelMutator(
            LabelPositionMutationOutcome.Refused(LabelPositionMutationRefusalReason.TargetProtected));
        List<string> messages = [];

        SetDescriptionLabelCommand.Run(mutator.Object, Labels.Selection(), messages.Add, GanttLabelPosition.Right);

        Assert.Contains("protected", Assert.Single(messages), StringComparison.Ordinal);
    }

    [Fact]
    public void A_presenter_failure_does_not_escape()
    {
        Mock<ILabelPositionMutator> mutator = Labels.LabelMutator(
            LabelPositionMutationOutcome.Refused(LabelPositionMutationRefusalReason.TargetProtected));

        SetDescriptionLabelCommand.Run(
            mutator.Object,
            Labels.Selection(),
            _ => throw new InvalidOperationException("the dialog is gone"),
            GanttLabelPosition.Right);
    }
}

/// <summary>Contract tests for the date-label command (the same shape).</summary>
public sealed class SetDateLabelCommandTests
{
    [Fact]
    public void A_null_mutator_is_a_programming_error() => _ = Assert.Throws<ArgumentNullException>(
        () => SetDateLabelCommand.Run(null!, Labels.Selection(), _ => { }, Labels.ChosenDateLabel));

    [Fact]
    public void A_non_single_row_selection_refuses_without_running_the_mutator()
    {
        Mock<IDateLabelPositionMutator> mutator = Labels.DateMutator(DateLabelMutationOutcome.Ok(1, 3));
        List<string> messages = [];

        SetDateLabelCommand.Run(
            mutator.Object,
            Labels.Selection(new SelectionFacts(true, false, null, false, 2)),
            messages.Add,
            Labels.ChosenDateLabel);

        Assert.Contains("one row at a time", Assert.Single(messages), StringComparison.Ordinal);
        mutator.Verify(m => m.SetDateLabelPosition(It.IsAny<GanttDateLabelPosition?>()), Times.Never);
    }

    [Fact]
    public void A_successful_write_is_silent_and_passes_the_choice()
    {
        Mock<IDateLabelPositionMutator> mutator = Labels.DateMutator(DateLabelMutationOutcome.Ok(1, 3));
        List<string> messages = [];

        SetDateLabelCommand.Run(mutator.Object, Labels.Selection(), messages.Add, Labels.ChosenDateLabel);

        Assert.Empty(messages);
        mutator.Verify(m => m.SetDateLabelPosition(Labels.ChosenDateLabel), Times.Once);
    }

    [Fact]
    public void A_refusal_is_presented_as_the_translated_message()
    {
        Mock<IDateLabelPositionMutator> mutator = Labels.DateMutator(
            DateLabelMutationOutcome.Refused(DateLabelMutationRefusalReason.WriteFailed));
        List<string> messages = [];

        SetDateLabelCommand.Run(mutator.Object, Labels.Selection(), messages.Add, Labels.ChosenDateLabel);

        Assert.Contains("Excel refused a write", Assert.Single(messages), StringComparison.Ordinal);
    }

    [Fact]
    public void A_presenter_failure_does_not_escape()
    {
        Mock<IDateLabelPositionMutator> mutator = Labels.DateMutator(
            DateLabelMutationOutcome.Refused(DateLabelMutationRefusalReason.WriteFailed));

        SetDateLabelCommand.Run(
            mutator.Object,
            Labels.Selection(),
            _ => throw new InvalidOperationException("the dialog is gone"),
            Labels.ChosenDateLabel);
    }
}

/// <summary>Contract tests for the date-label placement command (the same shape).</summary>
public sealed class SetDateLabelPlacementCommandTests
{
    [Fact]
    public void A_null_mutator_is_a_programming_error() => _ = Assert.Throws<ArgumentNullException>(
        () => SetDateLabelPlacementCommand.Run(null!, Labels.Selection(), _ => { }, GanttLabelPosition.TopLeft));

    [Fact]
    public void A_non_single_row_selection_refuses_without_running_the_mutator()
    {
        Mock<IDateLabelPositionMutator> mutator = Labels.DateMutator(DateLabelMutationOutcome.Ok(1, 3));
        List<string> messages = [];

        SetDateLabelPlacementCommand.Run(
            mutator.Object,
            Labels.Selection(SelectionFacts.None),
            messages.Add,
            GanttLabelPosition.TopLeft);

        Assert.Contains("No Gantt row is selected", Assert.Single(messages), StringComparison.Ordinal);
        mutator.Verify(m => m.SetDateLabelPlacement(It.IsAny<GanttLabelPosition?>()), Times.Never);
    }

    [Fact]
    public void A_successful_write_is_silent_and_passes_the_corner()
    {
        Mock<IDateLabelPositionMutator> mutator = Labels.DateMutator(DateLabelMutationOutcome.Ok(1, 3));
        List<string> messages = [];

        SetDateLabelPlacementCommand.Run(
            mutator.Object, Labels.Selection(), messages.Add, GanttLabelPosition.BottomRight);

        Assert.Empty(messages);
        mutator.Verify(m => m.SetDateLabelPlacement(GanttLabelPosition.BottomRight), Times.Once);
    }

    [Fact]
    public void A_refusal_is_presented_as_the_translated_message()
    {
        Mock<IDateLabelPositionMutator> mutator = Labels.DateMutator(
            DateLabelMutationOutcome.Refused(DateLabelMutationRefusalReason.TypeUnknown, 4));
        List<string> messages = [];

        SetDateLabelPlacementCommand.Run(
            mutator.Object, Labels.Selection(), messages.Add, GanttLabelPosition.BottomRight);

        Assert.StartsWith("Row 4 ", Assert.Single(messages), StringComparison.Ordinal);
    }

    [Fact]
    public void A_presenter_failure_does_not_escape()
    {
        Mock<IDateLabelPositionMutator> mutator = Labels.DateMutator(
            DateLabelMutationOutcome.Refused(DateLabelMutationRefusalReason.TypeUnknown, 4));

        SetDateLabelPlacementCommand.Run(
            mutator.Object,
            Labels.Selection(),
            _ => throw new InvalidOperationException("the dialog is gone"),
            GanttLabelPosition.BottomRight);
    }
}

/// <summary>The shared substitutes for the three label-command test classes.</summary>
internal static class Labels
{
    /// <summary>A date-label choice that is guaranteed to be a defined member.</summary>
    public static GanttDateLabelPosition ChosenDateLabel =>
        Enum.GetValues<GanttDateLabelPosition>().Last();

    /// <summary>A description-label mutator returning the supplied outcome.</summary>
    /// <param name="outcome">The outcome to return.</param>
    /// <returns>The mock, for call verification.</returns>
    public static Mock<ILabelPositionMutator> LabelMutator(LabelPositionMutationOutcome outcome)
    {
        Mock<ILabelPositionMutator> mutator = new();
        _ = mutator.Setup(m => m.SetLabelPosition(It.IsAny<GanttLabelPosition?>())).Returns(outcome);
        return mutator;
    }

    /// <summary>A date-label mutator returning the supplied outcome.</summary>
    /// <param name="outcome">The outcome to return.</param>
    /// <returns>The mock, for call verification.</returns>
    public static Mock<IDateLabelPositionMutator> DateMutator(DateLabelMutationOutcome outcome)
    {
        Mock<IDateLabelPositionMutator> mutator = new();
        _ = mutator.Setup(m => m.SetDateLabelPosition(It.IsAny<GanttDateLabelPosition?>())).Returns(outcome);
        _ = mutator.Setup(m => m.SetDateLabelPlacement(It.IsAny<GanttLabelPosition?>())).Returns(outcome);
        return mutator;
    }

    /// <summary>A selection identifier reporting exactly one body row unless told otherwise.</summary>
    /// <param name="facts">The facts to report.</param>
    /// <returns>The substitute selection identifier.</returns>
    public static ISelectionIdentifier Selection(SelectionFacts? facts = null)
    {
        Mock<ISelectionIdentifier> selection = new();
        _ = selection.Setup(s => s.Identify()).Returns(facts ?? new SelectionFacts(true, true, 3, false, 1));
        return selection.Object;
    }
}
