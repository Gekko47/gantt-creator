using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contract tests for the R5.3 WI-E add-child command and the Add child
/// dropdown that offers the permitted child types.
/// </summary>
public sealed class AddChildCommandTests
{
    private static GanttRowId FixedId() => GanttRowId.Parse("G-0123456789abcdef0123456789abcdef");

    private static Mock<IAddChildMutator> MutatorWith(AddChildMutationOutcome outcome)
    {
        var mutator = new Mock<IAddChildMutator>();
        _ = mutator
            .Setup(m => m.AddChild(It.IsAny<GanttEntityType>(), It.IsAny<Func<GanttRowId>>()))
            .Returns(outcome);
        return mutator;
    }

    [Fact]
    public void Run_throws_for_a_null_mutator()
    {
        Assert.Throws<ArgumentNullException>(
            () => AddChildCommand.Run(null!, FixedId, _ => { }, GanttEntityType.AsBuiltActivity));
    }

    [Fact]
    public void Run_throws_for_a_null_id_generator()
    {
        Mock<IAddChildMutator> mutator = MutatorWith(AddChildMutationOutcome.Ok(2, 1, 2));
        Assert.Throws<ArgumentNullException>(
            () => AddChildCommand.Run(mutator.Object, null!, _ => { }, GanttEntityType.AsBuiltActivity));
    }

    [Fact]
    public void Run_throws_for_a_null_presenter()
    {
        Mock<IAddChildMutator> mutator = MutatorWith(AddChildMutationOutcome.Ok(2, 1, 2));
        Assert.Throws<ArgumentNullException>(
            () => AddChildCommand.Run(mutator.Object, FixedId, null!, GanttEntityType.AsBuiltActivity));
    }

    [Fact]
    public void A_successful_add_is_silent()
    {
        // R2.2/R2.6: success is silent; the ribbon state refresh handles any
        // enabled-state consequences.
        Mock<IAddChildMutator> mutator = MutatorWith(AddChildMutationOutcome.Ok(2, 1, 2));
        List<string> presented = [];

        AddChildCommand.Run(mutator.Object, FixedId, presented.Add, GanttEntityType.AsBuiltActivity);

        Assert.Empty(presented);
    }

    [Fact]
    public void The_requested_type_reaches_the_mutator()
    {
        Mock<IAddChildMutator> mutator = MutatorWith(AddChildMutationOutcome.Ok(2, 1, 2));
        GanttEntityType? seen = null;
        _ = mutator
            .Setup(m => m.AddChild(It.IsAny<GanttEntityType>(), It.IsAny<Func<GanttRowId>>()))
            .Callback((GanttEntityType type, Func<GanttRowId> _) => seen = type)
            .Returns(AddChildMutationOutcome.Ok(2, 1, 2));

        AddChildCommand.Run(mutator.Object, FixedId, _ => { }, GanttEntityType.CustomMilestone);

        Assert.Equal(GanttEntityType.CustomMilestone, seen);
    }

    [Fact]
    public void A_planner_refusal_is_presented()
    {
        Mock<IAddChildMutator> mutator = MutatorWith(
            AddChildMutationOutcome.PlanRefused(AddChildRefusalReason.ParentCannotOwnChildren, 1));
        List<string> presented = [];

        AddChildCommand.Run(mutator.Object, FixedId, presented.Add, GanttEntityType.AsBuiltActivity);

        string message = Assert.Single(presented);
        Assert.Contains("cannot have children", message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_capacity_refusal_names_the_limit()
    {
        Mock<IAddChildMutator> mutator = MutatorWith(
            AddChildMutationOutcome.PlanRefused(AddChildRefusalReason.ParentAtCapacity, 1));
        List<string> presented = [];

        AddChildCommand.Run(mutator.Object, FixedId, presented.Add, GanttEntityType.AsBuiltActivity);

        string message = Assert.Single(presented);
        Assert.Contains("maximum number of children", message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_depth_refusal_names_the_reason()
    {
        Mock<IAddChildMutator> mutator = MutatorWith(
            AddChildMutationOutcome.PlanRefused(AddChildRefusalReason.TooDeep, 1));
        List<string> presented = [];

        AddChildCommand.Run(mutator.Object, FixedId, presented.Add, GanttEntityType.AsBuiltActivity);

        string message = Assert.Single(presented);
        Assert.Contains("three levels deep", message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_type_refusal_is_presented()
    {
        Mock<IAddChildMutator> mutator = MutatorWith(
            AddChildMutationOutcome.PlanRefused(AddChildRefusalReason.TypeCannotBeChild, 1));
        List<string> presented = [];

        AddChildCommand.Run(mutator.Object, FixedId, presented.Add, GanttEntityType.Spacer);

        string message = Assert.Single(presented);
        Assert.Contains("cannot be a child", message, StringComparison.Ordinal);
    }

    [Fact]
    public void An_infrastructure_refusal_is_presented()
    {
        Mock<IAddChildMutator> mutator = MutatorWith(
            AddChildMutationOutcome.Refused(AddChildMutationRefusalReason.TargetProtected));
        List<string> presented = [];

        AddChildCommand.Run(mutator.Object, FixedId, presented.Add, GanttEntityType.AsBuiltActivity);

        string message = Assert.Single(presented);
        Assert.Contains("protected", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Every_refusal_reason_has_a_message()
    {
        // The positive control for the translator: a refusal with no message
        // would leave the user with a silent no-op.
        foreach (AddChildRefusalReason reason in Enum.GetValues<AddChildRefusalReason>())
        {
            string message = AddChildCommand.Translate(
                AddChildMutationOutcome.PlanRefused(reason, 1));

            Assert.False(string.IsNullOrWhiteSpace(message), $"No message for {reason}.");
        }

        foreach (AddChildMutationRefusalReason reason in Enum.GetValues<AddChildMutationRefusalReason>())
        {
            string message = AddChildCommand.Translate(
                AddChildMutationOutcome.Refused(reason));

            Assert.False(string.IsNullOrWhiteSpace(message), $"No message for {reason}.");
        }
    }
}
