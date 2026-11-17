using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contract tests for the R5.4 collapse-all and expand-all commands: the
/// Ribbon surface over the native Excel outline.
/// </summary>
public class CollapseGroupsCommandTests
{
    private static Mock<IOutlineGroupPort> OutlineReturning(OutlineCollapseOutcome outcome)
    {
        var outline = new Mock<IOutlineGroupPort>();
        _ = outline
            .Setup(o => o.SetCollapsed(It.IsAny<OutlineCollapseDirection>()))
            .Returns(outcome);
        return outline;
    }

    [Fact]
    public void Collapse_throws_for_a_null_port()
    {
        Assert.Throws<ArgumentNullException>(
            () => CollapseGroupsCommand.Run(null!, OutlineCollapseDirection.Collapse, _ => { }));
    }

    [Fact]
    public void Collapse_throws_for_a_null_presenter()
    {
        Mock<IOutlineGroupPort> outline = OutlineReturning(OutlineCollapseOutcome.Ok(0));
        Assert.Throws<ArgumentNullException>(
            () => CollapseGroupsCommand.Run(outline.Object, OutlineCollapseDirection.Collapse, null!));
    }

    [Fact]
    public void Expand_throws_for_a_null_port()
    {
        Assert.Throws<ArgumentNullException>(
            () => CollapseGroupsCommand.Run(null!, OutlineCollapseDirection.Expand, _ => { }));
    }

    [Fact]
    public void Expand_throws_for_a_null_presenter()
    {
        Mock<IOutlineGroupPort> outline = OutlineReturning(OutlineCollapseOutcome.Ok(0));
        Assert.Throws<ArgumentNullException>(
            () => CollapseGroupsCommand.Run(outline.Object, OutlineCollapseDirection.Expand, null!));
    }

    [Fact]
    public void Collapse_success_is_silent_and_calls_the_port_exactly_once()
    {
        var messages = new List<string>();
        Mock<IOutlineGroupPort> outline = OutlineReturning(OutlineCollapseOutcome.Ok(6));

        CollapseGroupsCommand.Run(outline.Object, OutlineCollapseDirection.Collapse, messages.Add);

        Assert.Empty(messages);
        outline.Verify(o => o.SetCollapsed(OutlineCollapseDirection.Collapse), Times.Once);
    }

    [Fact]
    public void Expand_success_is_silent_and_calls_the_port_exactly_once()
    {
        var messages = new List<string>();
        Mock<IOutlineGroupPort> outline = OutlineReturning(OutlineCollapseOutcome.Ok(6));

        CollapseGroupsCommand.Run(outline.Object, OutlineCollapseDirection.Expand, messages.Add);

        Assert.Empty(messages);
        outline.Verify(o => o.SetCollapsed(OutlineCollapseDirection.Expand), Times.Once);
    }

    [Fact]
    public void A_sheet_with_no_group_is_a_silent_no_op()
    {
        // The port reports zero rows written when the sheet carries no group.
        // That is the state the user asked for, not a failure, so nothing is
        // presented.
        var messages = new List<string>();
        Mock<IOutlineGroupPort> outline = OutlineReturning(OutlineCollapseOutcome.Ok(0));

        CollapseGroupsCommand.Run(outline.Object, OutlineCollapseDirection.Collapse, messages.Add);

        Assert.Empty(messages);
    }

    [Fact]
    public void A_protected_sheet_presents_one_actionable_message()
    {
        var messages = new List<string>();
        Mock<IOutlineGroupPort> outline = OutlineReturning(
            OutlineCollapseOutcome.Refused(OutlineCollapseRefusalReason.TargetProtected));

        CollapseGroupsCommand.Run(outline.Object, OutlineCollapseDirection.Collapse, messages.Add);

        Assert.Single(messages);
        Assert.Contains("protected", messages[0], StringComparison.OrdinalIgnoreCase);
        Assert.Contains("collapsed", messages[0], StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void A_missing_table_presents_one_actionable_message()
    {
        var messages = new List<string>();
        Mock<IOutlineGroupPort> outline = OutlineReturning(
            OutlineCollapseOutcome.Refused(OutlineCollapseRefusalReason.TableMissing));

        CollapseGroupsCommand.Run(outline.Object, OutlineCollapseDirection.Expand, messages.Add);

        Assert.Single(messages);
        Assert.Contains("Initialise sheet", messages[0], StringComparison.OrdinalIgnoreCase);
        Assert.Contains("expanded", messages[0], StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void A_missing_workbook_presents_one_actionable_message()
    {
        var messages = new List<string>();
        Mock<IOutlineGroupPort> outline = OutlineReturning(
            OutlineCollapseOutcome.Refused(OutlineCollapseRefusalReason.NoActiveWorkbook));

        CollapseGroupsCommand.Run(outline.Object, OutlineCollapseDirection.Collapse, messages.Add);

        Assert.Single(messages);
        Assert.Contains("workbook", messages[0], StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void A_presenter_failure_does_not_escape_the_command()
    {
        // CA1031: a dialog failure is outside command behaviour. It must
        // degrade to silence, never escape into Excel.
        Mock<IOutlineGroupPort> outline = OutlineReturning(
            OutlineCollapseOutcome.Refused(OutlineCollapseRefusalReason.TargetProtected));

        CollapseGroupsCommand.Run(
            outline.Object,
            OutlineCollapseDirection.Collapse,
            _ => throw new InvalidOperationException("the host refused the dialog"));
    }
}
