using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>Contract tests for the R2.8 add-row command boundary.</summary>
public class AddRowCommandTests
{
    private static Mock<IGanttRowInserter> InserterWith(GanttRowInsertOutcome outcome)
    {
        var inserter = new Mock<IGanttRowInserter>();
        _ = inserter
            .Setup(i => i.Insert(It.IsAny<GanttEntityType>(), It.IsAny<Func<GanttRowId>>()))
            .Returns(outcome);
        return inserter;
    }

    private static GanttRowId FixedId() => GanttRowId.Parse("G-0123456789abcdef0123456789abcdef");

    [Fact]
    public void Run_throws_for_a_null_inserter()
    {
        Assert.Throws<ArgumentNullException>(
            () => AddRowCommand.Run(null!, FixedId, _ => { }, GanttEntityType.AsPlannedActivity));
    }

    [Fact]
    public void Run_throws_for_a_null_id_generator()
    {
        Mock<IGanttRowInserter> inserter = InserterWith(GanttRowInsertOutcome.Ok(1));
        Assert.Throws<ArgumentNullException>(
            () => AddRowCommand.Run(inserter.Object, null!, _ => { }, GanttEntityType.AsPlannedActivity));
    }

    [Fact]
    public void Run_throws_for_a_null_presenter()
    {
        Mock<IGanttRowInserter> inserter = InserterWith(GanttRowInsertOutcome.Ok(1));
        Assert.Throws<ArgumentNullException>(
            () => AddRowCommand.Run(inserter.Object, FixedId, null!, GanttEntityType.AsPlannedActivity));
    }

    [Fact]
    public void Success_is_silent_and_passes_the_requested_type_and_generator()
    {
        var messages = new List<string>();
        Mock<IGanttRowInserter> inserter = InserterWith(GanttRowInsertOutcome.Ok(4));

        AddRowCommand.Run(
            inserter.Object,
            FixedId,
            messages.Add,
            GanttEntityType.AsPlannedMilestone);

        Assert.Empty(messages);
        inserter.Verify(
            i => i.Insert(GanttEntityType.AsPlannedMilestone, It.IsAny<Func<GanttRowId>>()),
            Times.Once);
    }

    [Theory]
    [InlineData(GanttRowInsertRefusalReason.NoActiveWorkbook, "open workbook")]
    [InlineData(GanttRowInsertRefusalReason.TableMissing, "tblGanttData")]
    [InlineData(GanttRowInsertRefusalReason.TargetProtected, "protected")]
    [InlineData(GanttRowInsertRefusalReason.TypeOptionsUnavailable, "Type dropdown")]
    public void Refusal_surfaces_exactly_one_actionable_message(
        GanttRowInsertRefusalReason refusal,
        string expectedText)
    {
        var messages = new List<string>();
        Mock<IGanttRowInserter> inserter = InserterWith(GanttRowInsertOutcome.Refused(refusal));

        AddRowCommand.Run(
            inserter.Object,
            FixedId,
            messages.Add,
            GanttEntityType.Delineator);

        string message = Assert.Single(messages);
        Assert.Contains(expectedText, message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Presenter_failure_does_not_escape_the_command()
    {
        Mock<IGanttRowInserter> inserter = InserterWith(
            GanttRowInsertOutcome.Refused(GanttRowInsertRefusalReason.TargetProtected));

        Exception? exception = Record.Exception(() => AddRowCommand.Run(
            inserter.Object,
            FixedId,
            _ => throw new InvalidOperationException("dialog down"),
            GanttEntityType.AsPlannedActivity));

        Assert.Null(exception);
    }
}
