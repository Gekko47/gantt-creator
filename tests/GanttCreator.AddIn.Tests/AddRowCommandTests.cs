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

    private sealed class RecordingSelector : IInsertedRowSelector
    {
        public List<int> Selected { get; } = [];

        public int Calls { get; private set; }

        public bool Throws { get; set; }

        public void SelectBodyRow(int bodyIndex)
        {
            Calls++;
            if (Throws)
            {
                throw new InvalidOperationException("selection refused");
            }

            Selected.Add(bodyIndex);
        }
    }

    private static RecordingSelector NoopSelector() => new();

    [Fact]
    public void Run_throws_for_a_null_inserter()
    {
        Assert.Throws<ArgumentNullException>(
            () => AddRowCommand.Run(null!, FixedId, NoopSelector(), _ => { }, GanttEntityType.AsPlannedActivity));
    }

    [Fact]
    public void Run_throws_for_a_null_id_generator()
    {
        Mock<IGanttRowInserter> inserter = InserterWith(GanttRowInsertOutcome.Ok(1));
        Assert.Throws<ArgumentNullException>(
            () => AddRowCommand.Run(inserter.Object, null!, NoopSelector(), _ => { }, GanttEntityType.AsPlannedActivity));
    }

    [Fact]
    public void Run_throws_for_a_null_selector()
    {
        Mock<IGanttRowInserter> inserter = InserterWith(GanttRowInsertOutcome.Ok(1));
        Assert.Throws<ArgumentNullException>(
            () => AddRowCommand.Run(inserter.Object, FixedId, null!, _ => { }, GanttEntityType.AsPlannedActivity));
    }

    [Fact]
    public void Run_throws_for_a_null_presenter()
    {
        Mock<IGanttRowInserter> inserter = InserterWith(GanttRowInsertOutcome.Ok(1));
        Assert.Throws<ArgumentNullException>(
            () => AddRowCommand.Run(inserter.Object, FixedId, NoopSelector(), null!, GanttEntityType.AsPlannedActivity));
    }

    [Fact]
    public void Success_selects_the_inserted_row_body_index()
    {
        // The inserted row must be where the user's cursor lands, otherwise the
        // user has to hunt for the row they just created.
        var selector = new RecordingSelector();
        Mock<IGanttRowInserter> inserter = InserterWith(GanttRowInsertOutcome.Ok(4));

        AddRowCommand.Run(
            inserter.Object,
            FixedId,
            selector,
            _ => { },
            GanttEntityType.AsPlannedMilestone);

        Assert.Equal(4, Assert.Single(selector.Selected));
    }

    [Fact]
    public void Refusal_never_moves_the_selection()
    {
        var selector = new RecordingSelector();
        Mock<IGanttRowInserter> inserter = InserterWith(
            GanttRowInsertOutcome.Refused(GanttRowInsertRefusalReason.TargetProtected));

        AddRowCommand.Run(inserter.Object, FixedId, selector, _ => { }, GanttEntityType.AsPlannedActivity);

        Assert.Equal(0, selector.Calls);
    }

    [Fact]
    public void Selection_failure_does_not_escape_or_report_an_error()
    {
        // The row was inserted; a refused selection must neither throw into Excel
        // nor claim the add failed.
        var messages = new List<string>();
        var selector = new RecordingSelector { Throws = true };
        Mock<IGanttRowInserter> inserter = InserterWith(GanttRowInsertOutcome.Ok(2));

        Exception? exception = Record.Exception(() => AddRowCommand.Run(
            inserter.Object,
            FixedId,
            selector,
            messages.Add,
            GanttEntityType.AsPlannedActivity));

        Assert.Null(exception);
        Assert.Equal(1, selector.Calls);
        Assert.Empty(messages);
    }

    [Fact]
    public void Success_is_silent_and_passes_the_requested_type_and_generator()
    {
        var messages = new List<string>();
        Mock<IGanttRowInserter> inserter = InserterWith(GanttRowInsertOutcome.Ok(4));

        AddRowCommand.Run(
            inserter.Object,
            FixedId,
            NoopSelector(),
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
            NoopSelector(),
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
            NoopSelector(),
            _ => throw new InvalidOperationException("dialog down"),
            GanttEntityType.AsPlannedActivity));

        Assert.Null(exception);
    }
}
