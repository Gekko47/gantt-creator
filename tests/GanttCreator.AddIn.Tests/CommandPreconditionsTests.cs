using GanttCreator.Office;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The shared single-row precondition (finding 4). Each setter command refuses a selection
/// that is not exactly one table body row before it writes, and each keeps its own refusal
/// wording; this pins the shared half so the four call sites cannot drift apart.
/// </summary>
public class CommandPreconditionsTests
{
    [Fact]
    public void Single_body_row_passes_without_presenting()
    {
        List<string> presented = [];

        bool proceed = CommandPreconditions.TryRequireSingleBodyRow(
            Labels.Selection(),
            presented.Add,
            DescriptionLabelRefusalTranslator.TranslateNotSingleRow);

        Assert.True(proceed);
        Assert.Empty(presented);
    }

    [Fact]
    public void A_multi_row_selection_is_refused_with_the_commands_own_wording()
    {
        SelectionFacts facts = new(
            HasSelection: true,
            IsSingleBodyRow: false,
            BodyIndex: null,
            IsHeaderRow: false,
            SelectedRowCount: 3);
        List<string> presented = [];

        bool proceed = CommandPreconditions.TryRequireSingleBodyRow(
            Labels.Selection(facts),
            presented.Add,
            DescriptionLabelRefusalTranslator.TranslateNotSingleRow);

        Assert.False(proceed);
        Assert.Equal(
            DescriptionLabelRefusalTranslator.TranslateNotSingleRow(facts),
            Assert.Single(presented));
    }

    [Fact]
    public void A_missing_selection_is_refused()
    {
        List<string> presented = [];

        bool proceed = CommandPreconditions.TryRequireSingleBodyRow(
            Labels.Selection(SelectionFacts.None),
            presented.Add,
            StyleKeyRefusalTranslator.TranslateNotSingleRow);

        Assert.False(proceed);
        _ = Assert.Single(presented);
    }

    [Fact]
    public void Null_arguments_are_rejected()
    {
        _ = Assert.Throws<ArgumentNullException>(() => _ = CommandPreconditions.TryRequireSingleBodyRow(
            null!,
            _ => { },
            _ => "refused"));
        _ = Assert.Throws<ArgumentNullException>(() => _ = CommandPreconditions.TryRequireSingleBodyRow(
            Labels.Selection(),
            null!,
            _ => "refused"));
        _ = Assert.Throws<ArgumentNullException>(() => _ = CommandPreconditions.TryRequireSingleBodyRow(
            Labels.Selection(),
            _ => { },
            null!));
    }

    [Fact]
    public void A_presenter_failure_does_not_escape()
    {
        // The command boundary absorbs a dead dialog, exactly as it does for every other
        // refusal message the setter commands present.
        bool proceed = CommandPreconditions.TryRequireSingleBodyRow(
            Labels.Selection(SelectionFacts.None),
            _ => throw new InvalidOperationException("the dialog is gone"),
            DateLabelRefusalTranslator.TranslateNotSingleRow);

        Assert.False(proceed);
    }
}
