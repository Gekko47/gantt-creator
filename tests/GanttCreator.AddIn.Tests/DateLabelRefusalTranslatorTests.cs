using GanttCreator.Office;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contract tests for <see cref="DateLabelRefusalTranslator"/>: the pure
/// outcome-to-message mapping every refusal arm reaches, the row-scoped wording,
/// and the selection refusal.
/// </summary>
public sealed class DateLabelRefusalTranslatorTests
{
    /// <summary>Every defined refusal reason, so a new arm cannot ship untested.</summary>
    public static TheoryData<DateLabelMutationRefusalReason> Refusals
    {
        get
        {
            TheoryData<DateLabelMutationRefusalReason> data = [];
            foreach (DateLabelMutationRefusalReason reason in
                     Enum.GetValues<DateLabelMutationRefusalReason>())
            {
                data.Add(reason);
            }

            return data;
        }
    }

    [Fact]
    public void A_successful_outcome_has_nothing_to_report() =>
        Assert.Equal(
            string.Empty,
            DateLabelRefusalTranslator.Translate(DateLabelMutationOutcome.Ok(1, 4)));

    [Fact]
    public void A_null_outcome_is_a_programming_error() =>
        _ = Assert.Throws<ArgumentNullException>(() => DateLabelRefusalTranslator.Translate(null!));

    [Theory]
    [MemberData(nameof(Refusals))]
    public void Every_refusal_has_an_actionable_message(DateLabelMutationRefusalReason refusal)
    {
        string message = DateLabelRefusalTranslator.Translate(
            DateLabelMutationOutcome.Refused(refusal, 3));

        Assert.False(string.IsNullOrWhiteSpace(message));
        Assert.EndsWith(".", message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_missing_table_names_the_situation_rather_than_a_row() =>
        Assert.Equal(
            "Gantt Creator could not find tblGanttData. Run Initialise sheet before changing a date label.",
            DateLabelRefusalTranslator.Translate(
                DateLabelMutationOutcome.Refused(DateLabelMutationRefusalReason.TableMissing)));

    [Fact]
    public void A_row_scoped_refusal_names_the_row_it_refused_on() =>
        Assert.Equal(
            "Row 7 does not support that date label, so nothing was changed. "
            + "Choose one of the date labels this row's type offers.",
            DateLabelRefusalTranslator.Translate(DateLabelMutationOutcome.Refused(
                DateLabelMutationRefusalReason.DateLabelNotAllowedForType, 7)));

    [Fact]
    public void A_row_scoped_refusal_without_a_known_row_names_the_selection() =>
        Assert.StartsWith(
            "The selected row ",
            DateLabelRefusalTranslator.Translate(DateLabelMutationOutcome.Refused(
                DateLabelMutationRefusalReason.TypeUnknown)),
            StringComparison.Ordinal);

    [Fact]
    public void An_undefined_refusal_falls_back_to_the_diagnostics_message()
    {
        string message = DateLabelRefusalTranslator.Translate(
            DateLabelMutationOutcome.Refused((DateLabelMutationRefusalReason)99, 3));

        Assert.Contains("could not change the date label", message, StringComparison.Ordinal);
        Assert.DoesNotContain("Row ", message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_header_row_selection_asks_for_a_data_row() =>
        Assert.Equal(
            "Click a cell on a data row (not the table header), then try again.",
            DateLabelRefusalTranslator.TranslateNotSingleRow(
                new SelectionFacts(true, false, null, true, 0)));

    [Fact]
    public void A_multi_row_selection_asks_for_exactly_one_row() =>
        Assert.Equal(
            "Gantt Creator changes one row at a time. Select exactly one row, then try again.",
            DateLabelRefusalTranslator.TranslateNotSingleRow(
                new SelectionFacts(true, false, null, false, 3)));

    [Fact]
    public void No_selection_asks_for_a_cell_inside_the_table() =>
        Assert.Equal(
            "No Gantt row is selected. Click one cell inside the Gantt data table, then try again.",
            DateLabelRefusalTranslator.TranslateNotSingleRow(SelectionFacts.None));

    [Fact]
    public void A_null_selection_is_a_programming_error() =>
        _ = Assert.Throws<ArgumentNullException>(
            () => DateLabelRefusalTranslator.TranslateNotSingleRow(null!));
}
