using System.Globalization;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Pure, UI-free translation of a date-label mutation outcome (or a
/// selection refusal) into the single simple-English message the command
/// presents (R5.14; the <see cref="DescriptionLabelRefusalTranslator"/>
/// shape). Every refusal names the row it refused on using the outcome's
/// one-based body index; infrastructure and selection refusals name the
/// situation instead.
/// </summary>
internal static class DateLabelRefusalTranslator
{
    /// <summary>Translates a date-label outcome to its user-safe message.</summary>
    /// <param name="outcome">The typed mutation outcome.</param>
    /// <returns>The actionable message.</returns>
    internal static string Translate(DateLabelMutationOutcome outcome)
    {
        ArgumentNullException.ThrowIfNull(outcome);

        if (outcome.Refusal is not DateLabelMutationRefusalReason refusal)
        {
            return string.Empty;
        }

        var row = outcome.SelectedRowNumber is int number
            ? string.Create(CultureInfo.InvariantCulture, $"Row {number} ")
            : "The selected row ";

        return refusal switch
        {
            DateLabelMutationRefusalReason.NoActiveWorkbook =>
                "Gantt Creator needs an open workbook. Open or create a workbook, then try again.",
            DateLabelMutationRefusalReason.TableMissing =>
                "Gantt Creator could not find tblGanttData. Run Initialise sheet before changing a date label.",
            DateLabelMutationRefusalReason.TargetProtected =>
                "The worksheet or workbook is protected, so nothing was changed. Remove protection and try again.",
            DateLabelMutationRefusalReason.SelectionOutsideTable =>
                "No Gantt row is selected. Click one cell inside the Gantt data table, then try again.",
            DateLabelMutationRefusalReason.SelectionUnknown =>
                "Gantt Creator could not work out which row you selected. Click one cell inside the Gantt data table, then try again.",
            DateLabelMutationRefusalReason.TypeUnknown =>
                row + "has a Type this version does not recognise, so nothing was changed. "
                + "Run Validate to see which rows are wrong, fix them, then try again.",
            DateLabelMutationRefusalReason.DateLabelNotAllowedForType =>
                row + "does not support that date label, so nothing was changed. "
                + "Choose one of the date labels this row's type offers.",
            DateLabelMutationRefusalReason.WriteFailed =>
                "Excel refused a write, so the table may be out of order. Refresh the chart to repair it.",
            _ => "Gantt Creator could not change the date label. "
                + "Try again; if it keeps failing, see the Diagnostics dialog.",
        };
    }

    /// <summary>
    /// Translates the execution-time selection refusal (date labels are
    /// bespoke per entity, so the command re-checks the selection and
    /// refuses anything that is not exactly one body row).
    /// </summary>
    /// <param name="facts">The selection facts captured at execution time.</param>
    /// <returns>The actionable message.</returns>
    internal static string TranslateNotSingleRow(SelectionFacts facts)
    {
        ArgumentNullException.ThrowIfNull(facts);

        return facts.IsHeaderRow
            ? "Click a cell on a data row (not the table header), then try again."
            : facts.SelectedRowCount > 1
            ? "Gantt Creator changes one row at a time. Select exactly one row, then try again."
            : "No Gantt row is selected. Click one cell inside the Gantt data table, then try again.";
    }
}