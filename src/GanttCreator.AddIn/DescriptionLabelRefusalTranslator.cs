using System.Globalization;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Pure, UI-free translation of a description-label mutation outcome (or a
/// selection refusal) into the single simple-English message the command
/// presents (R5.5 D2; the <see cref="ParentChildRefusalTranslator"/> shape).
/// Every planner refusal names the row it refused on using the outcome's
/// one-based body index; infrastructure and selection refusals name the
/// situation instead.
/// </summary>
internal static class DescriptionLabelRefusalTranslator
{
    /// <summary>Translates a description-label outcome to its user-safe message.</summary>
    /// <param name="outcome">The typed mutation outcome.</param>
    /// <returns>The actionable message.</returns>
    internal static string Translate(LabelPositionMutationOutcome outcome)
    {
        ArgumentNullException.ThrowIfNull(outcome);

        if (outcome.Refusal is not LabelPositionMutationRefusalReason refusal)
        {
            return string.Empty;
        }

        string row = outcome.SelectedRowNumber is int number
            ? string.Create(CultureInfo.InvariantCulture, $"Row {number} ")
            : "The selected row ";

        return refusal switch
        {
            LabelPositionMutationRefusalReason.NoActiveWorkbook =>
                "Gantt Creator needs an open workbook. Open or create a workbook, then try again.",
            LabelPositionMutationRefusalReason.TableMissing =>
                "Gantt Creator could not find tblGanttData. Run Initialise sheet before changing a label.",
            LabelPositionMutationRefusalReason.TargetProtected =>
                "The worksheet or workbook is protected, so nothing was changed. Remove protection and try again.",
            LabelPositionMutationRefusalReason.SelectionOutsideTable =>
                "No Gantt row is selected. Click one cell inside the Gantt data table, then try again.",
            LabelPositionMutationRefusalReason.SelectionUnknown =>
                "Gantt Creator could not work out which row you selected. Click one cell inside the Gantt data table, then try again.",
            LabelPositionMutationRefusalReason.TableInvalid =>
                "The Gantt data table contains a row that cannot be read, so nothing was changed. "
                + "Run Validate to see which rows are wrong, fix them, then try again.",
            LabelPositionMutationRefusalReason.TypeUnknown =>
                row + "has a Type this version does not recognise, so nothing was changed. "
                + "Run Validate to see which rows are wrong, fix them, then try again.",
            LabelPositionMutationRefusalReason.PositionNotAllowedForType =>
                row + "does not support that label position, so nothing was changed. "
                + "Choose one of the positions this row's type offers.",
            LabelPositionMutationRefusalReason.StyleUnknown =>
                row + "is a Custom Activity whose style could not be found, so nothing was changed. "
                + "Choose a StyleKey for the row (or run Repair configuration), then try again.",
            LabelPositionMutationRefusalReason.WriteFailed =>
                "Excel refused a write, so the table may be out of order. Refresh the chart to repair it.",
            _ => "Gantt Creator could not change the label position. "
                + "Try again; if it keeps failing, see the Diagnostics dialog.",
        };
    }

    /// <summary>
    /// Translates the execution-time selection refusal (R5.5 ruling 1: labels
    /// are bespoke per entity, so the command re-checks the selection and
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
