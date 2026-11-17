using System.Globalization;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Pure, UI-free translation of a style-key mutation outcome (or a selection
/// refusal) into the single simple-English message the command presents
/// (R5.6a; the <see cref="DescriptionLabelRefusalTranslator"/> shape). Every
/// planner refusal names the row it refused on using the outcome's one-based
/// body index; infrastructure and selection refusals name the situation
/// instead.
/// </summary>
internal static class StyleKeyRefusalTranslator
{
    /// <summary>Translates a style-key outcome to its user-safe message.</summary>
    /// <param name="outcome">The typed mutation outcome.</param>
    /// <returns>The actionable message.</returns>
    internal static string Translate(StyleKeyMutationOutcome outcome)
    {
        ArgumentNullException.ThrowIfNull(outcome);

        if (outcome.Refusal is not StyleKeyMutationRefusalReason refusal)
        {
            return string.Empty;
        }

        string row = outcome.SelectedRowNumber is int number
            ? string.Create(CultureInfo.InvariantCulture, $"Row {number} ")
            : "The selected row ";

        return refusal switch
        {
            StyleKeyMutationRefusalReason.NoActiveWorkbook =>
                "Gantt Creator needs an open workbook. Open or create a workbook, then try again.",
            StyleKeyMutationRefusalReason.TableMissing =>
                "Gantt Creator could not find tblGanttData. Run Initialise sheet before changing a style.",
            StyleKeyMutationRefusalReason.TargetProtected =>
                "The worksheet or workbook is protected, so nothing was changed. Remove protection and try again.",
            StyleKeyMutationRefusalReason.SelectionOutsideTable =>
                "No Gantt row is selected. Click one cell inside the Gantt data table, then try again.",
            StyleKeyMutationRefusalReason.SelectionUnknown =>
                "Gantt Creator could not work out which row you selected. Click one cell inside the Gantt data table, then try again.",
            StyleKeyMutationRefusalReason.TableInvalid =>
                "The Gantt data table or the style catalogue could not be read, so nothing was changed. "
                + "Run Validate to see which rows are wrong, fix them, then try again.",
            StyleKeyMutationRefusalReason.TypeUnknown =>
                row + "has a Type this version does not recognise, so nothing was changed. "
                + "Run Validate to see which rows are wrong, fix them, then try again.",
            StyleKeyMutationRefusalReason.TypeNotCustomStyled =>
                row + "does not carry a custom style, so nothing was changed. "
                + "Custom styles apply to Custom Activity and Custom Milestone rows.",
            StyleKeyMutationRefusalReason.StyleUnknown =>
                row + "could not be given that style, because the style was not found in the "
                + "configuration. Create it in Theme settings (Custom styles tab), then try again.",
            StyleKeyMutationRefusalReason.WriteFailed =>
                "Excel refused a write, so the table may be out of order. Refresh the chart to repair it.",
            _ => "Gantt Creator could not change the custom style. "
                + "Try again; if it keeps failing, see the Diagnostics dialog.",
        };
    }

    /// <summary>
    /// Translates the execution-time selection refusal (R5.6a ruling 1: the
    /// style assignment is bespoke per entity, so the command re-checks the
    /// selection and refuses anything that is not exactly one body row).
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
