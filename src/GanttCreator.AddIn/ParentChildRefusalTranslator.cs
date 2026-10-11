using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Pure, UI-free translation of a parent/child mutation outcome into the
/// single simple-English message the command presents (R5.3 D6). Every
/// planner refusal names the row it refused on, using the outcome's
/// one-based body index; infrastructure refusals name the situation
/// instead, because they fire before the selection resolves to a row.
/// </summary>
/// <remarks>
/// The translation lives beside the commands rather than in
/// <see cref="CommandErrorTranslator"/>, which translates unexpected
/// exceptions at the boundary: a typed refusal is expected behaviour with
/// a precise remedy, not an error category. Both parent/child commands
/// share this one table so a message cannot drift between the verbs.
/// </remarks>
internal static class ParentChildRefusalTranslator
{
    /// <summary>Translates a promote outcome to its user-safe message.</summary>
    /// <param name="outcome">The typed mutation outcome.</param>
    /// <returns>The actionable message.</returns>
    internal static string TranslatePromote(ParentChildMutationOutcome outcome)
    {
        ArgumentNullException.ThrowIfNull(outcome);

        return outcome.Refusal is ParentChildMutationRefusalReason infrastructure
            ? TranslateInfrastructureRefusal(infrastructure)
            : TranslatePlannerRefusal(outcome.PlanRefusal, outcome.SelectedRowNumber, isPromote: true);
    }

    /// <summary>Translates a demote outcome to its user-safe message.</summary>
    /// <param name="outcome">The typed mutation outcome.</param>
    /// <returns>The actionable message.</returns>
    internal static string TranslateDemote(ParentChildMutationOutcome outcome)
    {
        ArgumentNullException.ThrowIfNull(outcome);

        return outcome.Refusal is ParentChildMutationRefusalReason infrastructure
            ? TranslateInfrastructureRefusal(infrastructure)
            : TranslatePlannerRefusal(outcome.PlanRefusal, outcome.SelectedRowNumber, isPromote: false);
    }

    /// <summary>
    /// Translates the adapter/host refusals, which fire before the
    /// selection resolves to a row and therefore carry no row number.
    /// </summary>
    /// <param name="refusal">The infrastructure refusal reason.</param>
    /// <returns>The actionable message.</returns>
    private static string TranslateInfrastructureRefusal(ParentChildMutationRefusalReason refusal) => refusal switch
    {
        ParentChildMutationRefusalReason.NoActiveWorkbook =>
            "Gantt Creator needs an open workbook. Open or create a workbook, then try again.",
        ParentChildMutationRefusalReason.TableMissing =>
            "Gantt Creator could not find tblGanttData. Run Initialise sheet before using the parent/child commands.",
        ParentChildMutationRefusalReason.TargetProtected =>
            "The worksheet or workbook is protected, so no row was changed. Remove protection and try again.",
        ParentChildMutationRefusalReason.SelectionOutsideTable =>
            "No Gantt row is selected. Click a cell inside the Gantt data table, then try again.",
        ParentChildMutationRefusalReason.SelectionUnknown =>
            "Gantt Creator could not work out which row you selected. Click a cell inside the Gantt data table, then try again.",
        ParentChildMutationRefusalReason.TableInvalid =>
            "The Gantt data table contains a row that cannot be read, so nothing was changed. "
            + "Run Validate to see which rows are wrong, fix them, then try again.",
        ParentChildMutationRefusalReason.WriteFailed =>
            "Excel refused a write, so the table may be out of order. Refresh the chart to repair it.",
        _ => "Gantt Creator could not change the table. Try again; if it keeps failing, see the Diagnostics dialog.",
    };


    /// <summary>
    /// Translates the Core planner's typed refusal, naming the row it
    /// refused on. Internal so the message table is testable without a
    /// full outcome.
    /// </summary>
    /// <param name="planRefusal">The planner's refusal reason, or null for an unknown enum value.</param>
    /// <param name="selectedRowNumber">The targeted row's one-based body index, when known.</param>
    /// <param name="isPromote">Whether the refused command was promote.</param>
    /// <returns>The actionable message.</returns>
    private static string TranslatePlannerRefusal(
        ParentChildRefusalReason? planRefusal,
        int? selectedRowNumber,
        bool isPromote)
    {
        var row = selectedRowNumber is int number
            ? string.Create(CultureInfo.InvariantCulture, $"Row {number} ")
            : "The selected row ";

        return planRefusal switch
        {
            ParentChildRefusalReason.SelectionUnknown =>
                "Gantt Creator could not find the selected row in the table. "
                + "Click a cell inside the Gantt data table, then try again.",
            ParentChildRefusalReason.NotAChild =>
                row + "is already a top-level activity, so there is nothing to promote.",
            ParentChildRefusalReason.AlreadyAChild =>
                row + "is already a child, so there is nothing to demote.",
            ParentChildRefusalReason.TypeCannotBeChild =>
                row + "cannot become a child: this row's type always stands on its own.",
            ParentChildRefusalReason.PromotionForbiddenForType =>
                row + "cannot be promoted: a critical interval is always a child of an activity.",
            ParentChildRefusalReason.HasChildren =>
                row + "cannot be promoted while it has children of its own. "
                + "Promote its children first, or delete them.",
            ParentChildRefusalReason.NoEligibleParentAbove =>
                row + "cannot become a child: there is no activity above it that can take a child.",
            ParentChildRefusalReason.ParentUnavailable =>
                row + "cannot become a child: the activities above it are already full "
                + "(seven children each). Delete or promote one of their children, then try again.",
            ParentChildRefusalReason.TooDeep =>
                row + "cannot become a child: it would become a grandchild, and Gantt Creator "
                + "supports one level of children.",
            ParentChildRefusalReason.Cycle =>
                row + "cannot become a child: the change would create a loop in the parent links.",
            ParentChildRefusalReason.ParentNotSpan =>
                row + "cannot become a child of that activity: a critical interval may only be a "
                + "child of a planned span (an activity with both start and finish dates).",
            _ => isPromote
                ? row + "could not be promoted. Run Validate to check the table, then try again."
                : row + "could not be demoted. Run Validate to check the table, then try again.",
        };
    }
}
