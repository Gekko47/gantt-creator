using GanttCreator.Core;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Sets the selected row's <c>DateLabelPosition</c> cell (R5.14: table
/// write, never shape edit; bespoke per entity, so the command re-reads
/// the selection at execution and refuses anything that is not exactly one
/// body row). The command owns outcome-to-message translation; selection
/// resolution and the cell write belong to the Office ports.
/// </summary>
internal static class SetDateLabelCommand
{
    /// <summary>Runs the production date-label command in the current Excel session.</summary>
    /// <param name="position">The choice to write, or null to clear the cell (Auto).</param>
    internal static void RunForExcel(GanttDateLabelPosition? position) =>
        Run(
            new ExcelDateLabelPositionMutator(ExcelDna.Integration.ExcelDnaUtil.Application),
            new ExcelSelectionIdentifier(ExcelDna.Integration.ExcelDnaUtil.Application),
            CommandErrorDialog.Show,
            position);

    /// <summary>Runs an injected date-label command.</summary>
    /// <param name="mutator">The date-label mutation port.</param>
    /// <param name="selection">The standalone selection identifier.</param>
    /// <param name="presenter">The user-safe refusal presenter.</param>
    /// <param name="position">The choice to write, or null to clear the cell (Auto).</param>
    internal static void Run(
        IDateLabelPositionMutator mutator,
        ISelectionIdentifier selection,
        Action<string> presenter,
        GanttDateLabelPosition? position)
    {
        ArgumentNullException.ThrowIfNull(mutator);
        ArgumentNullException.ThrowIfNull(selection);
        ArgumentNullException.ThrowIfNull(presenter);

        // The ribbon gate may have been stale, so the selection is re-read and a
        // non-single-row selection refuses before the mutator runs; the translator
        // keeps this command's own wording.
        if (!CommandPreconditions.TryRequireSingleBodyRow(
                selection,
                presenter,
                DateLabelRefusalTranslator.TranslateNotSingleRow))
        {
            return;
        }

        DateLabelMutationOutcome outcome = mutator.SetDateLabelPosition(position);

        // R2.2/R2.6: success is silent; the post-command ribbon refresh
        // handles any enabled-state consequences (R1.5 D3).
        if (outcome.Succeeded)
        {
            return;
        }

        CommandBoundary.SafePresent(presenter, DateLabelRefusalTranslator.Translate(outcome));
    }

}
