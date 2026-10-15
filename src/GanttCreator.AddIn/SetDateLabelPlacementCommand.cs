using GanttCreator.Core;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Sets the selected row's <c>DateLabelPosition</c> cell as a
/// date-label <em>position</em> (owner ruling 2026-10-14: the
/// delineator's date label is positionable with the description
/// label's corner set, so the cell names a corner rather than a
/// which-date choice; table write, never shape edit; bespoke per
/// entity, so the command re-reads the selection at execution and
/// refuses anything that is not exactly one body row). The command
/// owns outcome-to-message translation; selection resolution and the
/// cell write belong to the Office ports.
/// </summary>
internal static class SetDateLabelPlacementCommand
{
    /// <summary>Runs the production date-label placement command in the current Excel session.</summary>
    /// <param name="placement">The corner to write, or null to clear the cell (Auto).</param>
    internal static void RunForExcel(GanttLabelPosition? placement) =>
        Run(
            new ExcelDateLabelPositionMutator(ExcelDna.Integration.ExcelDnaUtil.Application),
            new ExcelSelectionIdentifier(ExcelDna.Integration.ExcelDnaUtil.Application),
            CommandErrorDialog.Show,
            placement);

    /// <summary>Runs an injected date-label placement command.</summary>
    /// <param name="mutator">The date-label mutation port.</param>
    /// <param name="selection">The standalone selection identifier.</param>
    /// <param name="presenter">The user-safe refusal presenter.</param>
    /// <param name="placement">The corner to write, or null to clear the cell (Auto).</param>
    internal static void Run(
        IDateLabelPositionMutator mutator,
        ISelectionIdentifier selection,
        Action<string> presenter,
        GanttLabelPosition? placement)
    {
        ArgumentNullException.ThrowIfNull(mutator);
        ArgumentNullException.ThrowIfNull(selection);
        ArgumentNullException.ThrowIfNull(presenter);

        // The ribbon gate may have been stale (selection changed between
        // render and click), so the selection is re-read here and a
        // non-single-row selection refuses before the mutator runs.
        SelectionFacts facts = selection.Identify();
        if (!facts.IsSingleBodyRow)
        {
            Present(presenter, DateLabelRefusalTranslator.TranslateNotSingleRow(facts));
            return;
        }

        DateLabelMutationOutcome outcome = mutator.SetDateLabelPlacement(placement);

        // R2.2/R2.6: success is silent; the post-command ribbon refresh
        // handles any enabled-state consequences (R1.5 D3).
        if (outcome.Succeeded)
        {
            return;
        }

        Present(presenter, DateLabelRefusalTranslator.Translate(outcome));
    }

    /// <summary>Presents a refusal message; isolated for the same reason as the date-label command's.</summary>
    /// <param name="presenter">The user-safe presenter.</param>
    /// <param name="message">The translated refusal message.</param>
    private static void Present(Action<string> presenter, string message) => presenter(message);
}
