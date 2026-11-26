using GanttCreator.Core;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Sets the selected row's <c>LabelPosition</c> cell (R5.5 D2: table write,
/// never shape edit; bespoke per entity — ruling 1 — so the command re-reads
/// the selection at execution and refuses anything that is not exactly one
/// body row). The command owns outcome-to-message translation; selection
/// resolution and the cell write belong to the Office ports.
/// </summary>
internal static class SetDescriptionLabelCommand
{
    /// <summary>Runs the production description-label command in the current Excel session.</summary>
    /// <param name="position">The position to write, or null to clear the cell (Default).</param>
    internal static void RunForExcel(GanttLabelPosition? position) =>
        Run(
            new ExcelLabelPositionMutator(ExcelDna.Integration.ExcelDnaUtil.Application),
            new ExcelSelectionIdentifier(ExcelDna.Integration.ExcelDnaUtil.Application),
            CommandErrorDialog.Show,
            position);

    /// <summary>Runs an injected description-label command.</summary>
    /// <param name="mutator">The description-label mutation port.</param>
    /// <param name="selection">The standalone selection identifier (ruling 6).</param>
    /// <param name="presenter">The user-safe refusal presenter.</param>
    /// <param name="position">The position to write, or null to clear the cell (Default).</param>
    internal static void Run(
        ILabelPositionMutator mutator,
        ISelectionIdentifier selection,
        Action<string> presenter,
        GanttLabelPosition? position)
    {
        ArgumentNullException.ThrowIfNull(mutator);
        ArgumentNullException.ThrowIfNull(selection);
        ArgumentNullException.ThrowIfNull(presenter);

        // Ruling 1 at execution time: the ribbon gate may have been stale
        // (selection changed between render and click), so the selection is
        // re-read here and a non-single-row selection refuses before the
        // mutator runs.
        SelectionFacts facts = selection.Identify();
        if (!facts.IsSingleBodyRow)
        {
            CommandBoundary.SafePresent(presenter, DescriptionLabelRefusalTranslator.TranslateNotSingleRow(facts));
            return;
        }

        LabelPositionMutationOutcome outcome = mutator.SetLabelPosition(position);

        // R2.2/R2.6: success is silent; the post-command ribbon refresh
        // handles any enabled-state consequences (R1.5 D3).
        if (outcome.Succeeded)
        {
            return;
        }

        CommandBoundary.SafePresent(presenter, DescriptionLabelRefusalTranslator.Translate(outcome));
    }

}
