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
internal static class SetLabelPositionCommand
{
    /// <summary>Runs the production label-position command in the current Excel session.</summary>
    /// <param name="position">The position to write, or null to clear the cell (Default).</param>
    internal static void RunForExcel(GanttLabelPosition? position) =>
        Run(
            new ExcelLabelPositionMutator(ExcelDna.Integration.ExcelDnaUtil.Application),
            new ExcelSelectionIdentifier(ExcelDna.Integration.ExcelDnaUtil.Application),
            CommandErrorDialog.Show,
            position);

    /// <summary>Runs an injected label-position command.</summary>
    /// <param name="mutator">The label-position mutation port.</param>
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
            Present(presenter, LabelPositionRefusalTranslator.TranslateNotSingleRow(facts));
            return;
        }

        LabelPositionMutationOutcome outcome = mutator.SetLabelPosition(position);

        // R2.2/R2.6: success is silent; the post-command ribbon refresh
        // handles any enabled-state consequences (R1.5 D3).
        if (outcome.Succeeded)
        {
            return;
        }

        Present(presenter, LabelPositionRefusalTranslator.Translate(outcome));
    }

    /// <summary>
    /// Shows a message, degrading to silence when the host refuses to present
    /// it. A dialog failure must never turn a completed command into an
    /// escaping exception (the <see cref="PromoteToParentCommand"/> shape).
    /// </summary>
    /// <param name="presenter">The user-safe message presenter.</param>
    /// <param name="message">The message to present.</param>
    private static void Present(Action<string> presenter, string message)
    {
        // CA1031: presenter failure is outside command behaviour.
#pragma warning disable CA1031
        try
        {
            presenter(message);
        }
        catch
        {
            // Intentionally empty: a dialog failure must not escape into Excel.
        }
#pragma warning restore CA1031
    }
}
