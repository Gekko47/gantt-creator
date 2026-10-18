using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Sets the selected row's <c>StyleKey</c> cell (R5.6a: table write, never
/// shape edit; bespoke per entity — ruling 1 — so the command re-reads the
/// selection at execution and refuses anything that is not exactly one body
/// row). The command owns outcome-to-message translation; selection
/// resolution and the cell write belong to the Office ports.
/// </summary>
internal static class SetStyleKeyCommand
{
    /// <summary>Runs the production style-key command in the current Excel session.</summary>
    /// <param name="styleKey">The exact style key to assign to the selected row.</param>
    internal static void RunForExcel(string styleKey) =>
        Run(
            new ExcelStyleKeyMutator(ExcelDna.Integration.ExcelDnaUtil.Application),
            new ExcelSelectionIdentifier(ExcelDna.Integration.ExcelDnaUtil.Application),
            CommandErrorDialog.Show,
            styleKey);

    /// <summary>Runs an injected style-key command.</summary>
    /// <param name="mutator">The style-key mutation port.</param>
    /// <param name="selection">The standalone selection identifier (ruling 1).</param>
    /// <param name="presenter">The user-safe refusal presenter.</param>
    /// <param name="styleKey">The exact style key to assign to the selected row.</param>
    internal static void Run(
        IStyleKeyMutator mutator,
        ISelectionIdentifier selection,
        Action<string> presenter,
        string styleKey)
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
            Present(presenter, StyleKeyRefusalTranslator.TranslateNotSingleRow(facts));
            return;
        }

        StyleKeyMutationOutcome outcome = mutator.SetStyleKey(styleKey);

        // R2.2/R2.6: success is silent; the post-command ribbon refresh
        // handles any enabled-state consequences (R1.5 D3).
        if (outcome.Succeeded)
        {
            return;
        }

        Present(presenter, StyleKeyRefusalTranslator.Translate(outcome));
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
