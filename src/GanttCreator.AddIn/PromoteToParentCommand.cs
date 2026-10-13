using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Promotes the child row under the active cell to top-level (R5.3 D2).
/// The command owns outcome-to-message translation; the selection
/// resolution, planning and cell writes belong to the Office mutator
/// port, exactly as the add-row command delegates row discovery to the
/// inserter port.
/// </summary>
internal static class PromoteToParentCommand
{
    /// <summary>Runs the production promote command in the current Excel session.</summary>
    internal static void RunForExcel() =>
        Run(
            new ExcelParentChildMutator(
                ExcelDna.Integration.ExcelDnaUtil.Application,
                outlineGroups: new ExcelOutlineGroupWriter(ExcelDna.Integration.ExcelDnaUtil.Application)),
            CommandErrorDialog.Show);

    /// <summary>Runs an injected promote command.</summary>
    /// <param name="mutator">The parent/child mutation port.</param>
    /// <param name="presenter">The user-safe refusal presenter.</param>
    internal static void Run(IParentChildMutator mutator, Action<string> presenter)
    {
        ArgumentNullException.ThrowIfNull(mutator);
        ArgumentNullException.ThrowIfNull(presenter);

        ParentChildMutationOutcome outcome = mutator.Promote();

        // R2.2/R2.6: success is silent; the ribbon state refresh handles
        // any enabled-state consequences (R1.5 D3).
        if (outcome.Succeeded)
        {
            return;
        }

        // CA1031: presenter failure is outside command behaviour; it degrades to no dialog.
        Present(presenter, ParentChildRefusalTranslator.TranslatePromote(outcome));
    }

    /// <summary>
    /// Shows a message, degrading to silence when the host refuses to present it. A
    /// dialog failure must never turn a completed command into an escaping exception.
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
