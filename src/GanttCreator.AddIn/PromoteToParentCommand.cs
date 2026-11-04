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
            new ExcelInsertedRowSelector(ExcelDna.Integration.ExcelDnaUtil.Application),
            CommandErrorDialog.Show);

    /// <summary>Runs an injected promote command.</summary>
    /// <param name="mutator">The parent/child mutation port.</param>
    /// <param name="selector">Moves the selection onto the promoted row's cell.</param>
    /// <param name="presenter">The user-safe refusal presenter.</param>
    internal static void Run(
        IParentChildMutator mutator,
        IInsertedRowSelector selector,
        Action<string> presenter)
    {
        ArgumentNullException.ThrowIfNull(mutator);
        ArgumentNullException.ThrowIfNull(selector);
        ArgumentNullException.ThrowIfNull(presenter);

        ParentChildMutationOutcome outcome = mutator.Promote();

        // R2.2/R2.6: success is silent; the ribbon state refresh handles
        // any enabled-state consequences (R1.5 D3).
        if (outcome.Succeeded)
        {
            // Promote RELOCATES the row's content to the bottom of its sibling run, so
            // the address the user was on now holds a different activity. The selection
            // follows the promoted row, in the column the user was working in; a host
            // that refuses leaves the cursor where it is, exactly as the move commands
            // degrade.
            if (outcome.MovedToRowNumber is { } movedTo)
            {
                SelectMovedCell(selector, movedTo);
            }

            return;
        }

        // CA1031: presenter failure is outside command behaviour; it degrades to no dialog.
        Present(presenter, ParentChildRefusalTranslator.TranslatePromote(outcome));
    }

    /// <summary>
    /// Moves the selection onto the promoted row's cell in the user's column, degrading
    /// to no selection change when the host refuses.
    /// </summary>
    /// <param name="selector">The selection port.</param>
    /// <param name="bodyIndex">The promoted row's new one-based body index.</param>
    private static void SelectMovedCell(IInsertedRowSelector selector, int bodyIndex)
    {
        // CA1031: a host selection failure is outside command behaviour.
#pragma warning disable CA1031
        try
        {
            selector.SelectBodyCell(bodyIndex);
        }
        catch
        {
            // Intentionally empty: the row moved; only the cursor stays put.
        }
#pragma warning restore CA1031
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
