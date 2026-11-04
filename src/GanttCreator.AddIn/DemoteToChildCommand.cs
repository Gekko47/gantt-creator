using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Demotes the top-level row under the active cell to a child of the
/// nearest eligible parent above it (R5.3 D3, amended 2026-10-11): the
/// row's content moves to sit directly beneath the new parent's child
/// run, and its existing children are re-homed to the new parent. The
/// command owns outcome-to-message translation; the selection
/// resolution, planning and cell writes belong to the Office mutator
/// port, exactly as the add-row command delegates row discovery to the
/// inserter port.
/// </summary>
internal static class DemoteToChildCommand
{
    /// <summary>Runs the production demote command in the current Excel session.</summary>
    internal static void RunForExcel()
    {
        // ADR-0055: this command changes row structure or shapes, which a protected sheet
        // refuses (measured), so it runs inside the release window. A render it triggers nests
        // inside the same window and only the outermost scope re-protects.
        using IDisposable window = RefreshSheetCommand.OpenProtectionWindow();

        Run(
            new ExcelParentChildMutator(
                ExcelDna.Integration.ExcelDnaUtil.Application,
                outlineGroups: new ExcelOutlineGroupWriter(ExcelDna.Integration.ExcelDnaUtil.Application)),
            new ExcelInsertedRowSelector(ExcelDna.Integration.ExcelDnaUtil.Application),
            CommandErrorDialog.Show);
    }

    /// <summary>Runs an injected demote command.</summary>
    /// <param name="mutator">The parent/child mutation port.</param>
    /// <param name="selector">Moves the selection onto the demoted row's cell.</param>
    /// <param name="presenter">The user-safe refusal presenter.</param>
    internal static void Run(
        IParentChildMutator mutator,
        IInsertedRowSelector selector,
        Action<string> presenter)
    {
        ArgumentNullException.ThrowIfNull(mutator);
        ArgumentNullException.ThrowIfNull(selector);
        ArgumentNullException.ThrowIfNull(presenter);

        ParentChildMutationOutcome outcome = mutator.Demote();

        // R2.2/R2.6: success is silent; the ribbon state refresh handles
        // any enabled-state consequences (R1.5 D3).
        if (outcome.Succeeded)
        {
            // Demote RELOCATES the row's content beneath the new parent's child run, so
            // the address the user was on now holds a different activity. The selection
            // follows the demoted row, in the column the user was working in; a host that
            // refuses leaves the cursor where it is, exactly as the move commands degrade.
            if (outcome.MovedToRowNumber is { } movedTo)
            {
                SelectMovedCell(selector, movedTo);
            }

            return;
        }

        // CA1031: presenter failure is outside command behaviour; it degrades to no dialog.
        Present(presenter, ParentChildRefusalTranslator.TranslateDemote(outcome));
    }

    /// <summary>
    /// Moves the selection onto the demoted row's cell in the user's column, degrading
    /// to no selection change when the host refuses.
    /// </summary>
    /// <param name="selector">The selection port.</param>
    /// <param name="bodyIndex">The demoted row's new one-based body index.</param>
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
