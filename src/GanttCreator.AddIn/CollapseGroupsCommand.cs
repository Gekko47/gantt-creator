using GanttCreator.Core;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Collapses or expands every parent/child group in the Gantt table (R5.4).
/// The command owns outcome-to-message translation; the plan and the row
/// writes belong to the Office outline port, exactly as the promote, demote
/// and move commands delegate theirs.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why the command drives the native state.</b> The user can already
/// collapse and expand a single group through Excel's own <c>-</c>/<c>+</c>
/// controls, and the owner has ruled that the visual and the function of that
/// behaviour are as intended. The Ribbon button adds the missing surface: one
/// action for every relationship at once. It writes the same state the native
/// control writes -- the child rows' hidden flag -- so the result is
/// indistinguishable from pressing the native control on each group.
/// </para>
/// <para>
/// <b>Why a successful request is silent.</b> R2.2 and R2.6 keep success
/// silent, and a sheet with no group is a no-op the user asked for rather than
/// a failure: there was nothing to collapse, and the state they wanted already
/// holds. Only a refusal is presented, because a refusal means the request
/// did not happen.
/// </para>
/// <para>
/// Successful outline changes schedule a coalesced refresh because visible
/// rows affect the rendered scene. A no-op does not schedule work.
/// </para>
/// </remarks>
internal static class CollapseGroupsCommand
{
    /// <summary>Runs the production collapse-all command in the current Excel session.</summary>
    internal static void RunForExcel(OutlineCollapseDirection direction)
        => Run(
            new ExcelOutlineGroupWriter(ExcelDna.Integration.ExcelDnaUtil.Application),
            direction,
            CommandErrorDialog.Show,
            RibbonStateService.Instance.ScheduleAutomaticRefresh);

    /// <summary>Runs an injected collapse-all command.</summary>
    /// <param name="outline">The outline port that performs the writes.</param>
    /// <param name="direction">Whether the groups are being collapsed or expanded.</param>
    /// <param name="presenter">The user-safe refusal presenter.</param>
    /// <param name="onChanged">Runs once after a successful change that wrote rows.</param>
    internal static void Run(
        IOutlineGroupPort outline,
        OutlineCollapseDirection direction,
        Action<string> presenter,
        Action? onChanged = null)
    {
        ArgumentNullException.ThrowIfNull(outline);
        ArgumentNullException.ThrowIfNull(presenter);

        OutlineCollapseOutcome outcome = outline.SetCollapsed(direction);

        // R2.2/R2.6: success is silent, including the no-op case where the
        // sheet carries no group.
        if (outcome.Succeeded)
        {
            if (outcome.RowsWritten > 0)
            {
                onChanged?.Invoke();
            }
            return;
        }

        // CA1031: presenter failure is outside command behaviour; it degrades to no dialog.
        Present(presenter, Translate(outcome, direction));
    }

    /// <summary>
    /// Translates a refusal into the single simple-English message the command
    /// presents.
    /// </summary>
    /// <param name="outcome">The typed outcome.</param>
    /// <param name="direction">The direction the request was moving.</param>
    /// <returns>The actionable message.</returns>
    private static string Translate(OutlineCollapseOutcome outcome, OutlineCollapseDirection direction)
    {
        ArgumentNullException.ThrowIfNull(outcome);

        string action = direction == OutlineCollapseDirection.Collapse
            ? "collapsed"
            : "expanded";

        return outcome.Refusal switch
        {
            OutlineCollapseRefusalReason.NoActiveWorkbook =>
                "Gantt Creator could not find an open workbook, so the groups were not " + action + ".",
            OutlineCollapseRefusalReason.TableMissing =>
                "Gantt Creator could not find the Gantt data table, so the groups were not " + action
                + ". Run Initialise sheet, then try again.",
            OutlineCollapseRefusalReason.TargetProtected =>
                "The worksheet is protected, so the groups were not " + action
                + ". Remove the protection, then try again.",
            _ => "Gantt Creator could not " + action + " the groups.",
        };
    }

    /// <summary>
    /// Shows a message, degrading to silence when the host refuses to present
    /// it. A dialog failure must never turn a completed command into an
    /// escaping exception.
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
