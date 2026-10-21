using GanttCreator.Core;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Adds one new row as a CHILD of the row under the active cell (R5.3 WI-E).
/// The command owns outcome-to-message translation; the selection
/// resolution, planning and cell writes belong to the Office mutator port,
/// exactly as the demote command delegates its planning and writes.
/// </summary>
internal static class AddChildCommand
{
    /// <summary>Runs the production add-child command in the current Excel session.</summary>
    /// <param name="childType">The catalogue type the new child will have.</param>
    internal static void RunForExcel(GanttEntityType childType) =>
        Run(
            new ExcelAddChildMutator(
                ExcelDna.Integration.ExcelDnaUtil.Application,
                rowInserter: new ExcelGanttRowInserter(
                    ExcelDna.Integration.ExcelDnaUtil.Application,
                    outlineGroups: new ExcelOutlineGroupWriter(ExcelDna.Integration.ExcelDnaUtil.Application)),
                outlineGroups: new ExcelOutlineGroupWriter(ExcelDna.Integration.ExcelDnaUtil.Application)),
            GanttRowId.New,
            CommandErrorDialog.Show,
            childType);

    /// <summary>Runs an injected add-child command.</summary>
    /// <param name="mutator">The add-child mutation port.</param>
    /// <param name="nextId">The stable row-ID generator.</param>
    /// <param name="presenter">The user-safe refusal presenter.</param>
    /// <param name="childType">The catalogue type the new child will have.</param>
    internal static void Run(
        IAddChildMutator mutator,
        Func<GanttRowId> nextId,
        Action<string> presenter,
        GanttEntityType childType)
    {
        ArgumentNullException.ThrowIfNull(mutator);
        ArgumentNullException.ThrowIfNull(nextId);
        ArgumentNullException.ThrowIfNull(presenter);

        AddChildMutationOutcome outcome = mutator.AddChild(childType, nextId);

        // R2.2/R2.6: success is silent; the ribbon state refresh handles any
        // enabled-state consequences (R1.5 D3).
        if (outcome.Succeeded)
        {
            return;
        }

        // CA1031: presenter failure is outside command behaviour; it degrades to no dialog.
        Present(presenter, Translate(outcome));
    }

    /// <summary>Translates an add-child refusal to a user-safe message.</summary>
    /// <param name="outcome">The typed refusal outcome.</param>
    /// <returns>The actionable message.</returns>
    internal static string Translate(AddChildMutationOutcome outcome) => outcome.PlanRefusal switch
    {
        AddChildRefusalReason.SelectionUnknown =>
            "Gantt Creator could not find the row to add a child to. Select the parent row, then try again.",
        AddChildRefusalReason.ParentCannotOwnChildren =>
            "That row cannot have children. Select a row that can own children -- an activity, a procurement or a delay event -- then try again.",
        AddChildRefusalReason.TypeCannotBeChild =>
            "That type cannot be a child. Choose a type that can appear under a parent.",
        AddChildRefusalReason.ParentAtCapacity =>
            "That row already has the maximum number of children. Promote or remove one of them first.",
        AddChildRefusalReason.TooDeep =>
            "That row is already a child, so a child of it would be three levels deep. Add the child to its parent instead.",
        _ => outcome.Refusal switch
        {
            AddChildMutationRefusalReason.NoActiveWorkbook =>
                "Gantt Creator needs an open workbook. Open or create a workbook, then try adding the child again.",
            AddChildMutationRefusalReason.TableMissing =>
                "Gantt Creator could not find tblGanttData. Run Initialise sheet before adding a child.",
            AddChildMutationRefusalReason.TargetProtected =>
                "The worksheet or workbook is protected, so no child was added. Remove protection and try again.",
            AddChildMutationRefusalReason.SelectionOutsideTable =>
                "Select the parent row in the Gantt table, then add the child.",
            AddChildMutationRefusalReason.TableInvalid =>
                "Gantt Creator could not read the table, so no child was added. Repair the configuration and try again.",
            AddChildMutationRefusalReason.InsertRefused =>
                "Gantt Creator could not add the row, so no child was added. Try again; if it keeps failing, see the Diagnostics dialog.",
            AddChildMutationRefusalReason.WriteFailed =>
                "Gantt Creator added the row, but Excel refused a follow-up write, so the row may be incomplete. Delete it and add it again.",
            _ => "Gantt Creator could not add the child. Try again; if it keeps failing, see the Diagnostics dialog.",
        },
    };

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
