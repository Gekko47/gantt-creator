using GanttCreator.Core;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Adds one scaffold row to the visible Gantt data table. The command owns
/// application-level error translation; row discovery and mutation belong to
/// the Office inserter port.
/// </summary>
internal static class AddRowCommand
{
    /// <summary>Runs the production add-row command in the current Excel session.</summary>
    /// <param name="type">The canonical catalogue type to insert.</param>
    internal static void RunForExcel(GanttEntityType type) =>
        Run(
            new ExcelGanttRowInserter(ExcelDna.Integration.ExcelDnaUtil.Application),
            GanttRowId.New,
            new ExcelInsertedRowSelector(ExcelDna.Integration.ExcelDnaUtil.Application),
            CommandErrorDialog.Show,
            type);

    /// <summary>Runs an injected add-row command.</summary>
    /// <param name="inserter">The row insertion port.</param>
    /// <param name="nextId">The stable row-ID generator.</param>
    /// <param name="selector">Moves the selection onto the inserted row.</param>
    /// <param name="presenter">The user-safe refusal presenter.</param>
    /// <param name="type">The canonical catalogue type to insert.</param>
    internal static void Run(
        IGanttRowInserter inserter,
        Func<GanttRowId> nextId,
        IInsertedRowSelector selector,
        Action<string> presenter,
        GanttEntityType type)
    {
        ArgumentNullException.ThrowIfNull(inserter);
        ArgumentNullException.ThrowIfNull(nextId);
        ArgumentNullException.ThrowIfNull(selector);
        ArgumentNullException.ThrowIfNull(presenter);

        GanttRowInsertOutcome outcome = inserter.Insert(type, nextId);
        if (outcome.Succeeded)
        {
            SelectInsertedRow(selector, outcome.BodyIndex!.Value);

            // The row is written either way, but a lost padding row is a visible
            // defect the user cannot diagnose from the sheet alone. ADR-0008 rules
            // out passing over it silently.
            if (!outcome.PaddingRowReserved)
            {
                Present(presenter, PaddingRowLostMessage);
            }

            return;
        }

        // CA1031: presenter failure is outside command behaviour; it degrades to no dialog.
        Present(presenter, TranslateRefusal(outcome.Refusal!.Value));
    }

    /// <summary>The message shown when the append consumed the chart's bottom margin.</summary>
    internal const string PaddingRowLostMessage =
        "Gantt Creator added the row, but could not reserve the blank row below the table, "
        + "so the chart's bottom margin has been lost. Refresh the chart, or undo the row and try again.";

    /// <summary>
    /// Shows a message, degrading to silence when the host refuses to present it. A
    /// dialog failure must never turn a completed insert into an escaping exception.
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

    /// <summary>
    /// Moves the selection onto the inserted row, degrading to no selection change
    /// when the host refuses. A selection failure must never turn a successful
    /// insert into a user-visible error or escape into Excel.
    /// </summary>
    /// <param name="selector">The selection port.</param>
    /// <param name="bodyIndex">The inserted row's one-based body index.</param>
    private static void SelectInsertedRow(IInsertedRowSelector selector, int bodyIndex)
    {
        // CA1031: a host selection failure is outside command behaviour.
#pragma warning disable CA1031
        try
        {
            selector.SelectBodyRow(bodyIndex);
        }
        catch
        {
            // Intentionally empty: the row exists; only the cursor stays put.
        }
#pragma warning restore CA1031
    }

    /// <summary>Translates a row-insertion refusal to a user-safe message.</summary>
    /// <param name="refusal">The typed refusal reason.</param>
    /// <returns>The actionable message.</returns>
    internal static string TranslateRefusal(GanttRowInsertRefusalReason refusal) => refusal switch
    {
        GanttRowInsertRefusalReason.NoActiveWorkbook =>
            "Gantt Creator needs an open workbook. Open or create a workbook, then try adding the row again.",
        GanttRowInsertRefusalReason.TableMissing =>
            "Gantt Creator could not find tblGanttData. Run Initialise sheet before adding a row.",
        GanttRowInsertRefusalReason.TargetProtected =>
            "The worksheet or workbook is protected, so no Gantt row was added. Remove protection and try again.",
        GanttRowInsertRefusalReason.TypeOptionsUnavailable =>
            "The row was not added because the Type dropdown could not be prepared. Repair the configuration and try again.",
        _ => "Gantt Creator could not add the row. Try again; if it keeps failing, see the Diagnostics dialog.",
    };
}
