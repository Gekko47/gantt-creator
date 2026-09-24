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
    internal static void RunForExcel(GanttEntityType type)
        => Run(
            new ExcelGanttRowInserter(ExcelDna.Integration.ExcelDnaUtil.Application),
            GanttRowId.New,
            CommandErrorDialog.Show,
            type);

    /// <summary>Runs an injected add-row command.</summary>
    /// <param name="inserter">The row insertion port.</param>
    /// <param name="nextId">The stable row-ID generator.</param>
    /// <param name="presenter">The user-safe refusal presenter.</param>
    /// <param name="type">The canonical catalogue type to insert.</param>
    internal static void Run(
        IGanttRowInserter inserter,
        Func<GanttRowId> nextId,
        Action<string> presenter,
        GanttEntityType type)
    {
        ArgumentNullException.ThrowIfNull(inserter);
        ArgumentNullException.ThrowIfNull(nextId);
        ArgumentNullException.ThrowIfNull(presenter);

        GanttRowInsertOutcome outcome = inserter.Insert(type, nextId);
        if (outcome.Succeeded)
        {
            return;
        }

        // CA1031: presenter failure is outside command behaviour; it degrades to no dialog.
#pragma warning disable CA1031
        try
        {
            presenter(TranslateRefusal(outcome.Refusal!.Value));
        }
        catch
        {
            // Intentionally empty: a dialog failure must not escape into Excel.
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
        _ => "Gantt Creator could not add the row. Try again; if it keeps failing, see the Diagnostics dialog.",
    };
}
