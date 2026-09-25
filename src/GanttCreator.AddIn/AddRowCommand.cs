using GanttCreator.Core;
using GanttCreator.Office;

using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.AddIn;

/// <summary>
/// Moves the user's selection onto a body row of the visible Gantt table.
/// Implemented in the AddIn layer because selection is a user-interface concern:
/// Core stays UI-free and the Office row-insertion port stays a data operation.
/// </summary>
internal interface IInsertedRowSelector
{
    /// <summary>Selects the given one-based body row of the visible Gantt table.</summary>
    /// <param name="bodyIndex">The one-based body-row index to select.</param>
    void SelectBodyRow(int bodyIndex);
}

/// <summary>Excel implementation of <see cref="IInsertedRowSelector"/>.</summary>
internal sealed class ExcelInsertedRowSelector(object? application) : IInsertedRowSelector
{
    private readonly Excel.Application? _application = application as Excel.Application;

    /// <inheritdoc />
    public void SelectBodyRow(int bodyIndex)
    {
        Excel.Workbook? workbook = _application?.ActiveWorkbook;
        if (workbook is null)
        {
            return;
        }

        Excel.Sheets sheets = workbook.Sheets;
        var count = sheets.Count;
        for (var sheetIndex = 1; sheetIndex <= count; sheetIndex++)
        {
            object sheet = sheets[sheetIndex];
            if (sheet is not Excel.Worksheet worksheet)
            {
                continue;
            }

            Excel.ListObjects listObjects = worksheet.ListObjects;
            var tableCount = listObjects.Count;
            for (var tableIndex = 1; tableIndex <= tableCount; tableIndex++)
            {
                Excel.ListObject table = listObjects[tableIndex];
                if (!string.Equals(table.Name, GanttTableSchema.TableName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                Excel.Range tableRange = table.Range;
                // The table's first row is the header, so body row N is header+N.
                Excel.Range target = tableRange.Rows[tableRange.Row + bodyIndex];
                target.Select();
                return;
            }
        }
    }
}

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
