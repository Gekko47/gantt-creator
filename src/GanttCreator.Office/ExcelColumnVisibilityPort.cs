using System.Runtime.InteropServices;
using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live adapter for the ribbon's column-visibility dropdown (R5.18h): reads and applies
/// the hidden state of the columns <see cref="GanttColumnVisibility.ToggleableColumnNames"/>
/// offers, over <c>tblGanttData</c>.
/// </summary>
/// <param name="application">
/// The Excel application object, or <see langword="null"/> when the host supplied none. A
/// foreign object fails the cast and degrades to the no-active-workbook refusal with no
/// mutation.
/// </param>
/// <param name="protectionGuard">
/// The protection guard, consulted before <see cref="Apply"/> writes, per ADR-0008 D4.
/// </param>
/// <remarks>
/// <para>
/// <b>Columns are matched by name, never by position.</b> A positional walk assumes the
/// live table is in the schema's order; when it is not, every flag lands on the wrong
/// column and the user's hidden choice hides something else. The names come from the Core
/// contract, so the dropdown and this writer cannot disagree about which columns exist.
/// </para>
/// <para>
/// <b>Resolution precedes mutation.</b> All five columns are resolved before any flag is
/// written, so a table missing one refuses with nothing changed rather than leaving a
/// half-applied choice behind — the ordering the presentation restorer adopted after its
/// own partial-write defect.
/// </para>
/// <para>
/// <b>Only drift is written.</b> A column already carrying the requested state is left
/// alone, so opening the workbook and picking the same item twice does not dirty it.
/// </para>
/// <para>
/// <b>COM ownership.</b> The <c>Application</c>, <c>Workbook</c>, <c>Worksheet</c>,
/// <c>ListObject</c>, <c>ListColumn</c> and <c>Range</c> objects reached here are
/// Excel-owned shared roots. This adapter takes no ownership, never calls
/// <c>FinalReleaseComObject</c>, and holds every proxy in a local used without chained
/// member expressions.
/// </para>
/// <para>
/// <b>Not sealed.</b> The <c>internal virtual</c> accessors are the test seams that let
/// contract tests drive both methods without a worksheet.
/// </para>
/// </remarks>
public class ExcelColumnVisibilityPort(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null) : IColumnVisibilityPort
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <inheritdoc />
    public ColumnVisibilityPortOutcome ReadHidden()
    {
        if (!TryActiveSheets(out Excel.Sheets? sheets))
        {
            return ColumnVisibilityPortOutcome.Refused(ColumnVisibilityRefusalReason.NoActiveWorkbook);
        }

        if (!TryFindGanttTable(sheets!, out _, out Excel.ListObject? table) || table is null)
        {
            return ColumnVisibilityPortOutcome.Refused(ColumnVisibilityRefusalReason.TableMissing);
        }

        if (!TryResolveColumns(table, out List<Excel.ListColumn>? columns))
        {
            return ColumnVisibilityPortOutcome.Refused(ColumnVisibilityRefusalReason.UnknownColumn);
        }

        HashSet<string> hidden = new(StringComparer.Ordinal);
        for (int index = 0; index < columns!.Count; index++)
        {
            Excel.Range columnRange = columns[index].Range;
            Excel.Range entireColumn = GetEntireColumn(columnRange);
            if (ReadHiddenFlag(entireColumn))
            {
                _ = hidden.Add(GanttColumnVisibility.ToggleableColumnNames[index]);
            }
        }

        return ColumnVisibilityPortOutcome.Ok(hidden);
    }

    /// <inheritdoc />
    public ColumnVisibilityPortOutcome Apply(IReadOnlySet<string> hidden)
    {
        ArgumentNullException.ThrowIfNull(hidden);

        // Refused before the workbook is touched: a name outside the Core contract is a
        // caller defect, and writing the rest of the set would hide a column and report a
        // failure that left the sheet half-changed.
        foreach (string name in hidden)
        {
            if (!GanttColumnVisibility.IsToggleable(name))
            {
                return ColumnVisibilityPortOutcome.Refused(ColumnVisibilityRefusalReason.UnknownColumn);
            }
        }

        if (!TryActiveSheets(out Excel.Sheets? sheets))
        {
            return ColumnVisibilityPortOutcome.Refused(ColumnVisibilityRefusalReason.NoActiveWorkbook);
        }

        if (!TryFindGanttTable(sheets!, out Excel.Worksheet? worksheet, out Excel.ListObject? table)
            || worksheet is null
            || table is null)
        {
            return ColumnVisibilityPortOutcome.Refused(ColumnVisibilityRefusalReason.TableMissing);
        }

        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(worksheet!);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return ColumnVisibilityPortOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? ColumnVisibilityRefusalReason.NoActiveWorkbook
                    : ColumnVisibilityRefusalReason.TargetProtected);
        }

        if (!TryResolveColumns(table, out List<Excel.ListColumn>? columns))
        {
            return ColumnVisibilityPortOutcome.Refused(ColumnVisibilityRefusalReason.UnknownColumn);
        }

        int written = 0;
        for (int index = 0; index < columns!.Count; index++)
        {
            bool wanted = hidden.Contains(GanttColumnVisibility.ToggleableColumnNames[index]);
            Excel.Range columnRange = columns[index].Range;
            Excel.Range entireColumn = GetEntireColumn(columnRange);
            if (ReadHiddenFlag(entireColumn) == wanted)
            {
                continue;
            }

            try
            {
                SetHiddenFlag(entireColumn, wanted);
            }
            catch (COMException)
            {
                return ColumnVisibilityPortOutcome.Refused(ColumnVisibilityRefusalReason.WriteFailed);
            }

            written++;
        }

        return ColumnVisibilityPortOutcome.Ok(hidden, written);
    }

    /// <summary>
    /// Finds the active workbook's Gantt worksheet and table.
    /// </summary>
    /// <param name="sheets">The active workbook's sheet collection, or <see langword="null"/>.</param>
    /// <returns><see langword="true"/> when there is an active workbook to search.</returns>
    private bool TryActiveSheets(out Excel.Sheets? sheets)
    {
        sheets = null;

        Excel.Application? app = _application;
        Excel.Workbook? workbook = app?.ActiveWorkbook;
        if (workbook is null)
        {
            return false;
        }

        sheets = workbook.Sheets;
        return true;
    }

    /// <summary>
    /// Resolves every offered column by name, in contract order.
    /// </summary>
    /// <param name="table">The live data table.</param>
    /// <param name="columns">The resolved columns, or <see langword="null"/>.</param>
    /// <returns><see langword="true"/> when every offered column was found.</returns>
    private bool TryResolveColumns(Excel.ListObject table, out List<Excel.ListColumn>? columns)
    {
        columns = null;

        Excel.ListColumns live = GetTableColumns(table);
        List<Excel.ListColumn> resolved = new(GanttColumnVisibility.ToggleableColumnNames.Count);
        foreach (string name in GanttColumnVisibility.ToggleableColumnNames)
        {
            if (!TryFindColumnByName(live, name, out Excel.ListColumn? column) || column is null)
            {
                return false;
            }

            resolved.Add(column);
        }

        columns = resolved;
        return true;
    }

    /// <summary>Finds a live table column by its header name.</summary>
    /// <param name="columns">The table's list columns.</param>
    /// <param name="name">The exact header name to find.</param>
    /// <param name="column">The found column, or <see langword="null"/>.</param>
    /// <returns><see langword="true"/> when the live table carries the name.</returns>
    private bool TryFindColumnByName(
        Excel.ListColumns columns,
        string name,
        out Excel.ListColumn? column)
    {
        int count = columns.Count;
        for (int index = 1; index <= count; index++)
        {
            Excel.ListColumn candidate = GetColumnAt(columns, index);
            if (string.Equals(GetColumnName(candidate), name, StringComparison.Ordinal))
            {
                column = candidate;
                return true;
            }
        }

        column = null;
        return false;
    }

    /// <summary>Finds the Gantt table in a sheet collection. Test seam.</summary>
    /// <param name="sheets">The workbook's sheet collection.</param>
    /// <param name="worksheet">The worksheet holding the table, or <see langword="null"/>.</param>
    /// <param name="table">The found table, or <see langword="null"/>.</param>
    /// <returns>Whether the table was found.</returns>
    internal virtual bool TryFindGanttTable(
        Excel.Sheets sheets,
        out Excel.Worksheet? worksheet,
        out Excel.ListObject? table)
    {
        worksheet = null;
        table = null;
        foreach (Excel.Worksheet candidate in sheets)
        {
            foreach (Excel.ListObject candidateTable in candidate.ListObjects)
            {
                if (GanttChartContract.IsDataTableNameExact(candidateTable.Name))
                {
                    worksheet = candidate;
                    table = candidateTable;
                    return true;
                }
            }
        }

        return false;
    }

    /// <summary>Gets a table's list columns. Test seam.</summary>
    /// <param name="table">The table.</param>
    /// <returns>The list columns.</returns>
    internal virtual Excel.ListColumns GetTableColumns(Excel.ListObject table) => table.ListColumns;

    /// <summary>Returns the list column at the one-based index. Test seam.</summary>
    /// <param name="columns">The table's list columns.</param>
    /// <param name="index">The one-based column index.</param>
    /// <returns>The list column.</returns>
    internal virtual Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index) => columns[index];

    /// <summary>Gets a column's header name. Test seam over the COM property.</summary>
    /// <param name="column">The list column.</param>
    /// <returns>The column's header text.</returns>
    internal virtual string GetColumnName(Excel.ListColumn column) => column.Name;

    /// <summary>Gets the whole worksheet column behind a table column. Test seam.</summary>
    /// <param name="range">The table column's range.</param>
    /// <returns>The entire column.</returns>
    internal virtual Excel.Range GetEntireColumn(Excel.Range range) => range.EntireColumn;

    /// <summary>Reads <c>Range.Hidden</c> as a real <see cref="bool"/>. Test seam.</summary>
    /// <param name="range">The whole worksheet column.</param>
    /// <returns>The unwrapped visibility flag.</returns>
    /// <remarks>
    /// The PIA declares <c>Range.Hidden</c> as <see cref="object"/>, so comparing it to a
    /// <see cref="bool"/> with <c>!=</c> would compare boxed references, read as always
    /// different, and rewrite every offered column on every call. An unrecognised value
    /// reads as "not hidden": the visible default causes the write that makes the column
    /// match the requested state.
    /// </remarks>
    internal virtual bool ReadHiddenFlag(Excel.Range range) => range.Hidden switch
    {
        bool value => value,
        int value => value != 0,
        double value => value != 0,
        _ => false,
    };

    /// <summary>Writes <c>Range.Hidden</c>. Test seam.</summary>
    /// <param name="range">The whole worksheet column.</param>
    /// <param name="hidden">The state to write.</param>
    internal virtual void SetHiddenFlag(Excel.Range range, bool hidden)
    {
        ArgumentNullException.ThrowIfNull(range);
        range.Hidden = hidden;
    }
}
