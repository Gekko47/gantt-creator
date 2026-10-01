using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live adapter restoring every managed column of <c>tblGanttData</c> to the
/// code-owned <see cref="GanttColumnAccess"/> classification (R4.7C D1/D2,
/// ADR-0029 D7/D8), as the R4.8A pipeline step D2 requires.
/// </summary>
/// <param name="application">
/// The Excel application object, or <see langword="null"/> when the host supplied
/// none. A foreign object fails the interface cast and degrades to the
/// no-active-workbook refusal with no mutation.
/// </param>
/// <param name="protectionGuard">
/// The protection guard, consulted first per ADR-0008 D4 because this adapter mutates.
/// </param>
/// <remarks>
/// <para>
/// <b>The classification is read from <see cref="GanttTableSchema.Default"/>, the same
/// source <c>ExcelWorkbookInitialiser</c> reads.</b> That is the whole reason this
/// adapter does not accept a column list: a hand-written list here would be a second
/// authority for which columns are engine columns, and the two would disagree the
/// first time R4.7C's schema changed.
/// </para>
/// <para>
/// <b>Only a drifted column is written.</b> A column already carrying the schema's
/// visibility is left alone, so a correct sheet yields an empty write set and a
/// Refresh does not mark the workbook dirty. This mirrors the initialiser's
/// behaviour, including its <c>Range.Hidden</c> unboxing: the PIA declares that
/// property as <see cref="object"/>, so comparing it to a <see cref="bool"/> with
/// <c>!=</c> compares boxed references and is always true — which would rewrite every
/// column on every Refresh and dirty the workbook.
/// </para>
/// <para>
/// <b>COM ownership.</b> The <c>Application</c>, <c>Workbook</c>, <c>Worksheet</c>,
/// <c>ListObject</c> and <c>Range</c> objects reached here are Excel-owned shared
/// roots. This adapter takes no ownership, never calls
/// <c>FinalReleaseComObject</c>, and holds every proxy in a local used without
/// chained member expressions.
/// </para>
/// <para>
/// <b>Not sealed.</b> The four <c>internal virtual</c> accessors are the test seams
/// that let contract tests drive the loop without a worksheet. Sealing the type would
/// make them unoverridable and force the tests to fake COM collections instead.
/// </para>
/// </remarks>
public class ExcelColumnPresentationRestorer(object? application, IWorksheetProtectionGuard? protectionGuard = null)
    : IColumnPresentationPort
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <inheritdoc />
    public ColumnPresentationOutcome EnsureClassification()
    {
        Excel.Application? app = _application;
        Excel.Workbook? workbook = app?.ActiveWorkbook;
        if (workbook is null)
        {
            return ColumnPresentationOutcome.Refused(ColumnPresentationRefusalReason.NoActiveWorkbook);
        }

        ProtectionGuardOutcome activeProtection = _protectionGuard.Query();
        if (activeProtection != ProtectionGuardOutcome.NotProtected)
        {
            return ColumnPresentationOutcome.Refused(
                activeProtection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? ColumnPresentationRefusalReason.NoActiveWorkbook
                    : ColumnPresentationRefusalReason.TargetProtected);
        }

        Excel.Sheets sheets = workbook.Sheets;
        if (!TryFindGanttTable(sheets, out Excel.ListObject? table, out Excel.Worksheet? worksheet)
            || table is null
            || worksheet is null)
        {
            return ColumnPresentationOutcome.Refused(ColumnPresentationRefusalReason.TableMissing);
        }

        ProtectionGuardOutcome targetProtection = _protectionGuard.QueryTarget(worksheet);
        if (targetProtection != ProtectionGuardOutcome.NotProtected)
        {
            return ColumnPresentationOutcome.Refused(
                targetProtection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? ColumnPresentationRefusalReason.NoActiveWorkbook
                    : ColumnPresentationRefusalReason.TargetProtected);
        }

        var restored = 0;
        Excel.ListColumns columns = GetTableColumns(table);
        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        for (var index = 1; index <= schema.Count; index++)
        {
            Excel.ListColumn column = GetColumnAt(columns, index);
            Excel.Range columnRange = column.Range;

            if (columnRange.Locked != schema[index - 1].IsLocked)
            {
                columnRange.Locked = schema[index - 1].IsLocked;
                restored++;
            }

            Excel.Range entireColumn = GetEntireColumn(columnRange);
            var wanted = schema[index - 1].IsHidden;
            if (ReadHiddenFlag(entireColumn) != wanted)
            {
                entireColumn.Hidden = wanted;
                restored++;
            }
        }

        return ColumnPresentationOutcome.Ok(restored);
    }

    /// <summary>
    /// Reads <c>Range.Hidden</c> as a real <see cref="bool"/>.
    /// </summary>
    /// <param name="range">The column range.</param>
    /// <returns>The unwrapped visibility flag.</returns>
    /// <remarks>
    /// The PIA declares <c>Range.Hidden</c> as <see cref="object"/>. Unboxed here for
    /// the same reason the initialiser unboxes it: a boxed comparison against
    /// <see langword="true"/> would never be equal, so every column would be rewritten
    /// on every Refresh and the workbook would be marked dirty every time.
    /// </remarks>
    private static bool ReadHiddenFlag(Excel.Range range) =>
        range.Hidden switch
        {
            bool value => value,
            int value => value != 0,
            double value => value != 0,
            _ => false,
        };

    /// <summary>
    /// Finds the Gantt table and the worksheet holding it. Test seam over the COM indexers.
    /// </summary>
    /// <param name="sheets">The workbook's sheet collection.</param>
    /// <param name="table">The found table, or <see langword="null"/>.</param>
    /// <param name="worksheet">The worksheet holding it, or <see langword="null"/>.</param>
    /// <returns>Whether the table was found.</returns>
    /// <remarks>
    /// Overridable so a contract test can supply a table without a live workbook, on
    /// the same pattern as <c>ExcelRowHeightNormaliser.TryFindTable</c>. The base
    /// implementation walks the real sheet collection.
    /// </remarks>
    internal virtual bool TryFindGanttTable(
        Excel.Sheets sheets,
        out Excel.ListObject? table,
        out Excel.Worksheet? worksheet)
    {
        table = null;
        worksheet = null;
        foreach (Excel.Worksheet candidate in sheets)
        {
            foreach (Excel.ListObject candidateTable in candidate.ListObjects)
            {
                if (string.Equals(candidateTable.Name, GanttTableSchema.TableName, StringComparison.Ordinal))
                {
                    table = candidateTable;
                    worksheet = candidate;
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

    /// <summary>Gets the whole worksheet column behind a table column. Test seam.</summary>
    /// <param name="range">The table column's range.</param>
    /// <returns>The entire column.</returns>
    internal virtual Excel.Range GetEntireColumn(Excel.Range range) => range.EntireColumn;
}