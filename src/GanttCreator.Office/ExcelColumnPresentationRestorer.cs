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
/// <b>Only a drifted column is written, and columns are matched by name.</b> A column
/// already carrying the schema's visibility is left alone, so a correct sheet yields
/// an empty write set and a Refresh does not mark the workbook dirty. This mirrors the
/// initialiser's behaviour, including its unboxing of the two properties this PIA
/// declares as <see cref="object"/>: <c>Range.Locked</c> and <c>Range.Hidden</c>. A
/// direct <c>!=</c> comparison against a <see cref="bool"/> compares boxed references
/// and is always true, which would rewrite every column on every Refresh and dirty the
/// workbook.
/// </para>
/// <para>
/// <b>Matching is by name, never by position,</b> because a positional walk silently
/// misclassifies every column when the live order differs from the schema's. A schema
/// column the live table lacks is a typed
/// <see cref="ColumnPresentationRefusalReason.SchemaColumnMissing"/>; a live column the
/// schema does not declare is left untouched.
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

        // Matched by NAME, not by position. A positional walk assumes the live
        // table's columns are in the schema's exact order, and when a user or an
        // older workbook has them in another order every classification is written
        // to the wrong column — hiding `Description` and leaving `Id` visible, with
        // no refusal and no warning. A name that the schema declares but the live
        // table lacks is a typed refusal rather than a silent skip.
        //
        // RESOLUTION IS A SEPARATE PASS, and it must be. The loop below writes as it
        // classifies, so resolving inside it meant a missing column discovered at
        // position N refused *after* positions 1..N-1 had already been written: the
        // refusal said "nothing was restored" while the worksheet carried a partial
        // classification, and a Refresh could repeat the half-write on every pass.
        // Resolving every required column first makes the refusal genuinely precede
        // any mutation, which is the ADR-0008 D4 contract the rest of the adapter
        // already follows.
        List<Excel.ListColumn> resolved = new(schema.Count);
        for (var index = 0; index < schema.Count; index++)
        {
            if (!TryFindColumnByName(columns, schema[index].Name, out Excel.ListColumn? live) || live is null)
            {
                return ColumnPresentationOutcome.Refused(
                    ColumnPresentationRefusalReason.SchemaColumnMissing);
            }

            resolved.Add(live);
        }

        for (var index = 0; index < schema.Count; index++)
        {
            GanttTableColumn column = schema[index];

            Excel.Range columnRange = resolved[index].Range;

            // The desired value is computed ONCE and used for both the comparison and
            // the assignment. Reading `IsLocked` twice would be correct only while the
            // schema stayed constant across the two reads; one local makes the
            // "already correct means no write" guarantee independent of that.
            var wantedLocked = column.IsLocked;

            // `Range.Locked` is declared `object` in this PIA (verified by reflection
            // over the installed Microsoft.Office.Interop.Excel), so it must be
            // unboxed before comparison for exactly the reason `Hidden` is. Comparing
            // a boxed value with `!=` against a bool compares references and is always
            // true, which would rewrite every column on every Refresh and dirty the
            // workbook.
            //
            // The comparison is EXPLICIT about drift rather than relying on the lifted
            // `bool? != bool` operator. ReadLockedFlag reports an unrecognised host
            // value as null, and null must differ from BOTH desired states — otherwise
            // a mixed-lock authoring column reads as already-correct and the adapter
            // reports success over a range it never repaired.
            var liveLocked = ReadLockedFlag(columnRange);
            if (liveLocked is not { } current || current != wantedLocked)
            {
                columnRange.Locked = wantedLocked;
                restored++;
            }

            Excel.Range entireColumn = GetEntireColumn(columnRange);
            var wantedHidden = column.IsHidden;
            if (ReadHiddenFlag(entireColumn) != wantedHidden)
            {
                entireColumn.Hidden = wantedHidden;
                restored++;
            }
        }

        return ColumnPresentationOutcome.Ok(restored);
    }

    /// <summary>
    /// Finds a live table column by its header name.
    /// </summary>
    /// <param name="columns">The table's list columns.</param>
    /// <param name="name">The schema column name to find.</param>
    /// <param name="column">The found column, or <see langword="null"/>.</param>
    /// <returns><see langword="true"/> when the live table carries the column.</returns>
    /// <remarks>
    /// <para>
    /// <b>Ordinal comparison.</b> A worksheet header is text the schema owns, not a
    /// user-facing label, so a case-insensitive match would let a renamed column
    /// silently receive another column's classification.
    /// </para>
    /// <para>
    /// <b>A live column the schema does not declare is left alone</b> rather than
    /// refused: the loop is driven by the schema, so an extra column is simply never
    /// visited. This adapter restores the add-in's own classification and does not own
    /// a user's addition to the table.
    /// </para>
    /// </remarks>
    private bool TryFindColumnByName(
        Excel.ListColumns columns,
        string name,
        out Excel.ListColumn? column)
    {
        var count = columns.Count;
        for (var index = 1; index <= count; index++)
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

    /// <summary>
    /// Reads <c>Range.Locked</c> as a real <see cref="bool"/>.
    /// </summary>
    /// <param name="range">The column range.</param>
    /// <returns>
    /// The unwrapped lock flag, or <see langword="null"/> when the host reported
    /// something that is not a lock state at all.
    /// </returns>
    /// <remarks>
    /// The same unboxing <see cref="ReadHiddenFlag"/> performs, applied to the other
    /// boxed PIA property this adapter writes.
    /// <para>
    /// <b>An unknown value is <see langword="null"/>, not <see langword="false"/>.</b>
    /// This method used to default to <see langword="false"/>, which is correct for
    /// every LOCKED schema column (false differs from true, so the column is written)
    /// but silently wrong for an AUTHORING column, whose desired state *is* false. A
    /// mixed-lock range — which is exactly what a user produces by locking some cells
    /// of a column — is reported by Excel as null, so it read as false, matched the
    /// desired false, and the adapter reported success while leaving the range mixed.
    /// <see langword="null"/> is not equal to either desired state, so the drift is
    /// corrected by writing, and the redundant assignment a genuinely-false authoring
    /// column used to incur is gone.
    /// </para>
    /// </remarks>
    private static bool? ReadLockedFlag(Excel.Range range) =>
        range.Locked switch
        {
            bool value => value,
            int value => value != 0,
            double value => value != 0,
            _ => null,
        };

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

    /// <summary>Gets a column's header name. Test seam over the COM property.</summary>
    /// <param name="column">The list column.</param>
    /// <returns>The column's header text.</returns>
    internal virtual string GetColumnName(Excel.ListColumn column) => column.Name;

    /// <summary>Gets the whole worksheet column behind a table column. Test seam.</summary>
    /// <param name="range">The table column's range.</param>
    /// <returns>The entire column.</returns>
    internal virtual Excel.Range GetEntireColumn(Excel.Range range) => range.EntireColumn;
}
