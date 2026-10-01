using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>Why managed row heights could not be normalised.</summary>
public enum RowHeightNormalisationRefusalReason
{
    /// <summary>The application object or active workbook was absent.</summary>
    NoActiveWorkbook = 0,

    /// <summary>The Gantt worksheet or its <c>tblGanttData</c> table was missing.</summary>
    TableMissing = 1,

    /// <summary>The target worksheet is protected.</summary>
    TargetProtected = 2,

    /// <summary>
    /// The measured rows could not be turned into a plan: a height was
    /// non-finite or non-positive, or a row kind was unrecognised. Refused rather
    /// than defaulted, because a wrong row height is the exact misalignment this
    /// write exists to remove.
    /// </summary>
    InvalidMeasurement = 3,
}

/// <summary>The typed result of normalising managed row heights.</summary>
/// <param name="RowsWritten">How many rows had a height written.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/> on success.</param>
public sealed record RowHeightNormalisationOutcome(int RowsWritten, RowHeightNormalisationRefusalReason? Refusal)
{
    /// <summary>Gets whether normalisation completed.</summary>
    public bool Succeeded => Refusal is null;

    /// <summary>Creates a successful outcome.</summary>
    /// <param name="rowsWritten">How many rows were written.</param>
    /// <returns>The successful outcome.</returns>
    public static RowHeightNormalisationOutcome Ok(int rowsWritten) => new(rowsWritten, null);

    /// <summary>Creates a refusal.</summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <returns>The refusal outcome.</returns>
    public static RowHeightNormalisationOutcome Refused(RowHeightNormalisationRefusalReason refusal) =>
        new(0, refusal);
}

/// <summary>
/// Live adapter that restores every managed row of <c>tblGanttData</c> to
/// <c>GanttRowHeightPt</c>, so the worksheet and the chart cannot disagree about
/// how tall a row is (R4.7D, ADR-0026 D3/D4).
/// </summary>
/// <param name="application">
/// The Excel application object, or <see langword="null"/> when the host supplied
/// none. A foreign object fails the interface cast and degrades to the
/// no-active-workbook refusal with no mutation.
/// </param>
/// <param name="protectionGuard">
/// The protection guard, consulted first per ADR-0008 D4 because this adapter
/// mutates.
/// </param>
/// <remarks>
/// <para>
/// <b>Why this is the add-in's job.</b> The lane height is the measured row height,
/// and the measurement is only meaningful once the sheet is uniform. A user who
/// drags one managed row taller would otherwise change one lane and desynchronise
/// the chart from the rows being read, so the drag is restored on Initialise, on
/// hierarchy mutation and on Refresh.
/// </para>
/// <para>
/// <b>Why rows already at the token are skipped.</b> The plan excludes them, so a
/// correctly normalised sheet produces an empty write set. Writing them anyway
/// would mark the workbook dirty on every Refresh.
/// </para>
/// <para>
/// <b>COM ownership.</b> The <c>Application</c>, <c>Workbook</c>, <c>Worksheet</c>,
/// <c>ListObject</c> and <c>Range</c> objects reached here are Excel-owned shared
/// roots. This adapter takes no ownership, never calls
/// <c>FinalReleaseComObject</c>, and holds every proxy in a local used without
/// chained member expressions.
/// </para>
/// <para>
/// The <c>internal virtual</c> seams isolate the COM parameterised properties so
/// contract tests can substitute them; the real <c>RowHeight</c> round-trip is
/// exercised by the tagged live-Office integration test.
/// </para>
/// </remarks>
public class ExcelRowHeightNormaliser(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null) : IRowHeightNormalisationPort
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IWorksheetProtectionGuard _protectionGuard = protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <inheritdoc />
    public RowHeightNormalisationOutcome Normalise(
        double managedHeightPt,
        double splitterHeightPt,
        double spacerHeightPt,
        double headerHeightPt,
        double reservedRowHeightPt,
        double paddingRowHeightPt)
    {
        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (workbook is null)
        {
            return RowHeightNormalisationOutcome.Refused(RowHeightNormalisationRefusalReason.NoActiveWorkbook);
        }

        if (!TryFindTable(workbook.Sheets, out Excel.Worksheet? worksheet, out Excel.ListObject? table)
            || worksheet is null
            || table is null)
        {
            return RowHeightNormalisationOutcome.Refused(RowHeightNormalisationRefusalReason.TableMissing);
        }

        // First check in a mutating adapter (ADR-0008 D4), before any COM write.
        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(worksheet);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return RowHeightNormalisationOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? RowHeightNormalisationRefusalReason.NoActiveWorkbook
                    : RowHeightNormalisationRefusalReason.TargetProtected);
        }

        // The layout rows are normalised FIRST, and before the body's zero-row early
        // return, because they exist independently of whether the table has any data
        // yet: the header row is the period band's row (D5), the reserved row carries
        // the title and year band (D4), and the top padding row is the chart's top
        // margin (ADR-0031 D2). A freshly initialised table has a header and no body,
        // and returning early on the body would leave them all at Excel's default and
        // the first lane misaligned again.
        var written = 0;
        written += NormaliseLayoutRow(
            worksheet,
            GanttSheetLayout.TopPaddingRowIndex,
            paddingRowHeightPt);

        written += NormaliseLayoutRow(
            worksheet,
            GanttSheetLayout.ReservedRowIndex,
            reservedRowHeightPt);

        written += NormaliseLayoutRow(worksheet, GanttSheetLayout.HeaderRowIndex, headerHeightPt);

        Excel.Range? body = GetTableBody(table);
        if (body is null)
        {
            return RowHeightNormalisationOutcome.Ok(written);
        }

        var rowCount = GetBodyRowCount(body);
        if (rowCount == 0)
        {
            return RowHeightNormalisationOutcome.Ok(written);
        }

        // The bottom padding row is derived from the body length, so it can only be
        // addressed once the row count is known - which is why it is normalised here
        // and not alongside the rows above. It is normalised BEFORE the body rows
        // themselves for the same reason they are: a user who drags it is asking for
        // a different margin, and restoring it is the whole point of this adapter.
        written += NormaliseLayoutRow(
            worksheet,
            GanttSheetLayout.BottomPaddingRowIndex(rowCount),
            paddingRowHeightPt);

        // Each row's KIND is read from its own Type cell, because the height policy
        // differs by kind: a Splitter follows `SplitterPt` and a Spacer `SpacerPt`,
        // never the managed height. Passing every row as MeasuredRowKind.Managed --
        // as this did -- meant a splitter or spacer row was dragged to the managed
        // height on the very Refresh that was supposed to restore the sheet, so the
        // section header and the blank separator were permanently the wrong height and
        // no lane could line up with its row.
        //
        // The Type column is resolved ONCE, before the row loop. KindOfRow used to
        // re-run TryGetTypeColumn for every row, so a body of N rows walked the
        // ListColumns collection N times over COM before any cell was even read, on
        // the exact path that already walks N rows of heights.
        //
        // The boolean is deliberately discarded: a false here means the Type column
        // could not be located, and `typeColumn` is then null, which KindOfRow
        // already treats as "every row is Managed". Branching here would need a
        // second outcome for a condition that has no different handling.
        _ = TryGetTypeColumn(table, out Excel.Range? typeColumn);
        List<MeasuredRowHeight> measured = [];
        for (var index = 1; index <= rowCount; index++)
        {
            if (GetBodyRowAt(body, index) is not { } row || ToPoints(row.RowHeight) is not { } height)
            {
                return RowHeightNormalisationOutcome.Refused(
                    RowHeightNormalisationRefusalReason.InvalidMeasurement);
            }

            measured.Add(new MeasuredRowHeight(index, height, KindOfRow(typeColumn, index)));
        }

        RowHeightNormalisationPlan? plan = RowHeightNormaliser.Plan(
            measured,
            managedHeightPt,
            splitterHeightPt,
            spacerHeightPt);
        if (plan is null)
        {
            return RowHeightNormalisationOutcome.Refused(
                RowHeightNormalisationRefusalReason.InvalidMeasurement);
        }

        var writtenAfterBody = 0;
        foreach (RowHeightNormalisation decision in plan.Rows)
        {
            if (decision.TargetHeightPt is not { } target || GetBodyRowAt(body, decision.RowNumber) is not { } targetRow)
            {
                continue;
            }

            targetRow.RowHeight = target;
            writtenAfterBody++;
        }

        return RowHeightNormalisationOutcome.Ok(written + writtenAfterBody);
    }

    /// <summary>
    /// Normalises one layout row — the reserved title row or the header row — to its
    /// token, writing only when the row is not already there.
    /// </summary>
    /// <param name="worksheet">The Gantt worksheet.</param>
    /// <param name="rowIndex">The one-based worksheet row.</param>
    /// <param name="targetHeightPt">The height the row should carry.</param>
    /// <returns>1 when the row was written, 0 when it was already correct or unreadable.</returns>
    /// <remarks>
    /// The same <c>RowHeightNormaliser.EqualityTolerancePt</c> comparison the body
    /// plan uses is applied, so a correctly normalised sheet writes nothing and the
    /// Refresh does not mark the workbook dirty. An unreadable row is skipped rather
    /// than refused: the body already refuses the whole normalisation when it cannot
    /// measure a row, and a layout row the host will not report is not a reason to
    /// abandon the managed rows that are.
    /// <para>
    /// The row is reached through <c>Worksheet.Rows[...]</c> rather than
    /// <c>Worksheet.Range[...]</c> for the same reason the outline writer changed
    /// accessor: the host's treatment of a range depends on which collection produced
    /// it, and <c>Rows</c> is the whole-row form.
    /// </para>
    /// </remarks>
    private int NormaliseLayoutRow(Excel.Worksheet worksheet, int rowIndex, double targetHeightPt)
    {
        if (!double.IsFinite(targetHeightPt) || targetHeightPt <= 0)
        {
            return 0;
        }

        Excel.Range? row = GetLayoutRow(worksheet, rowIndex);
        if (row is null || ToPoints(row.RowHeight) is not { } current)
        {
            return 0;
        }

        if (Math.Abs(current - targetHeightPt) <= RowHeightNormaliser.EqualityTolerancePt)
        {
            return 0;
        }

        row.RowHeight = targetHeightPt;
        return 1;
    }

    /// <summary>
    /// Returns the whole worksheet row at a one-based index. Test seam over the COM
    /// parameterised <c>Worksheet.Rows</c> property.
    /// </summary>
    /// <param name="worksheet">The Gantt worksheet.</param>
    /// <param name="rowIndex">The one-based worksheet row.</param>
    /// <returns>The row's range, or <see langword="null"/> when the host resolves none.</returns>
    internal virtual Excel.Range? GetLayoutRow(Excel.Worksheet worksheet, int rowIndex)
    {
        ArgumentNullException.ThrowIfNull(worksheet);
        return worksheet.Rows[rowIndex];
    }

    /// <summary>
    /// Maps one body row's <c>Type</c> cell to the height policy its row follows.
    /// </summary>
    /// <param name="typeColumn">
    /// The already-resolved <c>Type</c> column's body range, or
    /// <see langword="null"/> when the column could not be located. Resolving it
    /// once per normalisation, rather than once per row, is what keeps the column
    /// lookup off the per-row path.
    /// </param>
    /// <param name="index">The one-based body-row index.</param>
    /// <returns>The row kind; <see cref="MeasuredRowKind.Managed"/> for anything else.</returns>
    /// <remarks>
    /// The Type column is located through <see cref="GanttTableSchema"/> rather than a
    /// literal index, so a column reorder cannot silently start measuring the wrong
    /// cell. An unreadable, blank or unrecognised Type is <b>Managed</b>: that is the
    /// policy every ordinary activity, milestone and child row follows, and guessing a
    /// structural kind for a row whose Type could not be read would write a height the
    /// row's own type never asked for.
    /// </remarks>
    private MeasuredRowKind KindOfRow(Excel.Range? typeColumn, int index) =>
        typeColumn is not null
        && ReadTypeCellText(typeColumn, index) is { } text
        && EntityTypeCatalog.TryParse(text, out GanttEntityType parsed)
            ? parsed switch
            {
                // Only the two structural rows follow their own token. Every other type
                // -- activity, milestone, interval, procurement -- is an ordinary
                // managed row, and listing them explicitly (rather than a discard arm)
                // makes a future structural type fail here instead of being silently
                // normalised to the managed height.
                GanttEntityType.Splitter => MeasuredRowKind.Splitter,
                GanttEntityType.Spacer => MeasuredRowKind.Spacer,
                GanttEntityType.AsBuiltActivity
                or GanttEntityType.AsPlannedActivity
                or GanttEntityType.BaselineActivity
                or GanttEntityType.DelayEvent
                or GanttEntityType.AsBuiltProcurement
                or GanttEntityType.AsPlannedProcurement
                or GanttEntityType.BaselineProcurement
                or GanttEntityType.CustomActivity
                or GanttEntityType.AsBuiltMilestone
                or GanttEntityType.AsPlannedMilestone
                or GanttEntityType.BaselineMilestone
                or GanttEntityType.CriticalMilestone
                or GanttEntityType.CriticalInterval
                or GanttEntityType.Delineator => MeasuredRowKind.Managed,
                _ => MeasuredRowKind.Managed,
            }
            : MeasuredRowKind.Managed;

    /// <summary>
    /// Locates the <c>Type</c> column's body range, by name from the schema.
    /// Test seam over the COM list-column indexer.
    /// </summary>
    /// <param name="source">The Gantt data table.</param>
    /// <param name="typeColumn">The resolved column's body range.</param>
    /// <returns>Whether the column was found and has a body.</returns>
    internal virtual bool TryGetTypeColumn(Excel.ListObject source, out Excel.Range? typeColumn)
    {
        ArgumentNullException.ThrowIfNull(source);

        typeColumn = null;
        Excel.ListColumns? columns = source.ListColumns;
        if (columns is null
            || !GanttTableSchema.Default.TryGetColumn("Type", out GanttTableColumn? typeDef)
            || typeDef is null)
        {
            return false;
        }

        for (var index = 1; index <= columns.Count; index++)
        {
            Excel.ListColumn candidate = ColumnAt(columns, index);
            if (string.Equals(candidate.Name, typeDef.Name, StringComparison.Ordinal))
            {
                typeColumn = candidate.DataBodyRange;
                return typeColumn is not null;
            }
        }

        return false;
    }

    /// <summary>
    /// Reads one body row's <c>Type</c> text. Test seam over the COM cell read.
    /// </summary>
    /// <param name="typeColumn">The <c>Type</c> column's body range.</param>
    /// <param name="index">The one-based body-row index.</param>
    /// <returns>The cell's text, or <see langword="null"/> when it is not readable text.</returns>
    internal virtual string? ReadTypeCellText(Excel.Range typeColumn, int index)
    {
        ArgumentNullException.ThrowIfNull(typeColumn);

        return CellValueAt(typeColumn, index) as string;
    }

    /// <summary>Reads one cell of a column by its one-based row index. Test seam.</summary>
    /// <param name="column">The column range.</param>
    /// <param name="index">The one-based row index.</param>
    /// <returns>The cell's raw value.</returns>
    internal virtual object? CellValueAt(Excel.Range column, int index)
    {
        Excel.Range? rows = column.Rows;
        return rows?[index].Value2;
    }

    /// <summary>Returns the list column at the one-based index. Test seam.</summary>
    /// <param name="columns">The table's list columns.</param>
    /// <param name="index">The one-based column index.</param>
    /// <returns>The list column.</returns>
    internal virtual Excel.ListColumn ColumnAt(Excel.ListColumns columns, int index) => columns[index];

    /// <summary>Finds the Gantt worksheet and its table. Test seam over the COM collection indexers.</summary>
    /// <param name="sheets">The workbook's sheet collection.</param>
    /// <param name="worksheet">The resolved worksheet.</param>
    /// <param name="table">The resolved table.</param>
    /// <returns>Whether both were found.</returns>
    /// <remarks>
    /// Iterating <c>Sheets</c> with a <see cref="Excel.Worksheet"/> loop variable would
    /// throw on a workbook carrying a chart sheet: this PIA exposes no <c>Sheet</c>
    /// type, so the collection enumerates as <see cref="object"/> and a chart sheet
    /// cannot be cast. The cast is therefore done per entry and a non-worksheet is
    /// skipped, which is what <c>ExcelGanttTableReader</c> already does. The name
    /// comparison reads <see cref="GanttTableSchema.TableName"/> and is
    /// case-insensitive, because Excel preserves whatever case a table was created
    /// with.
    /// </remarks>
    internal virtual bool TryFindTable(
        Excel.Sheets sheets,
        out Excel.Worksheet? worksheet,
        out Excel.ListObject? table)
    {
        worksheet = null;
        table = null;
        foreach (var entry in sheets)
        {
            if (entry is not Excel.Worksheet candidate)
            {
                continue;
            }

            foreach (Excel.ListObject candidateTable in candidate.ListObjects)
            {
                if (string.Equals(
                    candidateTable.Name,
                    GanttTableSchema.TableName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    worksheet = candidate;
                    table = candidateTable;
                    return true;
                }
            }
        }

        return false;
    }

    /// <summary>Gets a table's data-body range. Test seam over the COM parameterised property.</summary>
    /// <param name="table">The table.</param>
    /// <returns>The body range, or <see langword="null"/> when there is no body.</returns>
    internal virtual Excel.Range? GetTableBody(Excel.ListObject table)
    {
        ArgumentNullException.ThrowIfNull(table);

        return table.DataBodyRange;
    }

    /// <summary>Reads the body row count. Test seam over <c>Range.Rows.Count</c>.</summary>
    /// <param name="body">The body range.</param>
    /// <returns>The row count.</returns>
    internal virtual int GetBodyRowCount(Excel.Range body)
    {
        ArgumentNullException.ThrowIfNull(body);

        Excel.Range? rows = body.Rows;
        return rows?.Count ?? 0;
    }

    /// <summary>Reads one body row by its 1-based index. Test seam over <c>Range.Rows.Item</c>.</summary>
    /// <param name="body">The body range.</param>
    /// <param name="index">The 1-based row index.</param>
    /// <returns>The row range, or <see langword="null"/> when the host does not resolve it.</returns>
    internal virtual Excel.Range? GetBodyRowAt(Excel.Range body, int index)
    {
        ArgumentNullException.ThrowIfNull(body);

        return body.Rows is { } rows ? rows[index] : null;
    }

    /// <summary>
    /// Converts a host-reported height to points, treating an absent measurement as
    /// absent rather than as zero.
    /// </summary>
    /// <param name="raw">The boxed host value.</param>
    /// <returns>The points, or <see langword="null"/> when the host reported none.</returns>
    /// <remarks>
    /// <c>DBNull.Value</c> is refused too, not just <see langword="null"/>: an
    /// <c>Object</c>-typed Excel member returns it for a range with no single value,
    /// and converting it would throw an <see cref="InvalidCastException"/> rather
    /// than the typed refusal this adapter reports.
    /// </remarks>
    private static double? ToPoints(object? raw) =>
        raw is null or DBNull ? null : Convert.ToDouble(raw, System.Globalization.CultureInfo.InvariantCulture);
}
