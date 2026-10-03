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

    /// <summary>
    /// The reserved bottom padding row could not be verified: the measured body
    /// span and the row below it do not form the reserved margin, or that row
    /// carries content (ADR-0035 D2).
    /// </summary>
    PaddingRowNotOwned = 4,

    /// <summary>
    /// The reserved anchor row could not be verified: the measured body span and the
    /// row below it do not form the reserved anchor, or that row carries content
    /// (ADR-0038 D6).
    /// </summary>
    /// <remarks>
    /// Its own reason rather than a member of the padding row's, because the two rows
    /// are separate reservations with different remedies and a diagnostic that cannot
    /// say which row is wrong is one the user cannot act on.
    /// </remarks>
    AnchorRowNotOwned = 5,
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
        double paddingRowHeightPt,
        double anchorRowHeightPt)
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

        // The layout rows exist independently of whether the table has any data yet:
        // the header row is the period band's row (D5), the reserved row carries the
        // title and year band (D4), and the top padding row is the chart's top margin
        // (ADR-0031 D2). A freshly initialised table has a header and no body, and
        // returning early on the body would leave them all at Excel's default and the
        // first lane misaligned again. So they are normalised on every path.
        //
        // They are written AFTER the padding-row resolution, not before it. This used
        // to normalise them first and comment that the refusal check came "before the
        // body rows are written, so a workbook in this state is not half-normalised" --
        // which was true of the body rows and false of the layout rows: a
        // PaddingRowNotOwned refusal already returned with three rows resized. The
        // refusal is the adapter's statement that the sheet is not in the layout the
        // add-in believes in, and a statement like that must be made before mutating
        // anything, not after. Ordering the resolution first costs nothing: the
        // layout rows still land on the same paths, and the zero-body path still
        // normalises them because it never needs a padding row to resolve.
        Excel.Range? body = GetTableBody(table);
        var rowCount = body is null ? 0 : GetBodyRowCount(body);

        int? bottomPaddingRow = null;
        int? anchorRow = null;
        if (body is not null && rowCount > 0)
        {
            // ADR-0035 D2: the bottom padding row is now RESOLVED AND VERIFIED against
            // the body's measured worksheet span rather than derived from the row count
            // on trust. The old arithmetic (`FirstBodyRowIndex + rowCount`) was correct
            // only while the table sat exactly where the layout authority assumed; a
            // table that had moved made this adapter write the chart's margin height to
            // an arbitrary user row. Refusing is the honest outcome -- the sheet does
            // not have the layout the add-in believes in, and Initialise/Repair is the
            // remedy.
            BottomPaddingResolution padding = ResolveBottomPaddingRow(worksheet, table, body, rowCount);
            if (!padding.Succeeded)
            {
                return RowHeightNormalisationOutcome.Refused(
                    RowHeightNormalisationRefusalReason.PaddingRowNotOwned);
            }

            bottomPaddingRow = padding.Row;

            // ADR-0038 D6: the anchor row is resolved and verified the same way, and
            // BEFORE anything is written. Ordering matters here for the reason the
            // comment above the layout rows gives: a refusal is this adapter's
            // statement that the sheet is not in the layout the add-in believes in,
            // and such a statement must be made before mutating, not after. Resolving
            // it here rather than alongside the padding row is what keeps the two
            // independently verifiable -- a caller that verified one and assumed the
            // other is exactly the substitution that let the padding row be consumed.
            AnchorRowResolution anchor = ResolveAnchorRow(worksheet, table, body, rowCount);
            if (!anchor.Succeeded)
            {
                return RowHeightNormalisationOutcome.Refused(
                    RowHeightNormalisationRefusalReason.AnchorRowNotOwned);
            }

            anchorRow = anchor.Row;
        }

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

        // Normalised after the body rows are MEASURED for the same reason the others
        // exist at all: a user who drags it is asking for a different margin, and
        // restoring it is the whole point of this adapter. Skipped when there is no
        // body, because then no padding row was resolved and none may be written.
        if (bottomPaddingRow is { } resolvedPaddingRow)
        {
            written += NormaliseLayoutRow(worksheet, resolvedPaddingRow, paddingRowHeightPt);
        }

        // The anchor row takes its OWN sub-row token (ADR-0038 D1), never the padding
        // token. Writing the padding height here would put the plot-spanning shapes'
        // bottom cell anchor a whole margin-height below the body, which is a visible
        // strip of blank sheet rather than the sub-row anchor the closing line's
        // arithmetic depends on.
        if (anchorRow is { } resolvedAnchorRow)
        {
            written += NormaliseLayoutRow(worksheet, resolvedAnchorRow, anchorRowHeightPt);
        }

        if (body is null || rowCount == 0)
        {
            return RowHeightNormalisationOutcome.Ok(written);
        }

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
    /// Resolves and verifies the reserved bottom padding row from the body's
    /// MEASURED worksheet span (ADR-0035 D2).
    /// </summary>
    /// <param name="worksheet">The Gantt worksheet.</param>
    /// <param name="table">The Gantt table, used only to reach its worksheet.</param>
    /// <param name="body">The table's data-body range.</param>
    /// <param name="rowCount">The measured body row count.</param>
    /// <returns>The verified padding row, or a typed refusal.</returns>
    /// <remarks>
    /// <para>
    /// The measured span is what makes the check capable of failing. Deriving both
    /// sides from the row count would compare the layout authority with itself and
    /// always agree, which is the shape of bug this replaces.
    /// </para>
    /// <para>
    /// Emptiness is verified through a seam rather than assumed, because "the user
    /// has not typed in it" is the one fact the previous code never established and
    /// the one that produced the reported defect. A row that cannot be read yields
    /// <see langword="null"/> -- not established -- which refuses.
    /// </para>
    /// </remarks>
    private BottomPaddingResolution ResolveBottomPaddingRow(
        Excel.Worksheet worksheet,
        Excel.ListObject table,
        Excel.Range body,
        int rowCount)
    {
        ArgumentNullException.ThrowIfNull(worksheet);
        ArgumentNullException.ThrowIfNull(table);
        ArgumentNullException.ThrowIfNull(body);

        var firstBodyRow = GetRangeRow(body);
        var lastBodyRow = rowCount > 0 && firstBodyRow > 0
            ? firstBodyRow + rowCount - 1
            : 0;
        // One past the reserved ANCHOR row (ADR-0038 D1). This was `lastBodyRow + 1`
        // until the anchor row was inserted between the body and the padding row;
        // left alone it would resolve the anchor row as the padding row and write the
        // chart's margin height into a 0.25pt anchor.
        var candidate = lastBodyRow > 0 ? lastBodyRow + 2 : (int?)null;

        return GanttSheetLayout.ResolveBottomPaddingRow(
            firstBodyRow,
            lastBodyRow,
            candidate,
            candidate is { } row ? IsLayoutRowEmpty(worksheet, row) : null);
    }

    /// <summary>
    /// Resolves and verifies the reserved anchor row from the body's MEASURED
    /// worksheet span (ADR-0038 D6).
    /// </summary>
    /// <param name="worksheet">The Gantt worksheet.</param>
    /// <param name="table">The Gantt table, used only to reach its worksheet.</param>
    /// <param name="body">The table's data-body range.</param>
    /// <param name="rowCount">The measured body row count.</param>
    /// <returns>The verified anchor row, or a typed refusal.</returns>
    /// <remarks>
    /// The measured span is what makes the check capable of failing, for the same
    /// reason it is on the padding row: deriving both sides from the row count would
    /// compare the layout authority with itself and always agree.
    /// </remarks>
    private AnchorRowResolution ResolveAnchorRow(
        Excel.Worksheet worksheet,
        Excel.ListObject table,
        Excel.Range body,
        int rowCount)
    {
        ArgumentNullException.ThrowIfNull(worksheet);
        ArgumentNullException.ThrowIfNull(table);
        ArgumentNullException.ThrowIfNull(body);

        var firstBodyRow = GetRangeRow(body);
        var lastBodyRow = rowCount > 0 && firstBodyRow > 0
            ? firstBodyRow + rowCount - 1
            : 0;
        var candidate = lastBodyRow > 0 ? lastBodyRow + 1 : (int?)null;

        return GanttSheetLayout.ResolveAnchorRow(
            firstBodyRow,
            lastBodyRow,
            candidate,
            candidate is { } row ? IsLayoutRowEmpty(worksheet, row) : null);
    }

    /// <summary>
    /// Reads a whole worksheet row's first worksheet index. Test seam over the COM
    /// <c>Range.Row</c> property.
    /// </summary>
    /// <param name="range">The range whose first row index is wanted.</param>
    /// <returns>The one-based worksheet row index.</returns>
    internal virtual int GetRangeRow(Excel.Range range)
    {
        ArgumentNullException.ThrowIfNull(range);
        return range.Row;
    }

    /// <summary>
    /// Whether a whole worksheet row carries no content.
    /// </summary>
    /// <param name="worksheet">The Gantt worksheet.</param>
    /// <param name="rowIndex">The one-based worksheet row.</param>
    /// <returns>
    /// <see langword="true"/> when the row was read and holds nothing;
    /// <see langword="false"/> when it holds something; <see langword="null"/> when
    /// the host would not report it.
    /// </returns>
    /// <remarks>
    /// <b>Three-valued on purpose.</b> "Not read" and "empty" are different facts,
    /// and collapsing them is how an unverified row came to be treated as the chart's
    /// margin. Only a genuine empty row resolves.
    /// </remarks>
    internal virtual bool? IsLayoutRowEmpty(Excel.Worksheet worksheet, int rowIndex)
    {
        ArgumentNullException.ThrowIfNull(worksheet);
        Excel.Range? row = GetLayoutRow(worksheet, rowIndex);
        if (row is null)
        {
            return null;
        }

        List<object?[]> values = ExcelValue2Matrix.ReadRows(GetRangeValue2(row));
        return values.Count == 1 && values[0].All(IsBlankCellValue);
    }

    /// <summary>Whether one cell value counts as blank for the padding-row check.</summary>
    /// <param name="value">The raw cell value.</param>
    /// <returns>Whether the cell is blank.</returns>
    private static bool IsBlankCellValue(object? value) =>
        value is null
        || value is System.Reflection.Missing
        || value is DBNull
        || (value is string text && string.IsNullOrWhiteSpace(text));

    /// <summary>
    /// Reads a range's <c>Value2</c>. Test seam over the COM parameterised member.
    /// </summary>
    /// <param name="range">The range to read.</param>
    /// <returns>The raw value.</returns>
    internal virtual object? GetRangeValue2(Excel.Range range)
    {
        ArgumentNullException.ThrowIfNull(range);
        return range.Value2;
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
