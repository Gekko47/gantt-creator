using System.Globalization;
using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live adapter that rebuilds the native Excel outline of the Gantt table from the
/// validated hierarchy (R4.7D, ADR-0026 D5).
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
/// <b>Body rows map to worksheet rows.</b> The plan works in
/// <see cref="GanttEvent.RowNumber"/> -- the body-row index -- because only the
/// Office layer knows where the table sits. This adapter resolves the body's first
/// worksheet row once and offsets, so the plan never has to know about Excel
/// coordinates.
/// </para>
/// <para>
/// <b>Why collapse state is not written.</b> <c>Group</c> is called on the child
/// range only. The parent row's <c>OutlineLevel</c> and the group's
/// <c>ShowLevels</c> are left alone, so a collapsed parent stays collapsed across a
/// Refresh instead of being re-expanded under the user.
/// </para>
/// <para>
/// The <c>internal virtual</c> seams isolate the COM parameterised properties for
/// contract tests; the real grouping round-trip is exercised by the tagged
/// live-Office integration test.
/// </para>
/// </remarks>
public class ExcelOutlineGroupWriter(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null) : IOutlineGroupPort
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IWorksheetProtectionGuard _protectionGuard = protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <summary>The outline level of a row that is not inside any group.</summary>
    internal const int TopLevelOutlineLevel = 1;

    /// <summary>The outline level of a child row, one group beneath its parent.</summary>
    internal const int ChildOutlineLevel = 2;

    /// <inheritdoc />
    public OutlineGroupOutcome Apply(IReadOnlyList<GanttEvent> events)
    {
        ArgumentNullException.ThrowIfNull(events);

        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (workbook is null)
        {
            return OutlineGroupOutcome.Refused(OutlineGroupRefusalReason.NoActiveWorkbook);
        }

        if (!TryFindTable(workbook.Sheets, out Excel.Worksheet? worksheet, out Excel.ListObject? table)
            || worksheet is null
            || table is null)
        {
            return OutlineGroupOutcome.Refused(OutlineGroupRefusalReason.TableMissing);
        }

        // First check in a mutating adapter (ADR-0008 D4), before any COM write.
        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(worksheet);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return OutlineGroupOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? OutlineGroupRefusalReason.NoActiveWorkbook
                    : OutlineGroupRefusalReason.TargetProtected);
        }

        Excel.Range? body = GetTableBody(table);
        if (body is null)
        {
            return OutlineGroupOutcome.Ok(0, 0);
        }

        if (ReadFirstWorksheetRow(body) is not { } firstRow)
        {
            return OutlineGroupOutcome.Ok(0, 0);
        }

        OutlineGroupPlan plan = OutlineGroupPlanner.Plan(events);

        // Stale groups are cleared first: Excel refuses to regroup a sheet whose
        // outline disagrees with the request. The two lists never overlap, so the
        // order of these phases cannot change the outcome.
        var ungrouped = 0;
        foreach (var bodyRow in plan.RowsToUngroup)
        {
            if (Ungroup(worksheet, firstRow, bodyRow))
            {
                ungrouped++;
            }
        }

        var applied = 0;
        foreach (OutlineGroup group in plan.Groups)
        {
            if (group.ChildRowNumbers.Count > 0
                && Group(worksheet, firstRow, group.ChildRowNumbers[0], group.ChildRowNumbers[^1]))
            {
                applied++;
            }
        }

        return OutlineGroupOutcome.Ok(applied, ungrouped);
    }

    /// <summary>
    /// Places the inclusive body-row span at outline level 2, which is what
    /// Excel shows as one collapsible group beneath its level-1 parent.
    /// </summary>
    /// <param name="worksheet">The target worksheet.</param>
    /// <param name="firstWorksheetRow">The worksheet row of body row 1.</param>
    /// <param name="firstChildBodyRow">The first child body row.</param>
    /// <param name="lastChildBodyRow">The last child body row.</param>
    /// <returns>Whether the host accepted the write.</returns>
    /// <remarks>
    /// <b>Why the outline level is written rather than <c>Group</c> being
    /// called.</b> The row-outlining form of <c>Group</c> is reached through the
    /// <c>Rows</c> collection, which this PIA surfaces as a plain
    /// <see cref="Excel.Range"/>; the only <c>Group</c> on that type is the
    /// PivotTable signature <c>(Start, End, By, Periods)</c>, verified by
    /// reflection over the installed <c>Microsoft.Office.Interop.Excel</c>
    /// 14.0.1 assembly. Calling it positionally would send the level as a
    /// PivotTable start value rather than as an outline level.
    /// <c>Range.OutlineLevel</c> is a read/write <c>Object</c> on the same type
    /// (also verified by reflection), and setting it to 2 on a contiguous row
    /// range is the same operation the row-outlining <c>Group</c> performs, with no
    /// overload ambiguity. If a future interop package restores the row-outlining
    /// overload, this is the one place to revisit.
    /// <para>
    /// Only the child rows are touched, so the parent's own level and the group's
    /// collapsed/expanded state are preserved.
    /// </para>
    /// </remarks>
    private bool Group(Excel.Worksheet worksheet, int firstWorksheetRow, int firstChildBodyRow, int lastChildBodyRow) =>
        SetOutlineLevel(
            worksheet,
            firstWorksheetRow + firstChildBodyRow - 1,
            firstWorksheetRow + lastChildBodyRow - 1,
            ChildOutlineLevel);

    /// <summary>Returns a body row to outline level 1, clearing any stale group.</summary>
    /// <param name="worksheet">The target worksheet.</param>
    /// <param name="firstWorksheetRow">The worksheet row of body row 1.</param>
    /// <param name="bodyRow">The body row to ungroup.</param>
    /// <returns>Whether the host accepted the write.</returns>
    private bool Ungroup(Excel.Worksheet worksheet, int firstWorksheetRow, int bodyRow) =>
        SetOutlineLevel(
            worksheet,
            firstWorksheetRow + bodyRow - 1,
            firstWorksheetRow + bodyRow - 1,
            TopLevelOutlineLevel);

    /// <summary>Writes an outline level across an inclusive worksheet row range.</summary>
    /// <param name="worksheet">The target worksheet.</param>
    /// <param name="firstRow">The first worksheet row.</param>
    /// <param name="lastRow">The last worksheet row.</param>
    /// <param name="level">The outline level to write.</param>
    /// <returns>Whether the host accepted the write.</returns>
    private bool SetOutlineLevel(Excel.Worksheet worksheet, int firstRow, int lastRow, int level) =>
        GetRowsRange(worksheet, firstRow, lastRow) is { } range && WriteOutlineLevel(range, level);

    /// <summary>Finds the Gantt worksheet and its table. Test seam over the COM collection indexers.</summary>
    /// <param name="sheets">The workbook's sheet collection.</param>
    /// <param name="worksheet">The resolved worksheet.</param>
    /// <param name="table">The resolved table.</param>
    /// <returns>Whether both were found.</returns>
    internal virtual bool TryFindTable(
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
                if (string.Equals(candidateTable.Name, "tblGanttData", StringComparison.Ordinal))
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

    /// <summary>Reads the worksheet row number of a range's first row. Test seam over <c>Range.Row</c>.</summary>
    /// <param name="range">The range.</param>
    /// <returns>The first worksheet row, or <see langword="null"/> when the host reports none.</returns>
    internal virtual int? ReadFirstWorksheetRow(Excel.Range range)
    {
        ArgumentNullException.ThrowIfNull(range);

        return range.Row is int row ? row : null;
    }

    /// <summary>Resolves an inclusive worksheet row range. Test seam over the range indexer.</summary>
    /// <param name="worksheet">The target worksheet.</param>
    /// <param name="firstRow">The first worksheet row.</param>
    /// <param name="lastRow">The last worksheet row.</param>
    /// <returns>The range, or <see langword="null"/> when the host does not resolve it.</returns>
    /// <remarks>
    /// <para>
    /// The span is selected by its whole-row A1 address (<c>"5:7"</c>), which spans
    /// every column of those rows. The previous form used the two-argument
    /// <c>Rows</c> indexer, and reflection over the installed
    /// <c>Microsoft.Office.Interop.Excel</c> 14.0.1 shows that indexer is
    /// <c>Item(RowIndex, ColumnIndex)</c> -- a single cell at
    /// <c>(firstRow, lastRow)</c>, not the rows between them. Every
    /// <c>WriteOutlineLevel</c> therefore reached exactly one cell, so a group of
    /// several children outlined a single cell and every other row in the group kept
    /// whatever level it already had.
    /// </para>
    /// <para>
    /// The whole row span is written rather than one cell per row because
    /// <c>OutlineLevel</c> is a row property: one write across the span is both
    /// correct and cheaper than N row writes. This PIA exposes no four-argument
    /// <c>Cells</c> indexer (verified by reflection: <c>Range</c> carries only the
    /// two-argument <c>Item</c>), so the address form is the way to name the span.
    /// </para>
    /// </remarks>
    internal virtual Excel.Range? GetRowsRange(Excel.Worksheet worksheet, int firstRow, int lastRow)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        return worksheet.Range[RowSpanAddress(firstRow, lastRow)];
    }

    /// <summary>
    /// Formats the whole-row A1 address for an inclusive row span, invariantly and
    /// without a sheet qualifier (the address is resolved against the worksheet it
    /// is read from). <c>"5:7"</c> spans every column of rows 5 through 7.
    /// </summary>
    /// <param name="firstRow">The first worksheet row.</param>
    /// <param name="lastRow">The last worksheet row.</param>
    /// <returns>The row-span address.</returns>
    internal static string RowSpanAddress(int firstRow, int lastRow) =>
        string.Create(CultureInfo.InvariantCulture, $"{firstRow}:{lastRow}");

    /// <summary>
    /// Writes an outline level to a row range. Test seam over the COM
    /// <c>Range.OutlineLevel</c> setter, which is <c>Object</c>-typed.
    /// </summary>
    /// <param name="range">The rows to write.</param>
    /// <param name="level">The outline level to write.</param>
    /// <returns>Whether the host accepted the write.</returns>
    /// <remarks>
    /// The value is boxed as <see cref="int"/> because <c>OutlineLevel</c> is
    /// <c>Object</c> in the interop assembly. The real write is exercised by the
    /// tagged live-Office integration test; a contract test can only observe that
    /// this seam is reached with the intended level.
    /// </remarks>
    internal virtual bool WriteOutlineLevel(Excel.Range range, int level)
    {
        ArgumentNullException.ThrowIfNull(range);

        range.OutlineLevel = level;
        return true;
    }
}
