using System.Globalization;
using System.Reflection;
using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live guarded row inserter for the visible <c>tblGanttData</c> table.
/// Inserts one scaffold row relative to the active table cell, with deterministic
/// append fallback, and never renders or changes any other workbook state.
/// </summary>
public class ExcelGanttRowInserter(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null,
    ITypeOptionsMaterialiser? typeOptionsMaterialiser = null,
    IRowHeightNormalisationPort? rowHeightNormaliser = null,
    IOutlineGroupPort? outlineGroups = null,
    Action<string>? trace = null) : IGanttRowInserter
{
    private readonly Excel.Application? _application = application as Excel.Application;

    /// <summary>
    /// Optional diagnostic sink for the group-aware parent decision and the regroup that follows
    /// it, or <see langword="null"/> for no trace.
    /// </summary>
    /// <remarks>
    /// Both lines exist because the decision and its consequence are one row apart in time but
    /// several components apart in code: the inserter chooses a <c>ParentId</c>, writes it, then
    /// the outline adapter rebuilds groups from a fresh read. A user report that "the new activity
    /// is not in the group" cannot distinguish "the parent was never assigned" from "the parent was
    /// assigned and the planner still dropped it", and the second is invisible in the sheet.
    /// </remarks>
    private readonly Action<string>? _trace = trace;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);
    private readonly ITypeOptionsMaterialiser _typeOptionsMaterialiser =
        typeOptionsMaterialiser ?? new ExcelTypeOptionsMaterialiser(application);

    /// <summary>
    /// The outline adapter, consulted after a successful positional
    /// insert so the native <c>-</c>/<c>+</c> controls reflect the
    /// new row's parent without waiting for the next Refresh (R5.13).
    /// Null in tests and when the host supplied no outline writer.
    /// </summary>
    private readonly IOutlineGroupPort? _outlineGroups = outlineGroups;

    /// <summary>
    /// The verified row-height normaliser, which owns the reserved rows below the body.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>Why the inserter normalises at all (ADR-0038 D1).</b> The reserved rows below
    /// the body were only ever touched through Refresh, so on the Add-Activity path they
    /// kept Excel's default height. Measured live 2026-10-03: after Initialise plus
    /// three appends every row below the table read <b>15pt</b> -- so the anchor row was
    /// a 15pt strip instead of 0.25pt and the chart's bottom margin was ~30pt instead
    /// of 6pt. The anchor still worked as a cell anchor, so the bands stretched, which
    /// is exactly why the defect survived a green suite: it was invisible to every
    /// geometry assertion and obvious on screen.
    /// </para>
    /// <para>
    /// <b>Why this port rather than writing heights here.</b> The inserter performs the
    /// worksheet-row insert, so it is the component that knows which rows moved -- but it
    /// does not own row heights, and re-deriving "the row below the body" here would be
    /// the second authority ADR-0035 D2 was raised to remove. The normaliser resolves
    /// both rows from the body's MEASURED span and REFUSES rather than writing a row it
    /// cannot verify, which is the discipline this needs.
    /// </para>
    /// </remarks>
    private readonly IRowHeightNormalisationPort _rowHeightNormaliser =
        rowHeightNormaliser ?? new ExcelRowHeightNormaliser(application);

    /// <inheritdoc />
    public GanttRowInsertOutcome Insert(GanttEntityType type, Func<GanttRowId> nextId)
    {
        // ADR-0008 D4: this is intentionally the first operation in the
        // mutating entry point, before any application/workbook/table access.
        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return GanttRowInsertOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? GanttRowInsertRefusalReason.NoActiveWorkbook
                    : GanttRowInsertRefusalReason.TargetProtected);
        }

        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (workbook is null)
        {
            return GanttRowInsertOutcome.Refused(GanttRowInsertRefusalReason.NoActiveWorkbook);
        }

        Excel.Sheets sheets = workbook.Sheets;
        if (!TryFindTable(sheets, out Excel.Worksheet? worksheet, out Excel.ListObject? table) || worksheet is null || table is null
            || !TryBuildColumnMap(table, out int[]? columnMap) || columnMap is null)
        {
            return GanttRowInsertOutcome.Refused(GanttRowInsertRefusalReason.TableMissing);
        }

        ProtectionGuardOutcome targetProtection = _protectionGuard.QueryTarget(worksheet);
        if (targetProtection != ProtectionGuardOutcome.NotProtected)
        {
            return GanttRowInsertOutcome.Refused(
                targetProtection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? GanttRowInsertRefusalReason.NoActiveWorkbook
                    : GanttRowInsertRefusalReason.TargetProtected);
        }

        IReadOnlyList<object?> values = GanttRowDefaults.Build(type, nextId);
        Excel.ListRows rows = GetListRows(table);
        Excel.Range? reusableRow = GetListRowCount(rows) == 0 ? GetReusableInitialBlankRow(table) : null;

        // ADR-0035 D3 (superseding the append-only form), as corrected by ADR-0036 D5:
        // the new row goes immediately BELOW the active cell when the active cell is
        // inside the table, and APPENDS before the padding row otherwise. BOTH branches
        // now insert a genuine worksheet row, because ListRows.Add(position) does not
        // shift the sheet (probe Q1: delta=0 on every shape).
        //
        // Shape tracking needs NO write from this adapter. Measured 2026-10-03
        // (scripts/probe-rowinsert-anchoring.ps1): all four creation families the
        // renderer uses -- AddShape, AddLine, AddTextbox, and the milestone diamond --
        // report Placement = xlMoveAndSize straight out of Excel, and inserting a
        // worksheet row moved every one of them down exactly one body row height with
        // TopLeftCell following. ADR-0035 reached the same conclusion from an
        // AddTextbox-only probe; this probe closes the AddShape/AddLine gap that
        // probe left open.
        Excel.ListRow? newRow = null;
        Excel.Range rowRange;

        // Hoisted out of the insert block so the regroup can report WHICH body row it inserted:
        // the plan's refusal list is a set of row numbers, and without this the log cannot say
        // whether the new row was refused or some unrelated row was.
        int? insertedBodyRow = null;

        // Only the APPEND branch can lose the padding row, because only it depends on
        // the insert below the table having succeeded. A positional insert shifts the
        // sheet by itself, so its padding row is displaced by definition.
        bool paddingRowReserved = true;

        // R5.16: resolve the new row's parent BEFORE any worksheet mutation, so a
        // Critical Interval with no eligible parent refuses without touching the
        // sheet. The decision reads the body only when there is something to
        // decide (see DecideParent), which keeps the append and no-outline paths
        // read-free.
        int? position = reusableRow is null
            ? GetInsertionPosition(application!, table, rows)
            : null;
        ParentDecision parentDecision = DecideParent(type, table, position, columnMap);
        if (parentDecision.Refused)
        {
            return GanttRowInsertOutcome.Refused(
                GanttRowInsertRefusalReason.CriticalIntervalRequiresParent);
        }

        position = parentDecision.Position;
        if (parentDecision.ParentId is not null)
        {
            values = WithParentId(values, parentDecision.ParentId);
        }

        if (reusableRow is not null)
        {
            // An initialised table whose body is empty can still own one blank row
            // that Excel does not count as a ListRow. Reusing it keeps the first
            // insert from leaving a trailing blank row inside the table, and is
            // independent of WHERE the row goes.
            rowRange = reusableRow;
        }
        else
        {
            insertedBodyRow = position;

            // BOTH branches insert a genuine WORKSHEET row, because
            // ListRows.Add(position) does NOT shift the sheet. Measured
            // 2026-10-03 (scripts/probe-positional-insert.ps1, Q1): adding at
            // position 2 moved all three probe shapes by delta=0 and pushed the row
            // below the table from 6pt to 15pt -- it rearranges rows INSIDE the
            // table and consumes the padding row, leaving every shape and cell
            // below untouched. That is the reported symptom exactly: the table's
            // row data moves, the cells and Gantt shapes do not.
            //
            // A real worksheet row insert does both. Q2 of the same probe:
            // Rows(3).Insert(xlShiftDown) inside the table's range moved all three
            // shapes by delta=18 with TopLeftCell following, AND auto-expanded the
            // ListObject into the new row (ListRows 3 -> 4) whose cell accepted a
            // written value (Q3). So the positional branch needs no ListRows call at
            // all -- the table absorbs the inserted row itself.
            //
            // This corrects ADR-0035 D3, which asserted the positional branch
            // "already inserts a real worksheet row". It does not.
            bool insideTable = position is not null;

            // Inside the table for a positional insert, so the ListObject absorbs the
            // row. One past the table's last row for an append, so the padding row is
            // displaced rather than consumed -- Q4 measured that a row inserted below
            // the table does NOT auto-expand it, which is why the append branch still
            // follows its insert with ListRows.Add().
            bool inserted = insideTable
                ? InsertWorksheetRow(
                    table,
                    GetRangeRow(GetTableRange(table)) + position!.Value)
                : InsertWorksheetRowBelowTable(table);

            // A REFUSED POSITIONAL insert is not a degraded insert, it is the wrong
            // row. `GetListRowAt(rows, position)` reads whatever row already occupies
            // that position, so continuing would hand the host a bulk write aimed at
            // the USER'S EXISTING ROW: the scaffold values would overwrite real
            // schedule data and report success. This is the only path where a refusal
            // can corrupt, so it refuses before the read.
            //
            // The append branch deliberately does NOT refuse here (ADR-0035 D3): its
            // `AddRow` creates a genuine new row even when the push-down below the
            // table failed, so the row exists and only the padding row was lost --
            // which `paddingRowReserved` reports instead.
            if (insideTable && !inserted)
            {
                return GanttRowInsertOutcome.Refused(GanttRowInsertRefusalReason.RowInsertRefused);
            }

            paddingRowReserved = insideTable || inserted;

            try
            {
                newRow = insideTable ? GetListRowAt(rows, position!.Value) : AddRow(rows);

                rowRange = GetRowRange(newRow);

                // The height is NOT right by Excel's format inheritance: the insert copies
                // from the row BELOW (xlFormatFromRightOrBelow), so an appended row inherits
                // the 6pt reserved row's height. Writing the token is therefore load-bearing
                // here, not merely this add-in taking ownership of a host behaviour.
                SetRowHeight(rowRange, GanttCatalogues.MetricDefault("GanttRowHeightPt"));
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                // Creating or sizing the row is the same class of host refusal as
                // writing it, and it is guarded on the same terms (ADR-0008 D4). These
                // three calls were unguarded while `WriteRow` was not, so a refusal
                // here escaped into the Ribbon callback with the worksheet row already
                // inserted -- the sequence stopped partway and nothing said so.
                //
                // `RowWriteRefused` is the honest reason: the row IS in the worksheet,
                // because the insert that precedes these calls succeeded.
                _ = ex;
                return GanttRowInsertOutcome.Refused(GanttRowInsertRefusalReason.RowWriteRefused);
            }

            // R5.16: the parent (if any) was resolved BEFORE the insert by
            // DecideParent and is already written into `values`. Nothing here
            // reads the user's rows after the row-creation seams.
        }

        try
        {
            WriteRow(rowRange, values, columnMap);
        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            // The worksheet row is ALREADY inserted at this point. Reporting this as
            // "the row was not added" would be a claim the user can disprove by looking
            // at the sheet, so it is reported as added-but-not-completed instead.
            //
            // This is the guard the sibling helper has always had (ADR-0008 D4: a host
            // refusal is reported, never thrown into the Ribbon callback). GetListRowAt,
            // AddRow and SetRowHeight were left unguarded, and a live workbook produced
            // exactly that: the insert sequence aborted partway and the plot was left
            // at its old extent, which reads as "the bands did not stretch".
            _ = ex;
            return GanttRowInsertOutcome.Refused(GanttRowInsertRefusalReason.RowWriteRefused);
        }

        // R5.13: rebuild the native outline from the post-insert table so the
        // -/+ controls reflect the new row's parent immediately, mirroring the
        // parent/child mutator's Finish. A regroup refusal degrades to the
        // pre-D12 behaviour -- the insert stands and the next Refresh rebuilds
        // the outline -- rather than failing work that already landed.
        RegroupOutline(table, columnMap, insertedBodyRow);

        // The materialiser is COM work too: it re-applies the Type column's list
        // validation when the name or the rule has gone stale, and the host refuses
        // `Validation.Add` on some workbooks with 0x800A03EC. This call was the last
        // unguarded COM entry in this method, and a live workbook proved what that
        // costs: the exception escaped into the Ribbon callback AFTER the row was
        // written, so the user got the generic "Excel could not complete the
        // operation" dialog, a scaffold row with no dates, and a refresh that
        // reported two validation errors against it. A host refusal here is now the
        // typed `TypeOptionsUnavailable`, and it takes the rollback below -- which is
        // what this block always intended to do.
        bool typeOptionsSucceeded;
        try
        {
            typeOptionsSucceeded = _typeOptionsMaterialiser.EnsureCurrent().Succeeded;
        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            _ = ex;
            typeOptionsSucceeded = false;
        }

        if (!typeOptionsSucceeded)
        {
            if (newRow is not null)
            {
                DeleteRow(newRow);
            }
            else
            {
                ClearRange(rowRange);
            }

            return GanttRowInsertOutcome.Refused(GanttRowInsertRefusalReason.TypeOptionsUnavailable);
        }

        // ADR-0038 D1: restore the reserved rows below the body NOW, not at the next
        // Refresh. Every insert displaces them, so this is the point at which their
        // heights are wrong; leaving it to Refresh meant the sheet sat wrong for as
        // long as the user did not refresh, which the live gate caught as 15pt rows.
        //
        // Deliberately AFTER WriteRow and the Type options: the body must already
        // include the new row, because the normaliser derives the anchor and padding
        // rows from the MEASURED body span. Doing it earlier would resolve them
        // against the pre-insert body and write to the wrong rows.
        //
        // A COMException here is a host refusal, not a fault: the row is written and
        // must not be reported as a failed insert. It is caught here and reported as an
        // un-normalised margin -- the same fact the normaliser's own typed refusal
        // produces -- so a refusal cannot escape `Insert` after the row is on the
        // sheet.
        bool reservedRowsNormalised;
        try
        {
            reservedRowsNormalised = _rowHeightNormaliser.Normalise(
                GanttCatalogues.MetricDefault("GanttRowHeightPt"),
                GanttCatalogues.MetricDefault("SplitterHeightPt"),
                GanttCatalogues.MetricDefault("SpacerHeightPt"),
                GanttCatalogues.MetricDefault("PeriodBandHeightPt"),
                GanttCatalogues.MetricDefault("YearBandHeightPt"),
                GanttCatalogues.MetricDefault("ChartPaddingRowHeightPt"),
                GanttCatalogues.MetricDefault("ChartAnchorRowHeightPt")).Succeeded;
        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            _ = ex;
            reservedRowsNormalised = false;
        }

        return GanttRowInsertOutcome.Ok(
            newRow is null ? 1 : GetRowIndex(newRow),
            paddingRowReserved,
            reservedRowsNormalised);
    }

    /// <summary>
    /// Resolves where the new row goes, from the active cell (ADR-0035 D3).
    /// </summary>
    /// <param name="application">The application, for the active cell.</param>
    /// <param name="table">The Gantt table.</param>
    /// <param name="rows">The table's list rows.</param>
    /// <returns>
    /// The one-based list position to insert at, or <see langword="null"/> to APPEND
    /// at the end of the body.
    /// </returns>
    /// <remarks>
    /// <para>
    /// This restores the R2.8 rule, which ADR-0035's append-only form had removed.
    /// The owner's ruling is conditional and both branches are load-bearing:
    /// an active cell INSIDE the table inserts immediately below it, and an active
    /// cell anywhere else -- another sheet, the header, the padding row -- appends.
    /// </para>
    /// <para>
    /// The active row is converted to a body index by subtracting the TABLE's first
    /// row, which is the header. That makes the arithmetic one-based already, so
    /// <c>bodyIndex + 1</c> is "the row after the selected one". A body index of 0
    /// means the header was selected and the new row goes to the top.
    /// </para>
    /// </remarks>
    private int? GetInsertionPosition(
        Excel.Application application,
        Excel.ListObject table,
        Excel.ListRows rows)
    {
        Excel.Range? activeCell = GetActiveCell(application);
        if (!GetTableActive(table) || activeCell is null)
        {
            return null;
        }

        Excel.Range tableRange = GetTableRange(table);
        int tableFirstRow = GetRangeRow(tableRange);
        int activeRow = GetRangeRow(activeCell);
        int bodyIndex = activeRow - tableFirstRow;
        int rowCount = GetListRowCount(rows);

        // Outside the body in either direction: not a position the user can name.
        if (bodyIndex < 0 || bodyIndex > rowCount)
        {
            return null;
        }

        if (bodyIndex == 0)
        {
            return 1;
        }

        // The LAST body row appends rather than inserting below itself, which would
        // be the same place by a different route.
        return bodyIndex == rowCount ? null : bodyIndex + 1;
    }

    // Excel's own values, carried as named constants rather than PIA enums: the
    // interop assembly Office compiles against does not expose XlInsertionShift or
    // XlInsertFormatOrigin, and these are the exact values a live probe drove
    // Rows.Insert with successfully.
    private const int _excelShiftDown = -4121; // xlShiftDown

    // 1 = XlInsertFormatOrigin.xlFormatFromRightOrBelow, NOT the default 0
    // (xlFormatFromLeftOrAbove). The default copies the format of the row ABOVE into
    // the inserted row, and the row above an appended activity is the last body row,
    // which carries the pastel-red Start/Finish fills of the validation report. Excel
    // copies cell interiors with the format, so every add cloned that pink into the
    // row it created -- exactly one row lower than the previous clone, which is the
    // reported "a warning fill appears one row below the entity, and the next add puts
    // another one below that" trail. This constant previously read -4142, which is not
    // a member of XlInsertFormatOrigin at all; Excel ignored it, fell back to the
    // default, and cloned the fill. Measured 2026-11-09 against Excel 16.0.
    private const int _excelFormatFromRightOrBelow = 1;

    /// <summary>
    /// Inserts one real worksheet row directly below the table, so the reserved
    /// bottom padding row moves down instead of being absorbed (ADR-0035 D3).
    /// </summary>
    /// <param name="table">The Gantt table.</param>
    /// <returns>
    /// <see langword="true"/> when a blank row now sits below the table; otherwise
    /// <see langword="false"/>, meaning the padding row was absorbed into the table.
    /// </returns>
    /// <remarks>
    /// <para>
    /// This used to swallow every failure and return, so a refused insert was
    /// indistinguishable from a successful one. That is the failure mode behind "the
    /// padding row still takes data": the append had already claimed the padding row,
    /// the push-down silently did not happen, and the command reported success
    /// regardless.
    /// </para>
    /// <para>
    /// The outcome is reported rather than thrown because the inserted row is already
    /// written and the user asked for it; throwing here would escape into the Ribbon
    /// callback with a half-applied insert. Returning <see langword="false"/> lets the
    /// caller tell the user the margin was lost, which ADR-0008 requires in place of a
    /// silent partial mutation.
    /// </para>
    /// <para>
    /// CA1031: an Excel host refusing a row insert is a host refusal, not a bug in
    /// this adapter, and must not surface as an exception into the Ribbon callback.
    /// </para>
    /// <para>
    /// The APPEND branch only. It must run BEFORE <c>ListRows.Add()</c>: called after,
    /// it cannot repair the damage, because <c>Add()</c> would already have claimed
    /// the padding row and turned it into a squashed 6pt body row.
    /// </para>
    /// </remarks>
    internal virtual bool InsertWorksheetRowBelowTable(Excel.ListObject table)
    {
        ArgumentNullException.ThrowIfNull(table);

        Excel.Range? tableRange = GetTableRange(table);
        if (tableRange is null)
        {
            return false;
        }

        // One past the table's last row. Reading it through the range rather than
        // assuming the layout keeps this correct if the table ever adopts at a
        // different row.
        int lastRow = GetRangeRow(tableRange) + (tableRange.Rows?.Count ?? 0) - 1;
        return InsertWorksheetRow(table, lastRow + 1);
    }

    /// <summary>
    /// Inserts one genuine worksheet row at <paramref name="worksheetRow"/>, shifting
    /// everything below it down (ADR-0036 D1/D5).
    /// </summary>
    /// <param name="table">The Gantt table, used to reach the owning worksheet.</param>
    /// <param name="worksheetRow">The one-based worksheet row to insert at.</param>
    /// <returns>
    /// <see langword="true"/> when the row was inserted; otherwise
    /// <see langword="false"/>, meaning the host refused.
    /// </returns>
    /// <remarks>
    /// <para>
    /// This is the ONLY insert on either branch that moves the sheet. A row inserted
    /// INSIDE the table's range is absorbed by the <c>ListObject</c>, which grows by
    /// one <c>ListRow</c>; a row inserted below it is not, so the caller pairs that
    /// case with <c>ListRows.Add()</c>.
    /// </para>
    /// <para>
    /// Deliberately NO height write. Inside the table the inserted row becomes a body
    /// row; below it, the row about to be created is the one <c>ListRows.Add()</c>
    /// claims. Either way the caller writes the body height, and writing the padding
    /// height here would reproduce the defect this method exists to prevent.
    /// </para>
    /// <para>
    /// CA1031: a host refusal is reported, not thrown, so it cannot escape into the
    /// Ribbon callback with a half-applied insert.
    /// </para>
    /// </remarks>
    internal virtual bool InsertWorksheetRow(Excel.ListObject table, int worksheetRow)
    {
        ArgumentNullException.ThrowIfNull(table);
        ArgumentOutOfRangeException.ThrowIfLessThan(worksheetRow, 1);

#pragma warning disable CA1031
        try
        {
            if (table.Parent is not Excel.Worksheet worksheet || worksheet.Rows is not { } rows)
            {
                return false;
            }

            rows[worksheetRow].Insert(_excelShiftDown, _excelFormatFromRightOrBelow);
            return true;
        }
        catch (Exception ex) when (ex is System.Runtime.InteropServices.COMException or InvalidCastException or ArgumentException)
        {
            // Reported as "not reserved" rather than thrown; see the remarks.
            _ = ex;
            return false;
        }
#pragma warning restore CA1031
    }

    private static bool IsBlankValue(object? value) =>
        value is null
        || value is Missing
        || value is DBNull
        || (value is string text && string.IsNullOrWhiteSpace(text));

    private Excel.Range? GetReusableInitialBlankRow(Excel.ListObject table)
    {
        Excel.Range tableRange = GetTableRange(table);
        Excel.Range tableRows = GetRangeRows(tableRange);
        int rowCount = GetRangeRowCount(tableRows);
        if (rowCount <= 1)
        {
            return null;
        }

        Excel.Range candidate = GetRangeAt(tableRows, rowCount);
        List<object?[]> values = ExcelValue2Matrix.ReadRows(GetRangeValue2(candidate));
        return values.Count == 1
            && values[0].Length > 0
            && values[0].All(IsBlankValue)
            ? candidate
            : null;
    }

    private bool TryFindTable(
        Excel.Sheets sheets,
        out Excel.Worksheet? worksheet,
        out Excel.ListObject? table)
    {
        worksheet = null;
        table = null;
        int sheetCount = sheets.Count;
        for (int sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
        {
            if (GetSheetAt(sheets, sheetIndex) is not Excel.Worksheet candidate)
            {
                continue;
            }

            Excel.ListObjects listObjects = candidate.ListObjects;
            int tableCount = listObjects.Count;
            for (int tableIndex = 1; tableIndex <= tableCount; tableIndex++)
            {
                Excel.ListObject candidateTable = GetTableAt(listObjects, tableIndex);
                if (string.Equals(candidateTable.Name, GanttTableSchema.TableName, StringComparison.OrdinalIgnoreCase))
                {
                    worksheet = candidate;
                    table = candidateTable;
                    return true;
                }
            }
        }

        return false;
    }

    private bool TryBuildColumnMap(Excel.ListObject table, out int[]? columnMap)
    {
        columnMap = null;
        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        int[] map = new int[schema.Count];
        Dictionary<string, int> byName = new Dictionary<string, int>(StringComparer.Ordinal);
        Excel.ListColumns columns = table.ListColumns;
        int columnCount = columns.Count;
        for (int index = 1; index <= columnCount; index++)
        {
            Excel.ListColumn column = GetColumnAt(columns, index);
            byName[column.Name] = index;
        }

        for (int schemaIndex = 0; schemaIndex < schema.Count; schemaIndex++)
        {
            GanttTableColumn expected = schema[schemaIndex];
            if (!byName.TryGetValue(expected.Name, out int tableIndex))
            {
                if (expected.IsRequired)
                {
                    return false;
                }

                map[schemaIndex] = -1;
            }
            else
            {
                map[schemaIndex] = tableIndex;
            }
        }

        columnMap = map;
        return true;
    }

    /// <summary>
    /// The parent decision for one insert (R5.16): the parent to write, the
    /// position to land at (null means append), and whether the insert must be
    /// refused before it mutates the sheet.
    /// </summary>
    private readonly record struct ParentDecision(
        GanttRowId? ParentId,
        int? Position,
        bool Refused);

    /// <summary>
    /// Resolves the new row's parent BEFORE the worksheet insert (R5.16).
    /// </summary>
    /// <remarks>
    /// <para>Four outcomes, in order:</para>
    /// <list type="number">
    /// <item><description>
    /// A Critical Interval attaches to the nearest eligible parent — searched
    /// upward from the selection first, then downward — and lands after that
    /// parent's last child. No eligible parent anywhere refuses the insert.
    /// </description></item>
    /// <item><description>
    /// Every other type keeps the R5.13 neighbour rule: inside a contiguous
    /// child run it joins that group; between a parent and its first child it
    /// becomes that parent's first child.
    /// </description></item>
    /// <item><description>
    /// The candidate parent is validated through
    /// <see cref="AddChildPlanner.Evaluate"/>. A rejection — or a type that may
    /// never be a child, such as a spacer or splitter — appends the row at the
    /// end of the table instead, so a child run is never split.
    /// </description></item>
    /// <item><description>
    /// Nothing to decide (an append, a header insert, or no outline writer)
    /// returns without reading the body at all.
    /// </description></item>
    /// </list>
    /// <para>
    /// This is the one place the insert reads before mutating. R5.16 records
    /// that exception, and the R5.13 "never reads" pin is amended to "never
    /// writes or regroups".
    /// </para>
    /// </remarks>
    /// <param name="type">The type the new row will have.</param>
    /// <param name="table">The Gantt data table, pre-insert.</param>
    /// <param name="position">The one-based body position the row would land at, or null for an append.</param>
    /// <param name="columnMap">The schema-column to table-column index map.</param>
    /// <returns>The decision.</returns>
    private ParentDecision DecideParent(
        GanttEntityType type,
        Excel.ListObject table,
        int? position,
        int[] columnMap)
    {
        bool critical = type == GanttEntityType.CriticalInterval;
        bool canJoinGroup = _outlineGroups is not null;

        // Nothing to decide: leave the append, header, and no-outline paths
        // exactly as they were, including their freedom from any body read.
        if (!critical && (!canJoinGroup || position is null || position.Value < 2))
        {
            return new ParentDecision(null, position, false);
        }

        List<object?[]> body = ReadBody(table);
        if (body.Count == 0)
        {
            return new ParentDecision(null, position, critical);
        }

        Dictionary<string, int> childCounts = CountChildren(body, columnMap);

        if (critical)
        {
            int anchor = position is null ? body.Count - 1 : position.Value - 2;
            if (anchor < 0 || anchor >= body.Count)
            {
                return new ParentDecision(null, position, true);
            }

            int? parentIndex = FindNearestEligibleParent(body, columnMap, anchor, childCounts, out GanttRowId? parentId);
            return parentIndex is null || parentId is null
                ? new ParentDecision(null, position, true)
                : new ParentDecision(
                    parentId,
                    PositionAfterChildren(body, columnMap, parentIndex.Value, parentId),
                    false);
        }

        int activeIndex = position!.Value - 2;
        if (activeIndex < 0 || activeIndex >= body.Count)
        {
            return new ParentDecision(null, position, false);
        }

        object?[] activeRow = body[activeIndex];
        object?[]? belowRow = position.Value < body.Count ? body[position.Value - 1] : null;

        GanttRowId? activeId = TryParseId(activeRow, columnMap);
        GanttRowId? activeParentId = TryParseParentId(activeRow, columnMap);
        GanttRowId? belowParentId = belowRow is null ? null : TryParseParentId(belowRow, columnMap);

        GanttRowId? candidate = null;
        if (activeParentId is not null && belowParentId == activeParentId)
        {
            candidate = activeParentId;
        }
        else if (activeParentId is null && activeId is not null && belowParentId == activeId)
        {
            candidate = activeId;
        }

        if (candidate is null)
        {
            return new ParentDecision(null, position, false);
        }

        int candidateIndex = -1;
        for (int index = 0; index < body.Count; index++)
        {
            if (TryParseId(body[index], columnMap) == candidate)
            {
                candidateIndex = index;
                break;
            }
        }

        if (candidateIndex < 0
            || !EntityTypeCatalog.TryParse(
                ExcelGanttRowProjection.CellText(body[candidateIndex], columnMap, "Type"),
                out GanttEntityType parentType))
        {
            return new ParentDecision(null, position, false);
        }

        object?[] parentRow = body[candidateIndex];
        string parentKey = candidate.Value;
        int existingChildren = childCounts.TryGetValue(parentKey, out int count) ? count : 0;
        AddChildRefusalReason? refusal = AddChildPlanner.Evaluate(
            parentType,
            candidate,
            TryParseParentId(parentRow, columnMap),
            existingChildren,
            type);

        // A rejected candidate, or a type that may never be a child, appends at
        // the end of the table: always a valid position, never splits a run.
        return refusal is null
            ? new ParentDecision(candidate, position, false)
            : new ParentDecision(null, null, false);
    }

    /// <summary>Counts each row's children by the parent id text it names.</summary>
    private static Dictionary<string, int> CountChildren(List<object?[]> body, int[] columnMap)
    {
        Dictionary<string, int> counts = new(StringComparer.Ordinal);
        foreach (object?[] row in body)
        {
            GanttRowId? parentId = TryParseParentId(row, columnMap);
            if (parentId is null)
            {
                continue;
            }

            string key = parentId.Value;
            counts[key] = counts.TryGetValue(key, out int count) ? count + 1 : 1;
        }

        return counts;
    }

    /// <summary>
    /// Finds the nearest row that can own a Critical Interval: upward from the
    /// anchor first, then downward (R5.16).
    /// </summary>
    private static int? FindNearestEligibleParent(
        List<object?[]> body,
        int[] columnMap,
        int anchor,
        Dictionary<string, int> childCounts,
        out GanttRowId? parentId)
    {
        parentId = null;
        for (int index = anchor; index >= 0; index--)
        {
            if (TryEligibleParent(body[index], columnMap, childCounts, out parentId))
            {
                return index;
            }
        }

        for (int index = anchor + 1; index < body.Count; index++)
        {
            if (TryEligibleParent(body[index], columnMap, childCounts, out parentId))
            {
                return index;
            }
        }

        return null;
    }

    /// <summary>Whether one row can own a Critical Interval.</summary>
    private static bool TryEligibleParent(
        object?[] row,
        int[] columnMap,
        Dictionary<string, int> childCounts,
        out GanttRowId? parentId)
    {
        parentId = null;
        GanttRowId? rowId = TryParseId(row, columnMap);
        if (rowId is null
            || !EntityTypeCatalog.TryParse(
                ExcelGanttRowProjection.CellText(row, columnMap, "Type"),
                out GanttEntityType rowType))
        {
            return false;
        }

        int childCount = childCounts.TryGetValue(rowId.Value, out int count) ? count : 0;
        if (AddChildPlanner.Evaluate(
                rowType,
                rowId,
                TryParseParentId(row, columnMap),
                childCount,
                GanttEntityType.CriticalInterval) is not null)
        {
            return false;
        }

        parentId = rowId;
        return true;
    }

    /// <summary>
    /// The position immediately after a parent's last existing child, or null
    /// when that is the end of the table (R5.16).
    /// </summary>
    private static int? PositionAfterChildren(
        List<object?[]> body,
        int[] columnMap,
        int parentIndex,
        GanttRowId parentId)
    {
        int lastIndex = parentIndex;
        for (int index = 0; index < body.Count; index++)
        {
            if (TryParseParentId(body[index], columnMap) == parentId)
            {
                lastIndex = index;
            }
        }

        return lastIndex == body.Count - 1 ? null : lastIndex + 2;
    }

    /// <summary>
    /// Reads the table's data body as a row matrix (R5.13), the same
    /// bulk read the projection authority and the regroup consume, so
    /// the inserter never re-derives a second copy of the body.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <returns>
    /// One array per body row, in body order; empty when the table
    /// has no body rows.
    /// </returns>
    private List<object?[]> ReadBody(Excel.ListObject table)
    {
        Excel.Range? dataBody = GetDataBodyRange(table);
        return dataBody is null
            ? []
            : ExcelValue2Matrix.ReadRows(GetRangeValue2(dataBody));
    }

    /// <summary>
    /// Parses a row's <c>Id</c> cell, or <see langword="null"/>
    /// when the cell is blank or unparseable.
    /// </summary>
    private static GanttRowId? TryParseId(object?[] row, int[] columnMap) =>
        GanttRowId.TryParse(
            ExcelGanttRowProjection.CellText(row, columnMap, "Id"),
            out GanttRowId? id) && id is not null
                ? id
                : null;

    /// <summary>
    /// Parses a row's <c>ParentId</c> cell, or <see langword="null"/>
    /// for a top-level row -- the same convention the shared projection
    /// applies, so a top-level run is matched by its blank cells.
    /// </summary>
    private static GanttRowId? TryParseParentId(object?[] row, int[] columnMap) =>
        GanttRowId.TryParse(
            ExcelGanttRowProjection.CellText(row, columnMap, "ParentId"),
            out GanttRowId? parentId) && parentId is not null
                ? parentId
                : null;

    /// <summary>
    /// Returns a copy of the scaffold with its <c>ParentId</c> set
    /// to the given parent, leaving every other column untouched.
    /// </summary>
    /// <param name="values">The scaffold row, in schema-column order.</param>
    /// <param name="parentId">The parent to assign.</param>
    /// <returns>The copied scaffold with the <c>ParentId</c> cell populated.</returns>
    private static object?[] WithParentId(
        IReadOnlyList<object?> values,
        GanttRowId parentId)
    {
        object?[] mutable = new object?[values.Count];
        for (int index = 0; index < values.Count; index++)
        {
            mutable[index] = values[index];
        }

        mutable[ExcelGanttRowProjection.SchemaIndexOf("ParentId")] = parentId.Value;
        return mutable;
    }

    /// <summary>
    /// Rebuilds the native outline from the post-insert table (R5.13),
    /// the same re-read/project/apply sequence the parent/child
    /// mutator's <c>Finish</c> uses, so the <c>-</c>/<c>+</c> controls
    /// reflect the new row's parent without waiting for the next
    /// Refresh. A regroup refusal degrades to the pre-D12 behaviour:
    /// the insert stands and the next Refresh rebuilds the outline.
    /// </summary>
    /// <param name="table">The Gantt data table, post-insert.</param>
    /// <param name="columnMap">The schema-column to table-column index map.</param>
    /// <param name="insertedBodyRow">
    /// The one-based body row the new row occupies, or <see langword="null"/> when the insert
    /// appended. Used only to report whether the new row is the one the plan refused.
    /// </param>
    private void RegroupOutline(Excel.ListObject table, int[] columnMap, int? insertedBodyRow)
    {
        if (_outlineGroups is null)
        {
            return;
        }

        List<object?[]> body = ReadBody(table);
        List<GanttEvent>? events = ExcelGanttRowProjection.ProjectRows(
            body,
            columnMap,
            out int? _,
            out int? _);
        if (events is null)
        {
            _trace?.Invoke("gantt.rowinsert regroup rows=" + body.Count + " projected=false");
            return;
        }

        // Forced: Excel copies the outline LEVEL of the row above into a row inserted inside a
        // group, so the new row already reports the child level while sitting outside the group's
        // span. The level-based skip would decline to group it and the row would be left ungrouped
        // in the sheet with no error - measured live 2026-11-09 (groups=0 while the plan had 2 and
        // the new row was inside one).
        OutlineGroupOutcome outcome = _outlineGroups.Apply(events, forceRegroup: true);

        // The plan is ALSO logged, because the writer's counts cannot distinguish the two ways a
        // regroup does nothing: an empty plan (nothing was grouped) and a plan whose groups the
        // writer judged already grouped (nothing was written). `RowsUngrouped` on the OUTCOME
        // counts rows the writer reset, not rows the planner refused, so it reads zero in both.
        OutlineGroupPlan plan = OutlineGroupPlanner.Plan(events);
        bool insertedGrouped = insertedBodyRow is { } row
            && plan.Groups.Any(group => group.ChildRowNumbers.Contains(row));
        bool insertedRefused = insertedBodyRow is { } refused
            && plan.RowsToUngroup.Contains(refused);

        _trace?.Invoke(
            "gantt.rowinsert regroup rows=" + body.Count
            + " events=" + events.Count
            + " groups=" + outcome.GroupsApplied
            + " ungrouped=" + outcome.RowsUngrouped
            + " refusal=" + (outcome.Refusal?.ToString() ?? "none")
            + " parents=" + events.Count(static @event => @event.ParentId is not null)
            + " plan-groups=" + plan.Groups.Count
            + " plan-ungroup=" + plan.RowsToUngroup.Count
            + " inserted=" + (insertedBodyRow?.ToString(CultureInfo.InvariantCulture) ?? "none")
            + " inserted-grouped=" + insertedGrouped
            + " inserted-refused=" + insertedRefused);
    }

    /// <summary>
    /// Returns the table's data body range. Test seam over the COM
    /// parameterised <c>ListObject.DataBodyRange</c> property.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <returns>
    /// The table's data body range, or <see langword="null"/> when the
    /// table has no data rows.
    /// </returns>
    internal virtual Excel.Range? GetDataBodyRange(Excel.ListObject table) => table.DataBodyRange;

    private static void WriteRow(
        Excel.Range rowRange,
        IReadOnlyList<object?> values,
        int[] columnMap)
    {
#pragma warning disable CA1814 // Excel Range.Value2 requires a rectangular SAFEARRAY.
        object[,] matrix = new object[1, columnMap.Length];
        for (int schemaIndex = 0; schemaIndex < values.Count; schemaIndex++)
        {
            int tableIndex = columnMap[schemaIndex];
            if (tableIndex > 0)
            {
                matrix[0, tableIndex - 1] = values[schemaIndex] ?? string.Empty;
            }
        }

        rowRange.Value2 = matrix;
#pragma warning restore CA1814
    }

    internal virtual Excel.Worksheet GetSheetAt(Excel.Sheets sheets, int index) => sheets[index];

    internal virtual Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index) =>
        listObjects[index];

    internal virtual Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index) =>
        columns[index];

    internal virtual Excel.ListRows GetListRows(Excel.ListObject table) => table.ListRows;

    internal virtual int GetListRowCount(Excel.ListRows rows) => rows.Count;

    internal virtual Excel.Range GetTableRange(Excel.ListObject table) => table.Range;

    internal virtual Excel.Range GetRangeRows(Excel.Range range)
    {
        Excel.Range rows = range.Rows;
        return rows;
    }

    internal virtual int GetRangeRowCount(Excel.Range rows)
    {
        int count = rows.Count;
        return count;
    }

    internal virtual Excel.Range GetRangeAt(Excel.Range rows, int index)
    {
        Excel.Range row = rows[index];
        return row;
    }

    internal virtual object? GetRangeValue2(Excel.Range range) => range.Value2;

    internal virtual void ClearRange(Excel.Range range) => range.ClearContents();

    /// <summary>Writes a row height in points. Test seam over <c>Range.RowHeight</c>.</summary>
    /// <param name="range">The row range.</param>
    /// <param name="heightPt">The height to write, in points.</param>
    internal virtual void SetRowHeight(Excel.Range range, double heightPt) =>
        range.RowHeight = heightPt;

    internal virtual Excel.ListRow AddRow(Excel.ListRows rows) => rows.Add(Type.Missing);

    /// <summary>The active cell, or null. Test seam over <c>Application.ActiveCell</c>.</summary>
    /// <param name="application">The application object.</param>
    /// <returns>The active cell.</returns>
    internal virtual Excel.Range? GetActiveCell(Excel.Application application) => application.ActiveCell;

    /// <summary>Whether the table is the active one. Test seam over <c>ListObject.Active</c>.</summary>
    /// <param name="table">The table.</param>
    /// <returns>Whether the table is active.</returns>
    internal virtual bool GetTableActive(Excel.ListObject table) => table.Active;

    /// <summary>A range's first worksheet row. Test seam over <c>Range.Row</c>.</summary>
    /// <param name="range">The range.</param>
    /// <returns>The one-based worksheet row index.</returns>
    internal virtual int GetRangeRow(Excel.Range range) => range.Row;

    internal virtual int GetRowIndex(Excel.ListRow row) => row.Index;

    /// <summary>Reads one body row by its one-based list index. Test seam.</summary>
    /// <param name="rows">The table's list rows.</param>
    /// <param name="index">The one-based body-row index.</param>
    /// <returns>The list row.</returns>
    internal virtual Excel.ListRow GetListRowAt(Excel.ListRows rows, int index) => rows[index];

    internal virtual Excel.Range GetRowRange(Excel.ListRow row) => row.Range;

    internal virtual void DeleteRow(Excel.ListRow row) => row.Delete();
}
