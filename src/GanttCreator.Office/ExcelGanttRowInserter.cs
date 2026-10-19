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
    IOutlineGroupPort? outlineGroups = null) : IGanttRowInserter
{
    private readonly Excel.Application? _application = application as Excel.Application;
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

        // Only the APPEND branch can lose the padding row, because only it depends on
        // the insert below the table having succeeded. A positional insert shifts the
        // sheet by itself, so its padding row is displaced by definition.
        bool paddingRowReserved = true;
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
            int? position = GetInsertionPosition(application!, table, rows);

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

                // The height is already right by Excel's format inheritance, but that is a
                // host behaviour rather than a stated invariant. Writing the token makes it
                // this add-in's decision.
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

            // R5.13: a positional insert lands BELOW the active cell. When that
            // puts the new row inside a group's contiguous child run, make it a
            // child of that group's parent so the run stays contiguous and the
            // outline planner can still represent the group. The append branch
            // (position null) and the header row (position 1) stay top-level.
            //
            // This runs only with an outline writer -- without one there is no
            // grouping to preserve, so the row is the plain top-level append the
            // contract pins. It runs only AFTER the row-creation seams succeeded:
            // a refused positional insert returned above, so a refusal never
            // reads the user's existing rows (ADR-0008 D4). The neighbours are
            // read as one bulk body read, never through GetListRowAt, so the
            // insert's own row-acquisition seam call count is unchanged.
            if (_outlineGroups is not null && position.HasValue && position.Value > 1)
            {
                values = ApplyGroupAwareParent(values, table, position.Value, columnMap);
            }
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
        RegroupOutline(table, columnMap);

        TypeOptionsMaterialiseOutcome typeOptions = _typeOptionsMaterialiser.EnsureCurrent();
        if (!typeOptions.Succeeded)
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
    // XlFormatFrom, and these are the exact values a live probe drove Rows.Insert
    // with successfully.
    private const int _excelShiftDown = -4121;              // xlShiftDown
    private const int _excelFormatFromLeftOrAbove = -4142; // xlFormatFromLeftOrAbove

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

            rows[worksheetRow].Insert(_excelShiftDown, _excelFormatFromLeftOrAbove);
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
    /// Applies group-aware parenting to a positional insert (R5.13).
    /// A positional insert lands BELOW the active cell, so once the
    /// worksheet-row insert has shifted the rows the active cell sits
    /// at the one-based body position <c>position - 1</c> and the row
    /// that ends up below the new row at <c>position + 1</c>. When
    /// the new row lands inside a group's contiguous child run it
    /// becomes a child of that group's parent; when it lands between
    /// a top-level parent and its first child it becomes that
    /// parent's first child; every other position -- including the
    /// header row (position 1) and the append (position
    /// <see langword="null"/>) -- stays top-level, matching the host's
    /// own outline-insert behaviour.
    /// </summary>
    /// <remarks>
    /// The two neighbours are read as ONE bulk body read
    /// (<see cref="GetDataBodyRange"/>/<see cref="GetRangeValue2"/>),
    /// never through <see cref="GetListRowAt"/>: that seam is the
    /// insert's own row-acquisition step, whose call count and order
    /// the contract pins. Reusing <see cref="ReadBody"/> also matches
    /// the projection authority the regroup that follows consumes.
    /// </remarks>
    /// <param name="values">The scaffold row, in schema-column order.</param>
    /// <param name="table">The Gantt data table, post-insert.</param>
    /// <param name="position">The one-based body position the new row lands at.</param>
    /// <param name="columnMap">The schema-column to table-column index map.</param>
    /// <returns>The scaffold, with its <c>ParentId</c> set when the new row joins a group.</returns>
    private IReadOnlyList<object?> ApplyGroupAwareParent(
        IReadOnlyList<object?> values,
        Excel.ListObject table,
        int position,
        int[] columnMap)
    {
        // The active cell is at 0-based body position position - 2 and the
        // row that ends up below the new row at 0-based position. Both must
        // be in range; position 1 (header selected) has no active data cell
        // above the new row, so it stays top-level.
        List<object?[]> body = ReadBody(table);
        if (position < 2 || position >= body.Count)
        {
            return values;
        }

        object?[] activeRow = body[position - 2];
        object?[] belowRow = body[position];

        GanttRowId? activeId = TryParseId(activeRow, columnMap);
        GanttRowId? activeParentId = TryParseParentId(activeRow, columnMap);
        GanttRowId? belowParentId = TryParseParentId(belowRow, columnMap);

        // The active cell is a child of P and the row below is also a
        // child of P: the new row lands inside P's contiguous child run,
        // so it joins that run.
        if (activeParentId is not null && belowParentId == activeParentId)
        {
            return WithParentId(values, activeParentId);
        }

        // The active cell is a top-level parent and the row below is
        // its child: the new row lands between the parent and its first
        // child, so it becomes that parent's first child.
        if (activeParentId is null && activeId is not null && belowParentId == activeId)
        {
            return WithParentId(values, activeId);
        }

        // Every other position -- below a top-level leaf, below a group's
        // last child, or between two different groups -- is top-level,
        // matching the host's own outline-insert behaviour.
        return values;
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
    private void RegroupOutline(Excel.ListObject table, int[] columnMap)
    {
        if (_outlineGroups is null)
        {
            return;
        }

        List<object?[]> body = ReadBody(table);
        List<GanttEvent>? events = ExcelGanttRowProjection.ProjectRows(body, columnMap, out int? _);
        if (events is null)
        {
            return;
        }

        _ = _outlineGroups.Apply(events);
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
