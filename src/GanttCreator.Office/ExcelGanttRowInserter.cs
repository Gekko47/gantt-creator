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
    ITypeOptionsMaterialiser? typeOptionsMaterialiser = null) : IGanttRowInserter
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);
    private readonly ITypeOptionsMaterialiser _typeOptionsMaterialiser =
        typeOptionsMaterialiser ?? new ExcelTypeOptionsMaterialiser(application);

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
            || !TryBuildColumnMap(table, out var columnMap) || columnMap is null)
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
        var paddingRowReserved = true;
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
            var position = GetInsertionPosition(application!, table, rows);

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
            var insideTable = position is not null;

            // Inside the table for a positional insert, so the ListObject absorbs the
            // row. One past the table's last row for an append, so the padding row is
            // displaced rather than consumed -- Q4 measured that a row inserted below
            // the table does NOT auto-expand it, which is why the append branch still
            // follows its insert with ListRows.Add().
            paddingRowReserved = insideTable
                ? InsertWorksheetRow(
                    table,
                    GetRangeRow(GetTableRange(table)) + position!.Value)
                : InsertWorksheetRowBelowTable(table);

            newRow = insideTable ? GetListRowAt(rows, position!.Value) : AddRow(rows);

            rowRange = GetRowRange(newRow);

            // The height is already right by Excel's format inheritance, but that is a
            // host behaviour rather than a stated invariant. Writing the token makes it
            // this add-in's decision.
            SetRowHeight(rowRange, GanttCatalogues.MetricDefault("GanttRowHeightPt"));
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

        return GanttRowInsertOutcome.Ok(newRow is null ? 1 : GetRowIndex(newRow), paddingRowReserved);
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
        var tableFirstRow = GetRangeRow(tableRange);
        var activeRow = GetRangeRow(activeCell);
        var bodyIndex = activeRow - tableFirstRow;
        var rowCount = GetListRowCount(rows);

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
        var lastRow = GetRangeRow(tableRange) + (tableRange.Rows?.Count ?? 0) - 1;
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
        var rowCount = GetRangeRowCount(tableRows);
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
        var sheetCount = sheets.Count;
        for (var sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
        {
            if (GetSheetAt(sheets, sheetIndex) is not Excel.Worksheet candidate)
            {
                continue;
            }

            Excel.ListObjects listObjects = candidate.ListObjects;
            var tableCount = listObjects.Count;
            for (var tableIndex = 1; tableIndex <= tableCount; tableIndex++)
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
        var map = new int[schema.Count];
        var byName = new Dictionary<string, int>(StringComparer.Ordinal);
        Excel.ListColumns columns = table.ListColumns;
        var columnCount = columns.Count;
        for (var index = 1; index <= columnCount; index++)
        {
            Excel.ListColumn column = GetColumnAt(columns, index);
            byName[column.Name] = index;
        }

        for (var schemaIndex = 0; schemaIndex < schema.Count; schemaIndex++)
        {
            GanttTableColumn expected = schema[schemaIndex];
            if (!byName.TryGetValue(expected.Name, out var tableIndex))
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

    private static void WriteRow(
        Excel.Range rowRange,
        IReadOnlyList<object?> values,
        int[] columnMap)
    {
#pragma warning disable CA1814 // Excel Range.Value2 requires a rectangular SAFEARRAY.
        var matrix = new object[1, columnMap.Length];
        for (var schemaIndex = 0; schemaIndex < values.Count; schemaIndex++)
        {
            var tableIndex = columnMap[schemaIndex];
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
        var count = rows.Count;
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
