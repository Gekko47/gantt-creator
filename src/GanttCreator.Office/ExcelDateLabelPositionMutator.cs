using System.Runtime.InteropServices;
using GanttCreator.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Applies R5.14 date-label writes to the visible Gantt table: exactly one
/// <c>DateLabelPosition</c> cell on one row (bespoke per entity, never
/// group-applied; a collapsed parent writes its own cell only).
/// </summary>
/// <remarks>
/// <para>
/// The pipeline mirrors the label-position mutator's discipline (R5.5): the
/// protection guard runs first, before any workbook or table access
/// (ADR-0008 D4); the table is located by name; the target worksheet's own
/// protection is queried; the active cell is resolved to a one-based body
/// index through the same active-cell shape the inserter uses; the body is
/// bulk-read; the catalogue decides; and exactly one cell is written back.
/// </para>
/// <para>
/// R5.14 D4: the capability is type-only. A <c>Custom Activity</c> style
/// carries no date mode, so — unlike the description-label mutator — this
/// adapter reads no style registry and has no style-related refusal; the
/// catalogue's <see cref="EntityTypeCatalog.IsDateLabelPositionAllowed"/>
/// decides directly from the row's <c>Type</c>.
/// </para>
/// <para>
/// Every refusal before the cell write leaves the workbook untouched. A null
/// choice clears the cell (blank = Auto per R5.14 D3). Refresh applies the
/// scene effect; this adapter never renders and never touches a shape.
/// </para>
/// </remarks>
public class ExcelDateLabelPositionMutator(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null) : IDateLabelPositionMutator
{
    // object? rather than the interop Application so the AddIn can pass
    // ExcelDnaUtil.Application without naming the interop type (the CS0433
    // duplicate-type hazard documented on IExcelApplicationAdapter) — the
    // same shape ExcelGanttRowInserter uses.
    private readonly Application? _application = application as Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <inheritdoc />
    public DateLabelMutationOutcome SetDateLabelPosition(GanttDateLabelPosition? position)
    {
        // ADR-0008 D4: the shared protection guard is intentionally the first
        // operation in this mutating entry point, before any workbook access.
        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return DateLabelMutationOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? DateLabelMutationRefusalReason.NoActiveWorkbook
                    : DateLabelMutationRefusalReason.TargetProtected);
        }

        Application? application = _application;
        Workbook? workbook = application?.ActiveWorkbook;
        if (workbook is null)
        {
            return DateLabelMutationOutcome.Refused(DateLabelMutationRefusalReason.NoActiveWorkbook);
        }

        Sheets sheets = workbook.Sheets;
        if (!TryFindTable(sheets, out Worksheet? worksheet, out ListObject? table) || worksheet is null || table is null)
        {
            return DateLabelMutationOutcome.Refused(DateLabelMutationRefusalReason.TableMissing);
        }

        ProtectionGuardOutcome targetProtection = _protectionGuard.QueryTarget(worksheet);
        if (targetProtection != ProtectionGuardOutcome.NotProtected)
        {
            return DateLabelMutationOutcome.Refused(
                targetProtection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? DateLabelMutationRefusalReason.NoActiveWorkbook
                    : DateLabelMutationRefusalReason.TargetProtected);
        }

        ListRows rows = GetListRows(table);
        var bodyIndex = ResolveActiveBodyIndex(table, rows);
        return bodyIndex == 0
            ? DateLabelMutationOutcome.Refused(DateLabelMutationRefusalReason.SelectionOutsideTable)
            : !TryReadBody(table, out List<object?[]>? body, out int[]? columnMap)
                || body is null || columnMap is null
                ? DateLabelMutationOutcome.Refused(DateLabelMutationRefusalReason.TableMissing)
                : ApplySetDateLabelPosition(table, body, columnMap, bodyIndex, position);
    }

    /// <summary>
    /// Testable apply path: validates the selected row and writes exactly its
    /// <c>DateLabelPosition</c> cell. Internal so contract tests drive it with a
    /// crafted body.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="body">The raw body matrix, in body order. Not mutated except the target cell.</param>
    /// <param name="columnMap">The schema-column to table-column index map.</param>
    /// <param name="bodyIndex">The one-based body index of the selection.</param>
    /// <param name="position">The choice to write, or null to clear.</param>
    /// <returns>The typed mutation outcome.</returns>
    internal DateLabelMutationOutcome ApplySetDateLabelPosition(
        ListObject table,
        List<object?[]> body,
        int[] columnMap,
        int bodyIndex,
        GanttDateLabelPosition? position)
    {
        var selectedIndex = bodyIndex - 1;
        if (selectedIndex < 0 || selectedIndex >= body.Count)
        {
            return DateLabelMutationOutcome.Refused(
                DateLabelMutationRefusalReason.SelectionOutsideTable);
        }

        if (!GanttRowId.TryParse(
                ExcelGanttRowProjection.CellText(body[selectedIndex], columnMap, "Id"),
                out GanttRowId? _))
        {
            return DateLabelMutationOutcome.Refused(
                DateLabelMutationRefusalReason.SelectionUnknown, bodyIndex);
        }

        var typeText = ExcelGanttRowProjection.CellText(body[selectedIndex], columnMap, "Type");
        if (!EntityTypeCatalog.TryParse(typeText, out GanttEntityType type))
        {
            return DateLabelMutationOutcome.Refused(
                DateLabelMutationRefusalReason.TypeUnknown, bodyIndex);
        }

        // R5.14 D4: the capability is type-only. A Custom Activity style
        // carries no date mode, so it cannot narrow the set the way it
        // narrows LabelPosition — there is no style-registry read here,
        // and therefore no style-related refusal.
        if (position.HasValue
            && !EntityTypeCatalog.IsDateLabelPositionAllowed(type, position.Value))
        {
            return DateLabelMutationOutcome.Refused(
                DateLabelMutationRefusalReason.DateLabelNotAllowedForType, bodyIndex);
        }

        // A null choice clears the cell (blank = Auto per R5.14 D3). Only the
        // selected row's cell is written: bespoke per entity, never the group;
        // a collapsed parent writes its own cell while hidden children keep
        // theirs (owner ruling 2026-10-14).
        return WriteCell(table, columnMap, selectedIndex, "DateLabelPosition", position?.ToString(), bodyIndex);
    }

    /// <summary>
    /// Locates the visible <c>tblGanttData</c> table across every worksheet of
    /// the active workbook. The table is matched by its name (Ordinal-
    /// insensitive), never by position, so a reordered sheet cannot redirect
    /// the write.
    /// </summary>
    internal virtual bool TryFindTable(
        Sheets sheets, out Worksheet? worksheet, out ListObject? table)
    {
        worksheet = null;
        table = null;
        var sheetCount = GetSheetCount(sheets);
        for (var sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
        {
            var candidate = GetSheetAt(sheets, sheetIndex);
            if (candidate is not Worksheet worksheetCandidate)
            {
                continue;
            }

            ListObjects listObjects = worksheetCandidate.ListObjects;
            var tableCount = listObjects.Count;
            for (var tableIndex = 1; tableIndex <= tableCount; tableIndex++)
            {
                ListObject located = GetTableAt(listObjects, tableIndex);
                if (string.Equals(
                    located.Name,
                    GanttTableSchema.TableName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    worksheet = worksheetCandidate;
                    table = located;
                    return true;
                }
            }
        }

        return false;
    }

    /// <summary>
    /// Builds the schema-column to table-column index map from live
    /// header names (Ordinal). Reuses the parent/child mutator's mapping.
    /// </summary>
    internal virtual bool TryBuildColumnMap(ListObject table, out int[]? columnMap)
    {
        columnMap = null;
        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        var map = new int[schema.Count];
        var byName = new Dictionary<string, int>(StringComparer.Ordinal);
        ListColumns columns = table.ListColumns;
        var columnCount = columns.Count;
        for (var index = 1; index <= columnCount; index++)
        {
            ListColumn column = GetColumnAt(columns, index);
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

    /// <summary>
    /// Bulk-reads the table body into raw cell rows, in body order.
    /// </summary>
    internal virtual bool TryReadBody(
        ListObject table,
        out List<object?[]>? body,
        out int[]? columnMap)
    {
        body = null;
        if (!TryBuildColumnMap(table, out columnMap) || columnMap is null)
        {
            return false;
        }

        Excel.Range? dataBody = GetDataBodyRange(table);
        if (dataBody is null)
        {
            body = [];
            return true;
        }

        body = ExcelValue2Matrix.ReadRows(GetRangeValue2(dataBody));
        return true;
    }

    /// <summary>
    /// Resolves the active cell to a one-based body index through the same
    /// active-cell shape the row inserter uses.
    /// </summary>
    internal virtual int ResolveActiveBodyIndex(ListObject table, ListRows rows)
    {
        Application? application = _application;
        if (application is null)
        {
            return 0;
        }

        Excel.Range? activeCell = GetActiveCell(application);
        if (!GetTableActive(table) || activeCell is null)
        {
            return 0;
        }

        Excel.Range tableRange = GetTableRange(table);
        var tableFirstRow = GetRangeRow(tableRange);
        var activeRow = GetRangeRow(activeCell);
        var bodyOffset = activeRow - tableFirstRow;
        var rowCount = GetListRowCount(rows);

        return bodyOffset < 1 || bodyOffset > rowCount ? 0 : bodyOffset;
    }

    /// <summary>
    /// Writes exactly one cell. The parent/child mutator's single-cell path,
    /// narrowed to the date-label column: the row is addressed through
    /// <c>ListRows</c> and the column through <c>ListColumns</c>, so no
    /// <c>Range.Rows</c> cast is needed.
    /// </summary>
    private DateLabelMutationOutcome WriteCell(
        ListObject table,
        int[] columnMap,
        int rowIndex,
        string columnName,
        object? value,
        int bodyIndex)
    {
        try
        {
            var tableIndex = columnMap[ExcelGanttRowProjection.SchemaIndexOf(columnName)];
            ListRow row = GetListRowAt(GetListRows(table), rowIndex + 1);
            Excel.Range cell = GetCellAt(GetListRowRange(row), tableIndex);
            SetRangeValues(cell, value);
        }
        catch (COMException)
        {
            return DateLabelMutationOutcome.Refused(
                DateLabelMutationRefusalReason.WriteFailed, bodyIndex);
        }

        return DateLabelMutationOutcome.Ok(1, bodyIndex);
    }

    /// <summary>Returns the active cell. Test seam over <c>Application.ActiveCell</c>.</summary>
    internal virtual Excel.Range? GetActiveCell(Application application) => application.ActiveCell;

    /// <summary>Test seam over the COM <c>ListObject.Active</c> property.</summary>
    internal virtual bool GetTableActive(ListObject table) => table.Active;

    /// <summary>Test seam over the COM <c>ListObject.Range</c> property.</summary>
    internal virtual Excel.Range GetTableRange(ListObject table) => table.Range;

    /// <summary>Test seam over the COM <c>Range.Row</c> property.</summary>
    internal virtual int GetRangeRow(Excel.Range range) => range.Row;

    /// <summary>Test seam over the COM <c>ListObject.DataBodyRange</c> property.</summary>
    internal virtual Excel.Range? GetDataBodyRange(ListObject table) => table.DataBodyRange;

    /// <summary>Test seam over the COM <c>Range.Value2</c> property.</summary>
    internal virtual object? GetRangeValue2(Excel.Range range) => range.Value2;

    /// <summary>
    /// Returns the body row at the one-based index. Test seam over the COM
    /// <c>ListRows</c> indexer.
    /// </summary>
    internal virtual ListRow GetListRowAt(ListRows rows, int index) => rows[index];

    /// <summary>Test seam over the COM <c>ListRow.Range</c> property.</summary>
    internal virtual Excel.Range GetListRowRange(ListRow row) => row.Range;

    /// <summary>
    /// Returns the cell at the one-based column index. Test seam over the COM
    /// parameterised <c>Range.Cells</c> property.
    /// </summary>
    internal virtual Excel.Range GetCellAt(Excel.Range row, int index) => row.Cells[index];

    /// <summary>Test seam over the COM <c>Range.Value2</c> setter.</summary>
    internal virtual void SetRangeValues(Excel.Range range, object? value) => range.Value2 = value;

    /// <summary>Test seam over the COM <c>Sheets.Count</c> property.</summary>
    internal virtual int GetSheetCount(Sheets sheets) => sheets.Count;

    /// <summary>Test seam over the COM parameterised <c>Sheets.Item</c> property.</summary>
    internal virtual object GetSheetAt(Sheets sheets, int index) => sheets[index];

    /// <summary>Test seam over the COM parameterised <c>ListObjects.Item</c> property.</summary>
    internal virtual ListObject GetTableAt(ListObjects listObjects, int index) =>
        listObjects[index];

    /// <summary>Test seam over the COM parameterised <c>ListColumns.Item</c> property.</summary>
    internal virtual ListColumn GetColumnAt(ListColumns columns, int index) =>
        columns[index];

    /// <summary>Test seam over the COM <c>ListObject.ListRows</c> property.</summary>
    internal virtual ListRows GetListRows(ListObject table) => table.ListRows;

    /// <summary>Test seam over the COM <c>ListRows.Count</c> property.</summary>
    internal virtual int GetListRowCount(ListRows rows) => rows.Count;
}