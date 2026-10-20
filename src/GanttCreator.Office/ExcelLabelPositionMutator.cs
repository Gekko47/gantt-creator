using System.Runtime.InteropServices;
using GanttCreator.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Applies R5.5 label-position writes to the visible Gantt table: exactly one
/// <c>LabelPosition</c> cell on one row (bespoke per entity, never
/// group-applied; a collapsed parent writes its own cell only).
/// </summary>
/// <remarks>
/// <para>
/// The pipeline mirrors the parent/child mutator's discipline (R5.3): the
/// protection guard runs first, before any workbook or table access
/// (ADR-0008 D4); the table is located by name; the target worksheet's own
/// protection is queried; the active cell is resolved to a one-based body
/// index through the same active-cell shape the inserter uses; the body is
/// bulk-read; the catalogue (plus the style registry for Custom Activity)
/// decides; and exactly one cell is written back.
/// </para>
/// <para>
/// Every refusal before the cell write leaves the workbook untouched. A null
/// position clears the cell (blank = Default per §22). Refresh applies the
/// scene effect; this adapter never renders and never touches a shape.
/// </para>
/// </remarks>
public class ExcelLabelPositionMutator(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null,
    IConfigCatalogueReader? catalogueReader = null) : ILabelPositionMutator
{
    // object? rather than the interop Application so the AddIn can pass
    // ExcelDnaUtil.Application without naming the interop type (the CS0433
    // duplicate-type hazard documented on IExcelApplicationAdapter) — the
    // same shape ExcelGanttRowInserter uses.
    private readonly Application? _application = application as Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);
    private readonly IConfigCatalogueReader _catalogueReader =
        catalogueReader ?? new ExcelConfigCatalogueReader(application);

    /// <inheritdoc />
    public LabelPositionMutationOutcome SetLabelPosition(GanttLabelPosition? position)
    {
        // ADR-0008 D4: the shared protection guard is intentionally the first
        // operation in this mutating entry point, before any workbook access.
        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return LabelPositionMutationOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? LabelPositionMutationRefusalReason.NoActiveWorkbook
                    : LabelPositionMutationRefusalReason.TargetProtected);
        }

        Application? application = _application;
        Workbook? workbook = application?.ActiveWorkbook;
        if (workbook is null)
        {
            return LabelPositionMutationOutcome.Refused(LabelPositionMutationRefusalReason.NoActiveWorkbook);
        }

        Sheets sheets = workbook.Sheets;
        if (!TryFindTable(sheets, out Worksheet? worksheet, out ListObject? table) || worksheet is null || table is null)
        {
            return LabelPositionMutationOutcome.Refused(LabelPositionMutationRefusalReason.TableMissing);
        }

        ProtectionGuardOutcome targetProtection = _protectionGuard.QueryTarget(worksheet);
        if (targetProtection != ProtectionGuardOutcome.NotProtected)
        {
            return LabelPositionMutationOutcome.Refused(
                targetProtection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? LabelPositionMutationRefusalReason.NoActiveWorkbook
                    : LabelPositionMutationRefusalReason.TargetProtected);
        }

        ListRows rows = GetListRows(table);
        int bodyIndex = ResolveActiveBodyIndex(table, rows);
        return bodyIndex == 0
            ? LabelPositionMutationOutcome.Refused(LabelPositionMutationRefusalReason.SelectionOutsideTable)
            : !TryReadBody(table, out List<object?[]>? body, out int[]? columnMap)
                || body is null || columnMap is null
                ? LabelPositionMutationOutcome.Refused(LabelPositionMutationRefusalReason.TableMissing)
                : ApplySetLabelPosition(table, body, columnMap, bodyIndex, position);
    }

    /// <summary>
    /// Testable apply path: validates the selected row and writes exactly its
    /// <c>LabelPosition</c> cell. Internal so contract tests drive it with a
    /// crafted body.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="body">The raw body matrix, in body order. Not mutated except the target cell.</param>
    /// <param name="columnMap">The schema-column to table-column index map.</param>
    /// <param name="bodyIndex">The one-based body index of the selection.</param>
    /// <param name="position">The position to write, or null to clear.</param>
    /// <returns>The typed mutation outcome.</returns>
    internal LabelPositionMutationOutcome ApplySetLabelPosition(
        ListObject table,
        List<object?[]> body,
        int[] columnMap,
        int bodyIndex,
        GanttLabelPosition? position)
    {
        int selectedIndex = bodyIndex - 1;
        if (selectedIndex < 0 || selectedIndex >= body.Count)
        {
            return LabelPositionMutationOutcome.Refused(
                LabelPositionMutationRefusalReason.SelectionOutsideTable);
        }

        if (!GanttRowId.TryParse(
                ExcelGanttRowProjection.CellText(body[selectedIndex], columnMap, "Id"),
                out GanttRowId? _))
        {
            return LabelPositionMutationOutcome.Refused(
                LabelPositionMutationRefusalReason.SelectionUnknown, bodyIndex);
        }

        string typeText = ExcelGanttRowProjection.CellText(body[selectedIndex], columnMap, "Type");
        if (!EntityTypeCatalog.TryParse(typeText, out GanttEntityType type))
        {
            return LabelPositionMutationOutcome.Refused(
                LabelPositionMutationRefusalReason.TypeUnknown, bodyIndex);
        }

        GanttStyleRegistry? registry = null;
        string? styleKey = null;

        // R5.6a: the catalogue's RequiresStyleKey flag is the single authority
        // for which types resolve from the style registry (Custom Activity and
        // Custom Milestone), so a Custom Milestone row validates against its
        // row style exactly like a Custom Activity does.
        EntityTypeDefinition? definition = EntityTypeCatalog.GetDefinition(type);
        if (definition is { RequiresStyleKey: true })
        {
            styleKey = ExcelGanttRowProjection.CellText(body[selectedIndex], columnMap, "StyleKey");
            ConfigReadOutcome catalogue = ReadCatalogue();
            if (!catalogue.Succeeded)
            {
                return LabelPositionMutationOutcome.Refused(
                    LabelPositionMutationRefusalReason.TableInvalid, bodyIndex);
            }

            registry = catalogue.Styles;
        }

        // R5.6d: the vocabulary is the row's style when the registry knows it
        // and the type's CODE-OWNED default style otherwise - the same preset
        // the scene falls back to when it renders the row. A row whose style
        // has not been chosen yet, or whose style was deleted, still renders
        // with that default, so it still offers the labels that default
        // permits; refusing it left the description-label dropdown disabled
        // on a row that was perfectly renderable. The vocabulary is never
        // empty for a known type - the registry rejects a style that permits
        // no position, and every catalogue type defines at least one - so a
        // position outside it is the ordinary not-permitted refusal.
        IReadOnlyList<GanttLabelPosition> allowed =
            EntityTypeCatalog.GetAllowedLabelPositions(type, styleKey, registry);

        if (position.HasValue && !allowed.Contains(position.Value))
        {
            return LabelPositionMutationOutcome.Refused(
                LabelPositionMutationRefusalReason.PositionNotAllowedForType, bodyIndex);
        }

        // A null position clears the cell (blank = Default per §22). Only the
        // selected row's cell is written: bespoke per entity, never the group;
        // a collapsed parent writes its own cell while hidden children keep
        // theirs (owner ruling 2026-10-14).
        object? value = position.HasValue ? position.Value.ToString() : null;
        return WriteCell(table, columnMap, selectedIndex, "LabelPosition", value, bodyIndex);
    }

    /// <summary>Reads the configuration catalogue (styles for Custom Activity).</summary>
    internal virtual ConfigReadOutcome ReadCatalogue() => _catalogueReader.Read();

    private bool TryFindTable(
        Sheets sheets,
        out Worksheet? worksheet,
        out ListObject? table)
    {
        worksheet = null;
        table = null;

        int sheetCount = GetSheetCount(sheets);
        for (int sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
        {
            if (GetSheetAt(sheets, sheetIndex) is not Worksheet candidate)
            {
                continue;
            }

            ListObjects listObjects = candidate.ListObjects;
            int tableCount = listObjects.Count;
            for (int tableIndex = 1; tableIndex <= tableCount; tableIndex++)
            {
                ListObject located = GetTableAt(listObjects, tableIndex);
                if (string.Equals(
                    located.Name,
                    GanttTableSchema.TableName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    worksheet = candidate;
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
        int[] map = new int[schema.Count];
        Dictionary<string, int> byName = new Dictionary<string, int>(StringComparer.Ordinal);
        ListColumns columns = table.ListColumns;
        int columnCount = columns.Count;
        for (int index = 1; index <= columnCount; index++)
        {
            ListColumn column = GetColumnAt(columns, index);
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
        int tableFirstRow = GetRangeRow(tableRange);
        int activeRow = GetRangeRow(activeCell);
        int bodyOffset = activeRow - tableFirstRow;
        int rowCount = GetListRowCount(rows);

        return bodyOffset < 1 || bodyOffset > rowCount ? 0 : bodyOffset;
    }

    /// <summary>
    /// Writes exactly one cell. The parent/child mutator's single-cell path,
    /// narrowed to the label column: the row is addressed through
    /// <c>ListRows</c> and the column through <c>ListColumns</c>, so no
    /// <c>Range.Rows</c> cast is needed.
    /// </summary>
    private LabelPositionMutationOutcome WriteCell(
        ListObject table,
        int[] columnMap,
        int rowIndex,
        string columnName,
        object? value,
        int bodyIndex)
    {
        try
        {
            int tableIndex = columnMap[ExcelGanttRowProjection.SchemaIndexOf(columnName)];
            ListRow row = GetListRowAt(GetListRows(table), rowIndex + 1);
            Excel.Range cell = GetCellAt(GetListRowRange(row), tableIndex);
            SetRangeValues(cell, value);
        }
        catch (COMException)
        {
            return LabelPositionMutationOutcome.Refused(
                LabelPositionMutationRefusalReason.WriteFailed, bodyIndex);
        }

        return LabelPositionMutationOutcome.Ok(1, bodyIndex);
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
