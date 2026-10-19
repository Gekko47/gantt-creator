using System.Runtime.InteropServices;
using GanttCreator.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Applies R5.6a style-key writes to the visible Gantt table: exactly one
/// <c>StyleKey</c> cell on one row (bespoke per entity, never group-applied).
/// </summary>
/// <remarks>
/// <para>
/// The pipeline mirrors the label-position mutator's discipline (R5.5): the
/// protection guard runs first, before any workbook or table access
/// (ADR-0008 D4); the table is located by name; the target worksheet's own
/// protection is queried; the active cell is resolved to a one-based body
/// index; the body is bulk-read; the catalogue decides (a style-key type via
/// <c>RequiresStyleKey</c>, a known key via the style registry); and exactly
/// one cell is written back.
/// </para>
/// <para>
/// Every refusal before the cell write leaves the workbook untouched. The
/// hidden <c>StyleKey</c> column is written through the table's list model by
/// header name, so its column visibility never affects the write. Refresh
/// applies the scene effect; this adapter never renders and never touches a
/// shape.
/// </para>
/// </remarks>
public class ExcelStyleKeyMutator(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null,
    IConfigCatalogueReader? catalogueReader = null) : IStyleKeyMutator
{
    // object? rather than the interop Application so the AddIn can pass
    // ExcelDnaUtil.Application without naming the interop type (the CS0433
    // duplicate-type hazard documented on IExcelApplicationAdapter) — the
    // same shape ExcelLabelPositionMutator uses.
    private readonly Application? _application = application as Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);
    private readonly IConfigCatalogueReader _catalogueReader =
        catalogueReader ?? new ExcelConfigCatalogueReader(application);

    /// <inheritdoc />
    public StyleKeyMutationOutcome SetStyleKey(string styleKey)
    {
        // ADR-0008 D4: the shared protection guard is intentionally the first
        // operation in this mutating entry point, before any workbook access.
        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return StyleKeyMutationOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? StyleKeyMutationRefusalReason.NoActiveWorkbook
                    : StyleKeyMutationRefusalReason.TargetProtected);
        }

        Application? application = _application;
        Workbook? workbook = application?.ActiveWorkbook;
        if (workbook is null)
        {
            return StyleKeyMutationOutcome.Refused(StyleKeyMutationRefusalReason.NoActiveWorkbook);
        }

        Sheets sheets = workbook.Sheets;
        if (!TryFindTable(sheets, out Worksheet? worksheet, out ListObject? table) || worksheet is null || table is null)
        {
            return StyleKeyMutationOutcome.Refused(StyleKeyMutationRefusalReason.TableMissing);
        }

        ProtectionGuardOutcome targetProtection = _protectionGuard.QueryTarget(worksheet);
        if (targetProtection != ProtectionGuardOutcome.NotProtected)
        {
            return StyleKeyMutationOutcome.Refused(
                targetProtection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? StyleKeyMutationRefusalReason.NoActiveWorkbook
                    : StyleKeyMutationRefusalReason.TargetProtected);
        }

        ListRows rows = GetListRows(table);
        int bodyIndex = ResolveActiveBodyIndex(table, rows);
        return bodyIndex == 0
            ? StyleKeyMutationOutcome.Refused(StyleKeyMutationRefusalReason.SelectionOutsideTable)
            : !TryReadBody(table, out List<object?[]>? body, out int[]? columnMap)
                || body is null || columnMap is null
                ? StyleKeyMutationOutcome.Refused(StyleKeyMutationRefusalReason.TableMissing)
                : ApplySetStyleKey(table, body, columnMap, bodyIndex, styleKey);
    }

    /// <summary>
    /// Testable apply path: validates the selected row and writes exactly its
    /// <c>StyleKey</c> cell. Internal so contract tests drive it with a
    /// crafted body.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="body">The raw body matrix, in body order. Not mutated except the target cell.</param>
    /// <param name="columnMap">The schema-column to table-column index map.</param>
    /// <param name="bodyIndex">The one-based body index of the selection.</param>
    /// <param name="styleKey">The exact style key to write.</param>
    /// <returns>The typed mutation outcome.</returns>
    internal StyleKeyMutationOutcome ApplySetStyleKey(
        ListObject table,
        List<object?[]> body,
        int[] columnMap,
        int bodyIndex,
        string styleKey)
    {
        string key = styleKey?.Trim() ?? string.Empty;

        int selectedIndex = bodyIndex - 1;
        if (selectedIndex < 0 || selectedIndex >= body.Count)
        {
            return StyleKeyMutationOutcome.Refused(
                StyleKeyMutationRefusalReason.SelectionOutsideTable);
        }

        if (!GanttRowId.TryParse(
                ExcelGanttRowProjection.CellText(body[selectedIndex], columnMap, "Id"),
                out GanttRowId? _))
        {
            return StyleKeyMutationOutcome.Refused(
                StyleKeyMutationRefusalReason.SelectionUnknown, bodyIndex);
        }

        string typeText = ExcelGanttRowProjection.CellText(body[selectedIndex], columnMap, "Type");
        if (!EntityTypeCatalog.TryParse(typeText, out GanttEntityType type))
        {
            return StyleKeyMutationOutcome.Refused(
                StyleKeyMutationRefusalReason.TypeUnknown, bodyIndex);
        }

        // The catalogue's RequiresStyleKey flag is the single authority for
        // which types carry a style: Custom Activity and Custom Milestone
        // (R5.6a); every other type is refused before any write.
        if (EntityTypeCatalog.GetDefinition(type) is not { RequiresStyleKey: true })
        {
            return StyleKeyMutationOutcome.Refused(
                StyleKeyMutationRefusalReason.TypeNotCustomStyled, bodyIndex);
        }

        // An explicit clear (the dropdown's <c>None</c> item) blanks the cell
        // on a style-key-type row; an unknown non-empty key is still refused.
        if (key.Length == 0)
        {
            return WriteCell(table, columnMap, selectedIndex, "StyleKey", string.Empty, bodyIndex);
        }

        ConfigReadOutcome catalogue = ReadCatalogue();
        return !catalogue.Succeeded
            ? StyleKeyMutationOutcome.Refused(
                StyleKeyMutationRefusalReason.TableInvalid, bodyIndex)
            : catalogue.Styles is null || !catalogue.Styles.TryGet(key, out _)
                ? StyleKeyMutationOutcome.Refused(StyleKeyMutationRefusalReason.StyleUnknown, bodyIndex)
                : WriteCell(table, columnMap, selectedIndex, "StyleKey", key, bodyIndex);
    }

    /// <summary>Reads the configuration catalogue (the style registry).</summary>
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
    /// header names (Ordinal). Reuses the label mutator's mapping.
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
    /// active-cell shape the label mutator uses.
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
    /// Writes exactly one cell. The row is addressed through
    /// <c>ListRows</c> and the column through <c>ListColumns</c> (resolved by
    /// header name), so the hidden StyleKey column is written exactly like a
    /// visible one.
    /// </summary>
    private StyleKeyMutationOutcome WriteCell(
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
            return StyleKeyMutationOutcome.Refused(
                StyleKeyMutationRefusalReason.WriteFailed, bodyIndex);
        }

        return StyleKeyMutationOutcome.Ok(1, bodyIndex);
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
    internal virtual ListColumn GetColumnAt(ListColumns columns, int index) => columns[index];

    /// <summary>Test seam over the COM <c>ListObject.ListRows</c> property.</summary>
    internal virtual ListRows GetListRows(ListObject table) => table.ListRows;

    /// <summary>Test seam over the COM <c>ListRows.Count</c> property.</summary>
    internal virtual int GetListRowCount(ListRows rows) => rows.Count;
}
