using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Excel adapter for <see cref="ITypeChangeNormaliser"/>. Reads the visible
/// Gantt table, corrects type-dependent engine-hidden columns so a <c>Type</c>
/// change does not leave stale values the validator would reject, and writes
/// the corrections back.
/// </summary>
public class ExcelTypeChangeNormaliser(object? application, IWorksheetProtectionGuard? protectionGuard = null)
    : ITypeChangeNormaliser
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IWorksheetProtectionGuard _protectionGuard = protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <inheritdoc />
    public GanttTypeChangeNormaliseOutcome Normalise()
    {
        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (workbook is null ||
            !TryFindTable(workbook.Sheets, out Excel.Worksheet? worksheet, out Excel.ListObject? table) ||
            worksheet is null || table is null)
        {
            return GanttTypeChangeNormaliseOutcome.Ok(0);
        }

        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(worksheet);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return GanttTypeChangeNormaliseOutcome.Ok(0);
        }

        Excel.Range? body = GetTableBody(table);
        if (body is null)
        {
            return GanttTypeChangeNormaliseOutcome.Ok(0);
        }

        // ONE walk of the columns, not one walk per name. Ten names resolved by scanning
        // the collection each time is 10 x 26 = 260 COM `Name` reads per refresh; the
        // owner's 2026-11-20 trace measured `types` at 7.2-26.7 ms (median 12.5) for ten
        // rows of data. The same defect was already fixed in StagingShapeAdapter and in
        // ExcelColumnPresentationRestorer; this was the third instance.
        Dictionary<string, int> columnIndexes = BuildColumnIndexMap(table);
        int typeIndex = IndexOf(columnIndexes, "Type");
        int styleKeyIndex = IndexOf(columnIndexes, "StyleKey");
        int labelPositionIndex = IndexOf(columnIndexes, "LabelPosition");
        int dateLabelPositionIndex = IndexOf(columnIndexes, "DateLabelPosition");
        int fillColourIndex = IndexOf(columnIndexes, "FillColour");
        int strokeColourIndex = IndexOf(columnIndexes, "StrokeColour");
        int parentIdIndex = IndexOf(columnIndexes, "ParentId");
        int laneIdIndex = IndexOf(columnIndexes, "LaneId");
        int stackIndexIndex = IndexOf(columnIndexes, "StackIndex");
        int siblingOrderIndex = IndexOf(columnIndexes, "SiblingOrder");

        if (typeIndex < 0)
        {
            return GanttTypeChangeNormaliseOutcome.Ok(0);
        }

        List<object?[]> rows = ExcelValue2Matrix.ReadRows(GetBodyValues(body));
        int rowCount = rows.Count;
        int written = 0;

        for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            object?[] row = rows[rowIndex];
            string? typeText = ToText(row.ElementAtOrDefault(typeIndex))?.Trim();
            if (string.IsNullOrEmpty(typeText) ||
                !GanttEntityCatalog.TryParse(typeText, out GanttEntityType type))
            {
                continue;
            }

            EntityTypeDefinition? definition = GanttEntityCatalog.GetDefinition(type);
            if (definition is null)
            {
                continue;
            }

            int excelRow = rowIndex + 1;

            NormaliseStyleKey(body, excelRow, row, definition, styleKeyIndex, ref written);
            NormaliseLabelPosition(body, excelRow, row, definition, labelPositionIndex, ref written);
            NormaliseDateLabelPosition(body, excelRow, row, type, dateLabelPositionIndex, ref written);
            NormaliseFillColour(body, excelRow, row, definition, fillColourIndex, ref written);
            NormaliseStrokeColour(body, excelRow, row, definition, strokeColourIndex, ref written);
            NormaliseParentId(body, excelRow, row, type, parentIdIndex, ref written);
            NormaliseLaneFields(body, excelRow, row, definition, laneIdIndex, stackIndexIndex, ref written);
            NormaliseSiblingOrder(body, excelRow, row, type, siblingOrderIndex, ref written);
        }

        return GanttTypeChangeNormaliseOutcome.Ok(written);
    }

    private void NormaliseStyleKey(Excel.Range body, int excelRow, object?[] row,
        EntityTypeDefinition definition, int styleKeyIndex, ref int written)
    {
        if (!definition.RequiresStyleKey || styleKeyIndex < 0)
        {
            return;
        }

        string current = ToText(row.ElementAtOrDefault(styleKeyIndex))?.Trim() ?? string.Empty;
        if (current.Length == 0 && definition.DefaultStyleKey.Length > 0)
        {
            GetCell(body, excelRow, styleKeyIndex + 1).Value2 = definition.DefaultStyleKey;
            written++;
        }
    }

    private void NormaliseLabelPosition(Excel.Range body, int excelRow, object?[] row,
        EntityTypeDefinition definition, int labelPositionIndex, ref int written)
    {
        if (labelPositionIndex < 0)
        {
            return;
        }

        string current = ToText(row.ElementAtOrDefault(labelPositionIndex))?.Trim() ?? string.Empty;
        if (current.Length > 0 &&
            GanttLabelCatalog.TryParse(current, out GanttLabelPosition position) &&
            !definition.AllowedLabelPositions.Contains(position))
        {
            GetCell(body, excelRow, labelPositionIndex + 1).Value2 = string.Empty;
            written++;
        }
    }

    private void NormaliseDateLabelPosition(Excel.Range body, int excelRow, object?[] row,
        GanttEntityType type, int dateLabelPositionIndex, ref int written)
    {
        if (dateLabelPositionIndex < 0)
        {
            return;
        }

        string current = ToText(row.ElementAtOrDefault(dateLabelPositionIndex))?.Trim() ?? string.Empty;
        if (current.Length > 0 &&
            GanttDateLabelCatalog.TryParse(current, out GanttDateLabelPosition position) &&
            !GanttEntityCatalog.IsDateLabelPositionAllowed(type, position))
        {
            GetCell(body, excelRow, dateLabelPositionIndex + 1).Value2 = string.Empty;
            written++;
        }
    }

    private void NormaliseFillColour(Excel.Range body, int excelRow, object?[] row,
        EntityTypeDefinition definition, int fillColourIndex, ref int written)
    {
        if (fillColourIndex < 0 ||
            (definition.ColourCapability & EntityColourCapability.Fill) != 0)
        {
            return;
        }

        string current = ToText(row.ElementAtOrDefault(fillColourIndex))?.Trim() ?? string.Empty;
        if (current.Length > 0)
        {
            GetCell(body, excelRow, fillColourIndex + 1).Value2 = string.Empty;
            written++;
        }
    }

    private void NormaliseStrokeColour(Excel.Range body, int excelRow, object?[] row,
        EntityTypeDefinition definition, int strokeColourIndex, ref int written)
    {
        if (strokeColourIndex < 0 ||
            (definition.ColourCapability & EntityColourCapability.Stroke) != 0)
        {
            return;
        }

        string current = ToText(row.ElementAtOrDefault(strokeColourIndex))?.Trim() ?? string.Empty;
        if (current.Length > 0)
        {
            GetCell(body, excelRow, strokeColourIndex + 1).Value2 = string.Empty;
            written++;
        }
    }

    private void NormaliseParentId(Excel.Range body, int excelRow, object?[] row,
        GanttEntityType type, int parentIdIndex, ref int written)
    {
        if (parentIdIndex < 0 || EntityHierarchyCatalog.MayBeChild(type))
        {
            return;
        }

        string current = ToText(row.ElementAtOrDefault(parentIdIndex))?.Trim() ?? string.Empty;
        if (current.Length > 0)
        {
            GetCell(body, excelRow, parentIdIndex + 1).Value2 = string.Empty;
            written++;
        }
    }

    private void NormaliseLaneFields(Excel.Range body, int excelRow, object?[] row,
        EntityTypeDefinition definition,
        int laneIdIndex, int stackIndexIndex, ref int written)
    {
        bool usesLane = definition.Kind is EntityKind.Span or EntityKind.Milestone;

        if (laneIdIndex >= 0 && !usesLane)
        {
            string current = ToText(row.ElementAtOrDefault(laneIdIndex))?.Trim() ?? string.Empty;
            if (current.Length > 0)
            {
                GetCell(body, excelRow, laneIdIndex + 1).Value2 = string.Empty;
                written++;
            }
        }

        if (stackIndexIndex >= 0 && !usesLane)
        {
            string current = ToText(row.ElementAtOrDefault(stackIndexIndex))?.Trim() ?? string.Empty;
            if (current.Length > 0)
            {
                GetCell(body, excelRow, stackIndexIndex + 1).Value2 = string.Empty;
                written++;
            }
        }
    }

    private void NormaliseSiblingOrder(Excel.Range body, int excelRow, object?[] row,
        GanttEntityType type, int siblingOrderIndex, ref int written)
    {
        if (siblingOrderIndex < 0 || EntityHierarchyCatalog.MayBeChild(type))
        {
            return;
        }

        string current = ToText(row.ElementAtOrDefault(siblingOrderIndex))?.Trim() ?? string.Empty;
        if (current.Length > 0)
        {
            GetCell(body, excelRow, siblingOrderIndex + 1).Value2 = string.Empty;
            written++;
        }
    }

    /// <summary>
    /// Builds the live column name to zero-based index map with one <c>Name</c> read per
    /// column.
    /// </summary>
    /// <param name="table">The Gantt table.</param>
    /// <returns>Each column name and its zero-based offset.</returns>
    /// <remarks>
    /// <para>
    /// <b>Zero-based</b> because every use of an index here is as an offset into the
    /// body's <c>Value2</c> matrix, which is zero-based, while a <c>ListColumns</c> index
    /// is one-based. The scan this replaces returned <c>index - 1</c> for the same
    /// reason, and <see cref="IndexOf"/> preserves it exactly.
    /// </para>
    /// <para>
    /// <b>First match wins.</b> The scan it replaces returned the first column carrying
    /// the name, so <c>TryAdd</c> is the faithful translation rather than an overwrite.
    /// A duplicated header is not a state the schema produces, but the behaviour is
    /// unchanged if one ever appears.
    /// </para>
    /// </remarks>
    private Dictionary<string, int> BuildColumnIndexMap(Excel.ListObject table)
    {
        Excel.ListColumns columns = table.ListColumns;
        int count = columns.Count;
        Dictionary<string, int> map = new(count, StringComparer.Ordinal);
        for (int index = 1; index <= count; index++)
        {
            _ = map.TryAdd(GetColumnAt(columns, index).Name, index - 1);
        }

        return map;
    }

    /// <summary>The zero-based offset of a named column, or -1 when the table lacks it.</summary>
    /// <param name="map">The column map.</param>
    /// <param name="name">The schema column name.</param>
    /// <returns>The zero-based offset, or -1.</returns>
    private static int IndexOf(Dictionary<string, int> map, string name) =>
        map.TryGetValue(name, out int index) ? index : -1;

    private static string ToText(object? value) =>
        Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;

    internal virtual Excel.Range GetCell(Excel.Range body, int row, int column)
    {
        Excel.Range cells = body.Cells;
        return cells[row, column];
    }

    internal virtual Excel.Range? GetTableBody(Excel.ListObject table) => table.DataBodyRange;

    internal virtual object? GetBodyValues(Excel.Range body) => body.Value2;

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

            Excel.ListObjects objects = candidate.ListObjects;
            int objectCount = objects.Count;
            for (int objectIndex = 1; objectIndex <= objectCount; objectIndex++)
            {
                Excel.ListObject candidateTable = GetTableAt(objects, objectIndex);
                if (GanttChartContract.IsDataTableName(candidateTable.Name))
                {
                    worksheet = candidate;
                    table = candidateTable;
                    return true;
                }
            }
        }

        return false;
    }

    internal virtual Excel.Worksheet GetSheetAt(Excel.Sheets sheets, int index) => sheets[index];

    internal virtual Excel.ListObject GetTableAt(Excel.ListObjects objects, int index) => objects[index];

    internal virtual Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index) => columns[index];
}

/// <summary>Catalog access for type metadata used by the normaliser.</summary>
internal static class GanttEntityCatalog
{
    public static bool TryParse(string? displayName, out GanttEntityType type) =>
        EntityTypeCatalog.TryParse(displayName, out type);

    public static EntityTypeDefinition? GetDefinition(GanttEntityType type) =>
        EntityTypeCatalog.GetDefinition(type);

    public static bool IsDateLabelPositionAllowed(GanttEntityType type, GanttDateLabelPosition position) =>
        EntityTypeCatalog.IsDateLabelPositionAllowed(type, position);
}

/// <summary>Label position parsing for the normaliser.</summary>
internal static class GanttLabelCatalog
{
    public static bool TryParse(string? text, out GanttLabelPosition position)
    {
        position = default;
        if (!string.IsNullOrWhiteSpace(text) &&
            Enum.TryParse(text.Trim(), out GanttLabelPosition parsed) &&
            Enum.IsDefined(parsed))
        {
            position = parsed;
            return true;
        }

        return false;
    }
}

/// <summary>Date label position parsing for the normaliser.</summary>
internal static class GanttDateLabelCatalog
{
    public static bool TryParse(string? text, out GanttDateLabelPosition position)
    {
        position = default;
        if (!string.IsNullOrWhiteSpace(text) &&
            Enum.TryParse(text.Trim(), out GanttDateLabelPosition parsed) &&
            Enum.IsDefined(parsed))
        {
            position = parsed;
            return true;
        }

        return false;
    }
}
