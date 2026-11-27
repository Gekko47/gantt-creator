using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Excel adapter for <see cref="ITypeChangeNormaliser"/>. Reads the visible Gantt table,
/// asks <see cref="EntityTypeTransition"/> what each row's type requires, and writes the
/// corrections back (R5.23).
/// </summary>
/// <remarks>
/// <para>
/// <b>The rules are not here.</b> Every decision about a type-dependent cell lives in
/// <see cref="EntityTypeTransition"/>, which uses the same catalogue and style
/// projections as the validator and the ribbon. This class is the writer: one column
/// walk, one body read, one write per changed cell.
/// </para>
/// <para>
/// <b>No old type.</b> The every-pass sweep sees only the row's current <c>Type</c>, so
/// it calls the transition with <c>oldType: null</c> and therefore applies only the
/// idempotent rules. The change-path heal supplies the old type from the cached pre-edit
/// rows; a valid stored value is never rewritten by an ordinary Refresh.
/// </para>
/// </remarks>
public class ExcelTypeChangeNormaliser(object? application, IWorksheetProtectionGuard? protectionGuard = null)
    : ITypeChangeNormaliser
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IWorksheetProtectionGuard _protectionGuard = protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <summary>
    /// Normalises with the caller's pre-edit types, which the change-path heal supplies
    /// from the cached rows (R5.23).
    /// </summary>
    /// <param name="styles">The workbook's named styles, or null when unavailable.</param>
    /// <param name="oldTypeOf">
    /// Resolves a row's previous Type by its engine Id, or null for the every-pass sweep,
    /// which cannot know it.
    /// </param>
    /// <returns>The write outcome.</returns>
    public GanttTypeChangeNormaliseOutcome Normalise(
        GanttStyleRegistry? styles,
        Func<GanttRowId, GanttEntityType?>? oldTypeOf)
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
                !EntityTypeCatalog.TryParse(typeText, out GanttEntityType type))
            {
                continue;
            }

            TypeTransitionRow input = new(
                StyleKey: TextAt(row, columnIndexes, "StyleKey"),
                LabelPosition: TextAt(row, columnIndexes, "LabelPosition"),
                DateLabelPosition: TextAt(row, columnIndexes, "DateLabelPosition"),
                FillColour: TextAt(row, columnIndexes, "FillColour"),
                StrokeColour: TextAt(row, columnIndexes, "StrokeColour"),
                ParentId: TextAt(row, columnIndexes, "ParentId"),
                LaneId: TextAt(row, columnIndexes, "LaneId"),
                StackIndex: TextAt(row, columnIndexes, "StackIndex"),
                SiblingOrder: TextAt(row, columnIndexes, "SiblingOrder"));

            GanttEntityType? oldType = null;
            if (oldTypeOf is not null
                && GanttRowId.TryParse(TextAt(row, columnIndexes, "Id"), out GanttRowId? rowId)
                && rowId is not null)
            {
                oldType = oldTypeOf(rowId);
            }

            IReadOnlyList<TypeTransitionChange> changes =
                EntityTypeTransition.Apply(input, type, oldType, styles);
            if (changes.Count == 0)
            {
                continue;
            }

            int excelRow = rowIndex + 1;
            foreach (TypeTransitionChange change in changes)
            {
                int column = IndexOf(columnIndexes, change.Column);
                if (column < 0)
                {
                    continue;
                }

                GetCell(body, excelRow, column + 1).Value2 = change.NewValue;
                written++;
            }
        }

        return GanttTypeChangeNormaliseOutcome.Ok(written);
    }

    /// <summary>Normalises with no style catalogue and no old type.</summary>
    /// <returns>The write outcome.</returns>
    public GanttTypeChangeNormaliseOutcome Normalise() => Normalise(null, null);

    /// <inheritdoc />
    public GanttTypeChangeNormaliseOutcome Normalise(GanttStyleRegistry? styles) =>
        Normalise(styles, null);

    /// <summary>
    /// Returns one cell's text by schema column name, or null when the table lacks it.
    /// </summary>
    /// <param name="row">The row's <c>Value2</c> values.</param>
    /// <param name="map">The column-name map.</param>
    /// <param name="name">The schema column name.</param>
    /// <returns>The cell text, or null when the column is absent or the cell is blank.</returns>
    private static string? TextAt(object?[] row, Dictionary<string, int> map, string name)
    {
        int index = IndexOf(map, name);
        return index < 0 ? null : ToText(row.ElementAtOrDefault(index));
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
