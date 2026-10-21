using System.Runtime.InteropServices;
using GanttCreator.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Inserts one new row as a CHILD of the row under the active cell
/// (R5.3 WI-E).
/// </summary>
/// <remarks>
/// <para>
/// The pipeline mirrors the parent/child mutator's discipline (R5.3 D4): the
/// protection guard runs first, before any workbook or table access
/// (ADR-0008 D4); the table is located by name; the target worksheet's own
/// protection is queried; the active cell is resolved to a one-based body
/// index; the body is bulk-read; the Core planner decides whether the
/// requested type may be added under the selected parent; the row is
/// inserted; and the new row's <c>ParentId</c> and <c>SiblingOrder</c> cells
/// are written.
/// </para>
/// <para>
/// <b>Why the row is inserted directly beneath the parent's child run.</b>
/// An Excel outline group covers a CONTIGUOUS run of rows. A child inserted
/// anywhere else would split the run, and the group writer would then have
/// to ungroup and regroup around it. Landing at the end of the run keeps the
/// group one group, and it is where the Core planner places the new row.
/// </para>
/// <para>
/// <b>Why this adapter writes the hierarchy cells rather than the row
/// inserter.</b> The parent/child mutator is the sole writer of
/// <c>ParentId</c> and <c>SiblingOrder</c>, and a second writer would be a
/// second authority for the same two columns. The inserter contributes only
/// the row creation.
/// </para>
/// </remarks>
public class ExcelAddChildMutator(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null,
    IGanttRowInserter? rowInserter = null,
    IOutlineGroupPort? outlineGroups = null) : IAddChildMutator
{
    // object? rather than the interop Application so the AddIn can pass
    // ExcelDnaUtil.Application without naming the interop type (the CS0433
    // duplicate-type hazard documented on IExcelApplicationAdapter).
    private readonly Application? _application = application as Application;

    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <summary>
    /// The row creator. It contributes the row and its scaffold values; the
    /// hierarchy cells are written here, by the sole writer of those columns.
    /// </summary>
    private readonly IGanttRowInserter? _rowInserter = rowInserter;

    /// <summary>
    /// The outline adapter, consulted after a successful insert so the native
    /// <c>-</c>/<c>+</c> controls reflect the new child without waiting for
    /// the next Refresh (R5.13).
    /// </summary>
    private readonly IOutlineGroupPort? _outlineGroups = outlineGroups;

    /// <inheritdoc />
    public AddChildMutationOutcome AddChild(GanttEntityType childType, Func<GanttRowId> nextId)
    {
        ArgumentNullException.ThrowIfNull(nextId);

        // ADR-0008 D4: this is intentionally the first operation in the
        // mutating entry point, before any application/workbook/table access.
        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return AddChildMutationOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? AddChildMutationRefusalReason.NoActiveWorkbook
                    : AddChildMutationRefusalReason.TargetProtected);
        }

        Application? application = _application;
        Workbook? workbook = application?.ActiveWorkbook;
        if (workbook is null)
        {
            return AddChildMutationOutcome.Refused(AddChildMutationRefusalReason.NoActiveWorkbook);
        }

        Sheets sheets = workbook.Sheets;
        if (!TryFindTable(sheets, out Worksheet? worksheet, out ListObject? table)
            || worksheet is null || table is null
            || !TryBuildColumnMap(table, out int[]? columnMap) || columnMap is null)
        {
            return AddChildMutationOutcome.Refused(AddChildMutationRefusalReason.TableMissing);
        }

        ProtectionGuardOutcome targetProtection = _protectionGuard.QueryTarget(worksheet);
        if (targetProtection != ProtectionGuardOutcome.NotProtected)
        {
            return AddChildMutationOutcome.Refused(
                targetProtection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? AddChildMutationRefusalReason.NoActiveWorkbook
                    : AddChildMutationRefusalReason.TargetProtected);
        }

        ListRows rows = GetListRows(table);
        int bodyIndex = ResolveActiveBodyIndex(table, rows);
        if (bodyIndex < 1)
        {
            return AddChildMutationOutcome.Refused(AddChildMutationRefusalReason.SelectionOutsideTable);
        }

        if (!TryReadBody(table, out List<object?[]>? body, out int[]? readColumnMap)
            || body is null || readColumnMap is null)
        {
            return AddChildMutationOutcome.Refused(AddChildMutationRefusalReason.TableMissing);
        }

        if (bodyIndex > body.Count)
        {
            return AddChildMutationOutcome.Refused(AddChildMutationRefusalReason.SelectionOutsideTable);
        }

        List<GanttEvent>? events = ProjectRows(
            body,
            readColumnMap,
            bodyIndex - 1,
            out AddChildMutationRefusalReason? projectionRefusal);
        if (events is null)
        {
            return AddChildMutationOutcome.Refused(projectionRefusal!.Value);
        }

        GanttRowId parentId = events[bodyIndex - 1].Id;
        GanttRowId newRowId = nextId();

        // The planner is the authority for whether the new row may join this
        // parent, and for the sibling order it takes. Every refusal happens
        // BEFORE the row is inserted, so a refused add leaves the workbook
        // untouched -- not even an empty row.
        AddChildPlan plan = AddChildPlanner.Plan(events, parentId, childType, newRowId);
        if (!plan.Succeeded)
        {
            return AddChildMutationOutcome.PlanRefused(plan.Refusal!.Value, bodyIndex);
        }

        AddChildWrite write = plan.Writes![0];

        // The row is inserted by the row inserter, which owns the row
        // creation and the scaffold values. It is optional so a caller that
        // only exercises the planning half does not have to supply one.
        if (_rowInserter is not { } inserter)
        {
            return AddChildMutationOutcome.Refused(AddChildMutationRefusalReason.InsertRefused);
        }

        GanttRowInsertOutcome inserted = inserter.Insert(childType, nextId);
        if (!inserted.Succeeded)
        {
            return AddChildMutationOutcome.Refused(AddChildMutationRefusalReason.InsertRefused);
        }

        // The new row's hierarchy cells are written HERE, by the sole writer
        // of ParentId and SiblingOrder (R5.3 D4).
        int insertedIndex = inserted.BodyIndex!.Value - 1;
        List<(int RowIndex, string ColumnName, object? Value)> singles =
        [
            (insertedIndex, "ParentId", write.ParentId.Value),
            (insertedIndex, "SiblingOrder", write.SiblingOrder),
        ];

        ParentChildMutationOutcome written = WriteCells(table, readColumnMap, singles);
        if (!written.Succeeded)
        {
            return AddChildMutationOutcome.Refused(AddChildMutationRefusalReason.WriteFailed);
        }

        // The outline is rebuilt so the native -/+ controls reflect the new
        // child without waiting for the next Refresh (R5.13). A regroup
        // refusal degrades to the pre-D12 behaviour: the insert stands and
        // the next Refresh rebuilds the outline.
        if (_outlineGroups is not null)
        {
            _ = RebuildOutline(table);
        }

        return AddChildMutationOutcome.Ok(singles.Count, bodyIndex, inserted.BodyIndex!.Value);
    }

    /// <summary>
    /// Finds the Gantt worksheet and its table. Test seam over the COM
    /// collection indexers.
    /// </summary>
    /// <param name="sheets">The workbook's sheet collection.</param>
    /// <param name="worksheet">The resolved worksheet.</param>
    /// <param name="table">The resolved table.</param>
    /// <returns>Whether both were found.</returns>
    internal virtual bool TryFindTable(
        Sheets sheets,
        out Worksheet? worksheet,
        out ListObject? table)
    {
        worksheet = null;
        table = null;
        foreach (object? entry in sheets)
        {
            if (entry is not Worksheet candidate)
            {
                continue;
            }

            foreach (ListObject candidateTable in candidate.ListObjects)
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

    /// <summary>
    /// Builds the schema-column to table-column index map. Test seam over the
    /// COM collection.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="columnMap">The map, or null when it could not be built.</param>
    /// <returns>Whether the map was built.</returns>
    internal virtual bool TryBuildColumnMap(ListObject table, out int[]? columnMap)
    {
        columnMap = null;
        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        int[] map = new int[schema.Count];
        Dictionary<string, int> byName = new Dictionary<string, int>(StringComparer.Ordinal);
        ListColumns columns = table.ListColumns;
        int columnCount = GetListColumnCount(columns);
        for (int index = 1; index <= columnCount; index++)
        {
            ListColumn column = GetListColumnAt(columns, index);
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

    /// <summary>Reads the table's data body as a row matrix. Test seam.</summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="body">The raw body rows, or null when it could not be read.</param>
    /// <param name="columnMap">The schema-column to table-column index map.</param>
    /// <returns>Whether the body was read.</returns>
    internal virtual bool TryReadBody(ListObject table, out List<object?[]>? body, out int[]? columnMap)
    {
        body = null;
        columnMap = null;
        if (!TryBuildColumnMap(table, out columnMap) || columnMap is null)
        {
            return false;
        }

        Excel.Range? dataBody = GetDataBodyRange(table);
        if (dataBody is null)
        {
            return false;
        }

        body = ExcelValue2Matrix.ReadRows(GetRangeValue2(dataBody));
        return true;
    }

    /// <summary>
    /// Resolves the active cell to a one-based body index. Test seam over the
    /// active-cell shape.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="rows">The table's list rows.</param>
    /// <returns>The one-based body index, or zero when the selection is not a body row.</returns>
    internal virtual int ResolveActiveBodyIndex(ListObject table, ListRows rows)
    {
        Excel.Range? activeCell = GetActiveCell(_application);
        if (activeCell is null || !GetTableActive(table))
        {
            return 0;
        }

        Excel.Range tableRange = GetTableRange(table);
        int tableFirstRow = GetRangeRow(tableRange);
        int activeRow = GetRangeRow(activeCell);
        int bodyIndex = activeRow - tableFirstRow;
        return bodyIndex < 1 || bodyIndex > GetListRowCount(rows) ? 0 : bodyIndex;
    }

    /// <summary>
    /// Projects the body into planner events. Test seam over the projection
    /// authority.
    /// </summary>
    /// <param name="body">The raw body rows.</param>
    /// <param name="columnMap">The schema-column to table-column index map.</param>
    /// <param name="selectedIndex">The selected row's zero-based body index.</param>
    /// <param name="refusal">The refusal reason, when the projection failed.</param>
    /// <returns>The projected events, or null.</returns>
    internal virtual List<GanttEvent>? ProjectRows(
        List<object?[]> body,
        int[] columnMap,
        int selectedIndex,
        out AddChildMutationRefusalReason? refusal)
    {
        refusal = null;
        List<GanttEvent>? events = ExcelGanttRowProjection.ProjectRows(body, columnMap, out int? failedIndex);
        if (events is null)
        {
            refusal = failedIndex == selectedIndex
                ? AddChildMutationRefusalReason.SelectionOutsideTable
                : AddChildMutationRefusalReason.TableInvalid;
            return null;
        }

        return events;
    }

    /// <summary>
    /// Writes the planned cells. Every write addresses the row's current
    /// position, and a host refusal is reported rather than swallowed.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="columnMap">The schema-column to table-column index map.</param>
    /// <param name="writes">The cells to write.</param>
    /// <returns>The outcome.</returns>
    internal virtual ParentChildMutationOutcome WriteCells(
        ListObject table,
        int[] columnMap,
        List<(int RowIndex, string ColumnName, object? Value)> writes)
    {
        Excel.Range? dataBody = GetDataBodyRange(table);
        if (dataBody is null)
        {
            return ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.TableMissing);
        }

        Excel.Range rows = GetRangeRows(dataBody);
        try
        {
            foreach ((int rowIndex, string? columnName, object? value) in writes)
            {
                int tableIndex = columnMap[ExcelGanttRowProjection.SchemaIndexOf(columnName)];
                Excel.Range row = GetRangeAt(rows, rowIndex + 1);
                Excel.Range cell = GetCellAt(row, tableIndex);
                SetRangeValues(cell, value!);
            }
        }
        catch (COMException)
        {
            return ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.WriteFailed);
        }

        return ParentChildMutationOutcome.Ok(writes.Count, 1);
    }

    /// <summary>
    /// Rebuilds the native outline from the post-insert table (R5.13), so
    /// the -/+ controls reflect the new child without waiting for the next
    /// Refresh.
    /// </summary>
    /// <param name="table">The Gantt data table, post-insert.</param>
    /// <returns>Whether the outline was rebuilt.</returns>
    internal virtual bool RebuildOutline(ListObject table)
    {
        if (_outlineGroups is null)
        {
            return false;
        }

        if (!TryReadBody(table, out List<object?[]>? body, out int[]? columnMap)
            || body is null || columnMap is null)
        {
            return false;
        }

        List<GanttEvent>? events = ExcelGanttRowProjection.ProjectRows(body, columnMap, out int? _);
        if (events is null)
        {
            return false;
        }

        _ = _outlineGroups.Apply(events);
        return true;
    }

    /// <summary>Gets the table's list rows. Test seam over the COM collection.</summary>
    /// <param name="table">The Gantt data table.</param>
    /// <returns>The list rows.</returns>
    internal virtual ListRows GetListRows(ListObject table) => table.ListRows;

    /// <summary>Gets the list-column count. Test seam over the COM collection.</summary>
    /// <param name="columns">The list columns.</param>
    /// <returns>The column count.</returns>
    internal virtual int GetListColumnCount(ListColumns columns) => columns.Count;

    /// <summary>Gets the list column at an index. Test seam over the COM collection.</summary>
    /// <param name="columns">The list columns.</param>
    /// <param name="index">The one-based index.</param>
    /// <returns>The column.</returns>
    internal virtual ListColumn GetListColumnAt(ListColumns columns, int index) => columns[index];

    /// <summary>Gets the list-row count. Test seam over the COM collection.</summary>
    /// <param name="rows">The list rows.</param>
    /// <returns>The row count.</returns>
    internal virtual int GetListRowCount(ListRows rows) => rows.Count;

    /// <summary>Gets the table's range. Test seam over the COM property.</summary>
    /// <param name="table">The Gantt data table.</param>
    /// <returns>The table's range.</returns>
    internal virtual Excel.Range GetTableRange(ListObject table) => table.Range;

    /// <summary>Gets the table's data-body range. Test seam over the COM property.</summary>
    /// <param name="table">The Gantt data table.</param>
    /// <returns>The data-body range, or null.</returns>
    internal virtual Excel.Range? GetDataBodyRange(ListObject table) => table.DataBodyRange;

    /// <summary>Gets the rows of a range. Test seam over the COM property.</summary>
    /// <param name="range">The range.</param>
    /// <returns>The rows.</returns>
    internal virtual Excel.Range GetRangeRows(Excel.Range range) => range.Rows;

    /// <summary>Gets the range at an index. Test seam over the COM indexer.</summary>
    /// <param name="rows">The rows collection.</param>
    /// <param name="index">The one-based index.</param>
    /// <returns>The range.</returns>
    internal virtual Excel.Range GetRangeAt(Excel.Range rows, int index) => rows[index];

    /// <summary>Gets the cell at a column. Test seam over the COM indexer.</summary>
    /// <param name="row">The row.</param>
    /// <param name="columnIndex">The one-based column index.</param>
    /// <returns>The cell.</returns>
    internal virtual Excel.Range GetCellAt(Excel.Range row, int columnIndex) => row.Cells[1, columnIndex];

    /// <summary>Reads a range's values. Test seam over the COM property.</summary>
    /// <param name="range">The range.</param>
    /// <returns>The range's values.</returns>
    internal virtual object? GetRangeValue2(Excel.Range range) => range.Value2;

    /// <summary>Writes a range's value. Test seam over the COM property.</summary>
    /// <param name="range">The range.</param>
    /// <param name="payload">The value.</param>
    internal virtual void SetRangeValues(Excel.Range range, object payload) => range.Value2 = payload;

    /// <summary>Gets the active cell. Test seam over the COM property.</summary>
    /// <param name="application">The Excel application.</param>
    /// <returns>The active cell, or null.</returns>
    internal virtual Excel.Range? GetActiveCell(Application? application) =>
        application?.ActiveCell;

    /// <summary>Whether the table is active. Test seam over the COM property.</summary>
    /// <param name="table">The Gantt data table.</param>
    /// <returns>Whether the table is active.</returns>
    internal virtual bool GetTableActive(ListObject table) => table.Active;

    /// <summary>Reads a range's worksheet row. Test seam over the COM property.</summary>
    /// <param name="range">The range.</param>
    /// <returns>The worksheet row.</returns>
    internal virtual int GetRangeRow(Excel.Range range) => range.Row;
}
