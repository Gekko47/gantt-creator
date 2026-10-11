using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using GanttCreator.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Applies the R5.3 parent/child plans to the visible Gantt table.
/// </summary>
/// <remarks>
/// <para>
/// The pipeline mirrors the row inserter's discipline (R2.8): the
/// protection guard runs first, before any workbook or table access
/// (ADR-0008 D4); the table is located by name; the target
/// worksheet's own protection is queried; the active cell is resolved
/// to a one-based body index through the same active-cell shape the
/// inserter uses; the body is bulk-read; the Core planner computes
/// the change; and exactly the planned cells are written back.
/// </para>
/// <para>
/// This adapter is the sole writer of <c>ParentId</c> and
/// <c>SiblingOrder</c> (R5.3 D4). Refresh re-reads both columns but
/// never rewrites them, and no other adapter derives them from row
/// position. A promote rewrites the moved window's full width because
/// the row's content physically moves down (R5.3 D2), plus the
/// resequenced <c>SiblingOrder</c> cells of run members above the
/// window, which the move does not touch; a demote likewise rewrites
/// its moved window's full width because the row's content physically
/// moves up to sit beneath the new parent's child run (R5.3 D3, amended
/// 2026-10-11), plus the resequenced cells outside the window and the
/// re-homed children's <c>ParentId</c> cells, which stay put.
/// </para>
/// <para>
/// Every refusal before the first cell write leaves the workbook
/// untouched. A host refusal DURING the write phase can follow
/// partially applied writes; the <see cref="ParentChildMutationRefusalReason.WriteFailed"/>
/// refusal reports that state rather than hiding it, which is the
/// ADR-0008 discipline the row inserter applies to its own
/// partially applied inserts.
/// </para>
/// <para>
/// The promoted row's own <c>SiblingOrder</c> cell is not rewritten:
/// the Core plan emits no write for it (it emits the remaining run's
/// resequencing only), and the planner -- not this adapter -- is the
/// authority on which <c>SiblingOrder</c> cells change.
/// </para>
/// </remarks>
public class ExcelParentChildMutator(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null) : IParentChildMutator
{
    // object? rather than the interop Application so the AddIn can pass
    // ExcelDnaUtil.Application without naming the interop type (the CS0433
    // duplicate-type hazard documented on IExcelApplicationAdapter) — the
    // same shape ExcelGanttRowInserter uses.
    private readonly Application? _application = application as Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <inheritdoc />
    public ParentChildMutationOutcome Promote() => Mutate(isPromote: true);

    /// <inheritdoc />
    public ParentChildMutationOutcome Demote() => Mutate(isPromote: false);

    /// <summary>
    /// The shared pipeline: guard, locate, resolve the selection, read,
    /// plan, then apply. Every refusal happens before the first COM
    /// write, and every local (in-memory) mutation happens before the
    /// first COM write, so a failure anywhere leaves the workbook
    /// untouched.
    /// </summary>
    /// <param name="isPromote">
    /// <see langword="true"/> for promote-to-parent,
    /// <see langword="false"/> for demote-to-child.
    /// </param>
    /// <returns>The typed mutation outcome.</returns>
    private ParentChildMutationOutcome Mutate(bool isPromote)
    {
        // ADR-0008 D4: intentionally the first operation in the mutating
        // entry point, before any application/workbook/table access.
        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return ParentChildMutationOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? ParentChildMutationRefusalReason.NoActiveWorkbook
                    : ParentChildMutationRefusalReason.TargetProtected);
        }

        Application? application = _application;
        Workbook? workbook = application?.ActiveWorkbook;
        if (workbook is null)
        {
            return ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.NoActiveWorkbook);
        }

        if (!TryFindTable(workbook.Sheets, out Worksheet? worksheet, out ListObject? table)
            || worksheet is null || table is null)
        {
            return ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.TableMissing);
        }

        // The authoritative check for the worksheet the mutation targets:
        // the active-sheet guard above says nothing about this sheet.
        ProtectionGuardOutcome targetProtection = _protectionGuard.QueryTarget(worksheet);
        if (targetProtection != ProtectionGuardOutcome.NotProtected)
        {
            return ParentChildMutationOutcome.Refused(
                targetProtection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? ParentChildMutationRefusalReason.NoActiveWorkbook
                    : ParentChildMutationRefusalReason.TargetProtected);
        }

        ListRows rows = GetListRows(table);
        var bodyIndex = ResolveActiveBodyIndex(table, rows);
        if (bodyIndex < 1)
        {
            return ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.SelectionOutsideTable);
        }

        if (!TryReadBody(table, out List<object?[]>? body, out var columnMap)
            || body is null || columnMap is null)
        {
            return ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.TableMissing);
        }

        // The selection must name a row the bulk read actually saw.
        if (bodyIndex > body.Count)
        {
            return ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.SelectionOutsideTable);
        }

        List<GanttEvent>? events = ProjectRows(
            body,
            columnMap,
            bodyIndex - 1,
            out ParentChildMutationRefusalReason? projectionRefusal);
        if (events is null)
        {
            return ParentChildMutationOutcome.Refused(projectionRefusal!.Value);
        }

        GanttRowId selectedId = events[bodyIndex - 1].Id;
        return isPromote
            ? ApplyPromote(table, columnMap, body, events, selectedId, bodyIndex)
            : ApplyDemote(table, columnMap, body, events, selectedId, bodyIndex);
    }

    /// <summary>
    /// Applies a promote plan (R5.3 D2): the target row's content
    /// moves to the bottom of its sibling run, its <c>ParentId</c>
    /// is cleared, and the remaining run is resequenced densely.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="columnMap">Schema-column to table-column index map.</param>
    /// <param name="body">The raw body matrix, in body order. Mutated in place.</param>
    /// <param name="events">The projected planner events, in body order.</param>
    /// <param name="selectedId">The identifier of the row being promoted.</param>
    /// <param name="bodyIndex">The one-based body index of the selection.</param>
    /// <returns>The typed mutation outcome.</returns>
    private ParentChildMutationOutcome ApplyPromote(
        ListObject table,
        int[] columnMap,
        List<object?[]> body,
        IReadOnlyList<GanttEvent> events,
        GanttRowId selectedId,
        int bodyIndex)
    {
        PromotePlan plan = CreatePromotePlan(events, selectedId);
        if (!plan.Succeeded)
        {
            return ParentChildMutationOutcome.PlanRefused(plan.Refusal!.Value, bodyIndex);
        }

        var fromIndex = bodyIndex - 1;
        var toIndex = plan.MoveToBodyIndex!.Value - 1;
        if (toIndex < fromIndex)
        {
            // The planner's move target is the bottom of the sibling run,
            // so it cannot precede the selection. Reaching this branch
            // means the plan and the table disagree about row positions;
            // refusing is safer than rotating the wrong window.
            return ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.TableInvalid);
        }

        // The content move, applied to the in-memory body before any COM
        // write: the rows between the selection and the run bottom shift
        // up one slot, and the promoted row's content lands at the bottom.
        var moved = body[fromIndex];
        for (var index = fromIndex; index < toIndex; index++)
        {
            body[index] = body[index + 1];
        }

        body[toIndex] = moved;

        // The promoted row leaves its parent: clear its ParentId cell at
        // its new position.
        SetCell(body[toIndex], columnMap, "ParentId", null);

        // The remaining run resequences densely. The writes are keyed by
        // row identifier, so they resolve against the POST-move positions.
        Dictionary<GanttRowId, int> positions = IndexById(body, columnMap);
        var outsideCells = new List<(int RowIndex, string ColumnName, object? Value)>();
        foreach (SiblingOrderWrite write in plan.SiblingOrderWrites)
        {
            if (!positions.TryGetValue(write.RowId, out var position))
            {
                continue;
            }

            SetCell(body[position], columnMap, "SiblingOrder", write.SiblingOrder);
            if (position < fromIndex || position > toIndex)
            {
                // A run member above the selection keeps its position when
                // the promoted row moves down, so its resequenced
                // SiblingOrder cell sits outside the moved window and is
                // written as a single cell: nothing else about the row
                // changes.
                outsideCells.Add((position, "SiblingOrder", write.SiblingOrder));
            }
        }

        // The content move rewrites every column of every row in the
        // moved window, full width.
        ParentChildMutationOutcome window = WriteRows(table, body, fromIndex, toIndex, bodyIndex);
        if (!window.Succeeded)
        {
            return window;
        }

        if (outsideCells.Count == 0)
        {
            return window;
        }

        ParentChildMutationOutcome outside = WriteCells(table, columnMap, outsideCells, bodyIndex);
        return outside.Succeeded
            ? ParentChildMutationOutcome.Ok(window.CellsWritten + outside.CellsWritten, bodyIndex)
            : outside;
    }

    /// <summary>
    /// Applies a demote plan (R5.3 D3, amended 2026-10-11): the target
    /// row's content moves to sit directly beneath the new parent's
    /// existing child run, it gains the new <c>ParentId</c>, its existing
    /// children are re-homed to the new parent, and the runs resequence
    /// densely.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="columnMap">Schema-column to table-column index map.</param>
    /// <param name="body">The raw body matrix, in body order. Mutated in place.</param>
    /// <param name="events">The projected planner events, in body order.</param>
    /// <param name="selectedId">The identifier of the row being demoted.</param>
    /// <param name="bodyIndex">The one-based body index of the selection.</param>
    /// <returns>The typed mutation outcome.</returns>
    private ParentChildMutationOutcome ApplyDemote(
        ListObject table,
        int[] columnMap,
        List<object?[]> body,
        IReadOnlyList<GanttEvent> events,
        GanttRowId selectedId,
        int bodyIndex)
    {
        DemotePlan plan = CreateDemotePlan(events, selectedId);
        if (!plan.Succeeded)
        {
            return ParentChildMutationOutcome.PlanRefused(plan.Refusal!.Value, bodyIndex);
        }

        var fromIndex = bodyIndex - 1;
        var toIndex = plan.MoveToBodyIndex!.Value - 1;
        if (toIndex > fromIndex)
        {
            // The planner's move target sits at or above the bottom of the
            // new parent's child run, which is above the selection. Reaching
            // this branch means the plan and the table disagree about row
            // positions; refusing is safer than rotating the wrong window.
            return ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.TableInvalid);
        }

        // The content move, applied to the in-memory body before any COM
        // write: the demoted row's content lands beneath the new parent's
        // child run, and the rows between shift down one slot to make room.
        var moved = body[fromIndex];
        for (var index = fromIndex; index > toIndex; index--)
        {
            body[index] = body[index - 1];
        }

        body[toIndex] = moved;

        // The demoted row joins the new parent, and its existing children
        // move with it so they stay depth 2.
        var newParentId = plan.NewParentId!.Value;
        SetCell(body[toIndex], columnMap, "ParentId", newParentId);
        foreach (GanttRowId rehomedId in plan.RehomedChildIds)
        {
            for (var rowIndex = 0; rowIndex < body.Count; rowIndex++)
            {
                if (CellText(body[rowIndex], columnMap, "Id") == rehomedId.Value)
                {
                    SetCell(body[rowIndex], columnMap, "ParentId", newParentId);
                    break;
                }
            }
        }

        // The planned SiblingOrder writes, keyed by row identifier so they
        // resolve against the POST-move positions.
        Dictionary<GanttRowId, int> positions = IndexById(body, columnMap);
        var outsideCells = new List<(int RowIndex, string ColumnName, object? Value)>();
        foreach (SiblingOrderWrite write in plan.SiblingOrderWrites)
        {
            if (!positions.TryGetValue(write.RowId, out var position))
            {
                continue;
            }

            SetCell(body[position], columnMap, "SiblingOrder", write.SiblingOrder);
            if (position < toIndex || position > fromIndex)
            {
                // A resequenced cell outside the moved window is written as
                // a single cell: nothing else about the row changes.
                outsideCells.Add((position, "SiblingOrder", write.SiblingOrder));
            }
        }

        // The re-homed children's ParentId cells sit outside the window
        // (they moved with the selection's own rows, which stay put) and
        // are written as single cells.
        var rehomeCells = new List<(int RowIndex, string ColumnName, object? Value)>();
        for (var index = 0; index < body.Count; index++)
        {
            if (index >= toIndex && index <= fromIndex)
            {
                continue;
            }

            if (plan.RehomedChildIds.Any(rehomedId => CellText(body[index], columnMap, "Id") == rehomedId.Value))
            {
                rehomeCells.Add((index, "ParentId", newParentId));
            }
        }

        // The content move rewrites every column of every row in the moved
        // window, full width.
        ParentChildMutationOutcome window = WriteRows(table, body, toIndex, fromIndex, bodyIndex);
        if (!window.Succeeded)
        {
            return window;
        }

        List<(int RowIndex, string ColumnName, object? Value)> singles = [.. outsideCells, .. rehomeCells];
        if (singles.Count == 0)
        {
            return window;
        }

        ParentChildMutationOutcome outside = WriteCells(table, columnMap, singles, bodyIndex);
        return outside.Succeeded
            ? ParentChildMutationOutcome.Ok(window.CellsWritten + outside.CellsWritten, bodyIndex)
            : outside;
    }

    /// <summary>
    /// Computes the promote plan. An internal seam so contract tests can
    /// drive the apply path with a crafted plan.
    /// </summary>
    /// <param name="rows">The projected events, in body order.</param>
    /// <param name="selectedId">The identifier of the row being promoted.</param>
    /// <returns>The promotion plan, or a typed refusal.</returns>
    internal virtual PromotePlan CreatePromotePlan(IReadOnlyList<GanttEvent> rows, GanttRowId selectedId) =>
        HierarchyMutations.PlanPromoteToParent(rows, selectedId);

    /// <summary>
    /// Computes the demote plan. An internal seam so contract tests can
    /// drive the apply path with a crafted plan.
    /// </summary>
    /// <param name="rows">The projected events, in body order.</param>
    /// <param name="selectedId">The identifier of the row being demoted.</param>
    /// <returns>The demotion plan, or a typed refusal.</returns>
    internal virtual DemotePlan CreateDemotePlan(IReadOnlyList<GanttEvent> rows, GanttRowId selectedId) =>
        HierarchyMutations.PlanDemoteToChild(rows, selectedId);

    /// <summary>
    /// Finds <c>tblGanttData</c> on any worksheet. Chart sheets carry no
    /// list objects and are skipped via the worksheet cast.
    /// </summary>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <param name="worksheet">The worksheet hosting the table.</param>
    /// <param name="table">The table when found.</param>
    /// <returns><see langword="true"/> when the table was found.</returns>
    internal virtual bool TryFindTable(
        Sheets sheets,
        out Worksheet? worksheet,
        out ListObject? table)
    {
        worksheet = null;
        table = null;
        var count = sheets.Count;
        for (var index = 1; index <= count; index++)
        {
            var sheet = GetSheetAt(sheets, index);
            if (sheet is not Worksheet candidate)
            {
                continue;
            }

            ListObjects listObjects = candidate.ListObjects;
            var tableCount = listObjects.Count;
            for (var tableIndex = 1; tableIndex <= tableCount; tableIndex++)
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
    /// header names (Ordinal), so a reordered table still maps
    /// correctly. A missing required column refuses the mutation.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="columnMap">Per-schema-column one-based table indexes, or <see langword="null"/>.</param>
    /// <returns><see langword="true"/> when every required header resolved.</returns>
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
    /// Bulk-reads the table body into raw cell rows, in body order,
    /// exactly as <c>Value2</c> returned them. An empty body reads as
    /// an empty list.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="body">The raw body rows, one per body row.</param>
    /// <param name="columnMap">The schema-column to table-column index map.</param>
    /// <returns><see langword="true"/> when the body was read.</returns>
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
    /// Resolves the active cell to a one-based body index through the
    /// same active-cell shape the row inserter uses: an active table
    /// claims the active cell's row, offset by the table's first
    /// (header) row.
    /// </summary>
    /// <remarks>
    /// Returns 0 when the selection is outside the table body -- no
    /// active cell, the table not active, the header row itself, or the
    /// reserved padding row below the body. None of those names a data
    /// row a parent/child command can act on.
    /// </remarks>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="rows">The table's body rows.</param>
    /// <returns>The one-based body index, or 0 when the selection names no data row.</returns>
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

        // The header row (offset 0) and anything at or below the reserved
        // padding row (offset past the body) name no data row.
        return bodyOffset < 1 || bodyOffset > rowCount ? 0 : bodyOffset;
    }

    /// <summary>
    /// Projects the raw body into the neutral events the Core planner
    /// consumes. Only the fields the hierarchy planners read are
    /// populated; the rest stay at their defaults.
    /// </summary>
    /// <remarks>
    /// A row whose <c>Id</c> cannot be parsed makes the table unusable:
    /// the selection row yields <see cref="ParentChildMutationRefusalReason.SelectionUnknown"/>
    /// and any other row yields <see cref="ParentChildMutationRefusalReason.TableInvalid"/>.
    /// An unparseable <c>Type</c> is projected as the enum default
    /// (<see cref="GanttEntityType.Splitter"/>, standalone): the
    /// conservative reading that refuses both commands with
    /// <c>TypeCannotBeChild</c> rather than guessing a role.
    /// </remarks>
    /// <param name="body">The raw body rows, in body order.</param>
    /// <param name="columnMap">The schema-column to table-column index map.</param>
    /// <param name="selectedIndex">The zero-based index of the selected row.</param>
    /// <param name="refusal">The refusal reason when the projection returns <see langword="null"/>.</param>
    /// <returns>The projected events, or <see langword="null"/>.</returns>
    private static List<GanttEvent>? ProjectRows(
        List<object?[]> body,
        int[] columnMap,
        int selectedIndex,
        out ParentChildMutationRefusalReason? refusal)
    {
        refusal = null;
        var events = new List<GanttEvent>(body.Count);
        for (var index = 0; index < body.Count; index++)
        {
            var row = body[index];
            if (!GanttRowId.TryParse(CellText(row, columnMap, "Id"), out GanttRowId? id) || id is null)
            {
                refusal = index == selectedIndex
                    ? ParentChildMutationRefusalReason.SelectionUnknown
                    : ParentChildMutationRefusalReason.TableInvalid;
                return null;
            }

            _ = EntityTypeCatalog.TryParse(
                CellText(row, columnMap, "Type"),
                out GanttEntityType type);
            var parentText = CellText(row, columnMap, "ParentId");
            GanttRowId? parentId =
                GanttRowId.TryParse(parentText, out GanttRowId? parsedParent) && parsedParent is not null
                    ? parsedParent
                    : null;

            events.Add(new GanttEvent(
                index + 1,
                id,
                null,
                null,
                type,
                null,
                null,
                null,
                parentId,
                null,
                null,
                null,
                null,
                true,
                null));
        }

        return events;
    }

    /// <summary>
    /// Reads one text cell from a raw body row by schema column name.
    /// </summary>
    /// <param name="row">The raw row, one value per table column.</param>
    /// <param name="columnMap">The schema-column to table-column index map.</param>
    /// <param name="columnName">The schema column name.</param>
    /// <returns>The cell's text, or <see cref="string.Empty"/>.</returns>
    private static string CellText(
        object?[] row,
        int[] columnMap,
        string columnName)
    {
        var schemaIndex = SchemaIndexOf(columnName);
        var tableIndex = columnMap[schemaIndex];
        return tableIndex < 1 || tableIndex > row.Length
            ? string.Empty
            : row[tableIndex - 1] as string ?? string.Empty;
    }

    /// <summary>
    /// Locates a schema column's index by name.
    /// </summary>
    /// <param name="columnName">The schema column name.</param>
    /// <returns>The zero-based schema index.</returns>
    private static int SchemaIndexOf(string columnName)
    {
        for (var index = 0; index < GanttTableSchema.Default.Columns.Count; index++)
        {
            if (GanttTableSchema.Default.Columns[index].Name == columnName)
            {
                return index;
            }
        }

        // Unreachable for the catalogue-owned column names this adapter
        // addresses. A new column name is a schema change, not a value
        // this method may guess.
        throw new InvalidOperationException($"'{columnName}' is not a Gantt schema column.");
    }

    /// <summary>
    /// Writes the moved window of body rows back to the table, full
    /// width. This is the promote content move: every column of every
    /// row in the window is rewritten because the rows' content
    /// physically moved within the table body.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="body">The mutated body matrix.</param>
    /// <param name="fromIndex">The zero-based index the window starts at.</param>
    /// <param name="toIndex">The zero-based index the window ends at.</param>
    /// <param name="bodyIndex">The one-based body index of the selection.</param>
    /// <returns>The typed mutation outcome.</returns>
    private ParentChildMutationOutcome WriteRows(
        ListObject table,
        List<object?[]> body,
        int fromIndex,
        int toIndex,
        int bodyIndex)
    {
        Excel.Range? dataBody = GetDataBodyRange(table);
        if (dataBody is null)
        {
            return ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.TableMissing);
        }

        Excel.Range rows = GetRangeRows(dataBody);
        try
        {
            for (var index = fromIndex; index <= toIndex; index++)
            {
                Excel.Range row = GetRangeAt(rows, index + 1);
                SetRangeValues(row, ToRowPayload(body[index]));
            }
        }
        catch (COMException)
        {
            return ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.WriteFailed);
        }

        var cellsWritten = (toIndex - fromIndex + 1) * body[fromIndex].Length;
        return ParentChildMutationOutcome.Ok(cellsWritten, bodyIndex);
    }

    /// <summary>
    /// Writes the planned cells back to the table, one cell per write.
    /// This is the demote mutation set: only the planned
    /// <c>ParentId</c> and <c>SiblingOrder</c> cells change, so only
    /// those cells are written.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="columnMap">The schema-column to table-column index map.</param>
    /// <param name="writes">The planned cell writes, as (zero-based row index, column name, value).</param>
    /// <param name="bodyIndex">The one-based body index of the selection.</param>
    /// <returns>The typed mutation outcome.</returns>
    private ParentChildMutationOutcome WriteCells(
        ListObject table,
        int[] columnMap,
        List<(int RowIndex, string ColumnName, object? Value)> writes,
        int bodyIndex)
    {
        Excel.Range? dataBody = GetDataBodyRange(table);
        if (dataBody is null)
        {
            return ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.TableMissing);
        }

        Excel.Range rows = GetRangeRows(dataBody);
        try
        {
            foreach ((var rowIndex, var columnName, var value) in writes)
            {
                var tableIndex = columnMap[SchemaIndexOf(columnName)];
                Excel.Range row = GetRangeAt(rows, rowIndex + 1);
                Excel.Range cell = GetCellAt(row, tableIndex);
                SetRangeValues(cell, value!);
            }
        }
        catch (COMException)
        {
            return ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.WriteFailed);
        }

        return ParentChildMutationOutcome.Ok(writes.Count, bodyIndex);
    }

    /// <summary>
    /// Sets one cell in the in-memory body matrix.
    /// </summary>
    /// <remarks>
    /// <see cref="TryBuildColumnMap"/> guarantees every required column
    /// -- which includes both columns this adapter writes -- resolved to
    /// a one-based table index within the row's width, so the index is
    /// direct rather than guarded.
    /// </remarks>
    /// <param name="row">The raw body row.</param>
    /// <param name="columnMap">The schema-column to table-column index map.</param>
    /// <param name="columnName">The schema column name.</param>
    /// <param name="value">The cell value.</param>
    private static void SetCell(
        object?[] row,
        int[] columnMap,
        string columnName,
        object? value) => row[columnMap[SchemaIndexOf(columnName)] - 1] = value;

    /// <summary>
    /// Maps every parseable row identifier to its zero-based position
    /// in the body, for resolving plan writes keyed by row identifier.
    /// </summary>
    /// <param name="body">The raw body rows.</param>
    /// <param name="columnMap">The schema-column to table-column index map.</param>
    /// <returns>Row identifier to zero-based body position.</returns>
    private static Dictionary<GanttRowId, int> IndexById(
        List<object?[]> body,
        int[] columnMap)
    {
        var positions = new Dictionary<GanttRowId, int>();
        for (var index = 0; index < body.Count; index++)
        {
            if (GanttRowId.TryParse(CellText(body[index], columnMap, "Id"), out GanttRowId? id)
                && id is not null)
            {
                positions[id] = index;
            }
        }

        return positions;
    }

    /// <summary>
    /// Converts one raw body row into the rectangular payload Excel's
    /// <c>Value2</c> assignment requires.
    /// </summary>
    /// <param name="row">The raw body row, one value per table column.</param>
    /// <returns>A one-based <c>object[,]</c> for assignment.</returns>
    [SuppressMessage(
        "Performance",
        "CA1814:MultiDimArrays",
        Justification = "Excel's Value2 is a SAFEARRAY and is written through a multidimensional array. A jagged array cannot express it, so the analyzer's preferred shape is not available here.")]
    [SuppressMessage(
        "Performance",
        "CA1859:UseConcreteTypesWhenPossible",
        Justification = "The payload is assigned to a COM property typed as object, so the declared return type must stay object.")]
    private static object ToRowPayload(object?[] row)
    {
        var payload = new object[1, row.Length];
        for (var column = 0; column < row.Length; column++)
        {
            payload[0, column] = row[column]!;
        }

        return payload;
    }

    /// <summary>
    /// Returns the sheet at the one-based index. Test seam over the
    /// COM parameterised <c>Sheets.Item</c> property.
    /// </summary>
    /// <remarks>
    /// Returns <see cref="object"/> rather than <see cref="Worksheet"/>
    /// so a chart sheet does not throw here; the caller filters by
    /// cast, exactly as the table reader does.
    /// </remarks>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <param name="index">The one-based sheet index.</param>
    /// <returns>The sheet at the index.</returns>
    internal virtual object GetSheetAt(Sheets sheets, int index) => sheets[index];

    /// <summary>
    /// Returns the list object at the one-based index. Test seam over
    /// the COM parameterised <c>ListObjects.Item</c> property.
    /// </summary>
    /// <param name="listObjects">The worksheet's list objects.</param>
    /// <param name="index">The one-based table index.</param>
    /// <returns>The list object at the index.</returns>
    internal virtual ListObject GetTableAt(ListObjects listObjects, int index) =>
        listObjects[index];

    /// <summary>
    /// Returns the list column at the one-based index. Test seam over
    /// the COM parameterised <c>ListColumns.Item</c> property.
    /// </summary>
    /// <param name="columns">The table's list columns.</param>
    /// <param name="index">The one-based column index.</param>
    /// <returns>The list column at the index.</returns>
    internal virtual ListColumn GetColumnAt(ListColumns columns, int index) =>
        columns[index];

    /// <summary>
    /// Returns the table's body rows. Test seam over the COM
    /// <c>ListObject.ListRows</c> property.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <returns>The table's body rows.</returns>
    internal virtual ListRows GetListRows(ListObject table) => table.ListRows;

    /// <summary>
    /// Returns the body row count. Test seam over the COM
    /// <c>ListRows.Count</c> property.
    /// </summary>
    /// <param name="rows">The table's body rows.</param>
    /// <returns>The number of body rows.</returns>
    internal virtual int GetListRowCount(ListRows rows) => rows.Count;

    /// <summary>
    /// Returns the table's full range. Test seam over the COM
    /// <c>ListObject.Range</c> property.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <returns>The table's range, header row included.</returns>
    internal virtual Excel.Range GetTableRange(ListObject table) => table.Range;

    /// <summary>
    /// Returns the table's body range. Test seam over the COM
    /// <c>ListObject.DataBodyRange</c> property.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <returns>The body range, or <see langword="null"/> for an empty body.</returns>
    internal virtual Excel.Range? GetDataBodyRange(ListObject table) => table.DataBodyRange;

    /// <summary>
    /// Returns a range's rows. Test seam over the COM
    /// <c>Range.Rows</c> property.
    /// </summary>
    /// <param name="range">The range whose rows are wanted.</param>
    /// <returns>The range's rows.</returns>
    internal virtual Excel.Range GetRangeRows(Excel.Range range) => range.Rows;

    /// <summary>
    /// Returns the row at the one-based index within a rows range.
    /// Test seam over the COM parameterised <c>Range.Item</c> property.
    /// </summary>
    /// <param name="rows">The rows range.</param>
    /// <param name="index">The one-based row index.</param>
    /// <returns>The row range.</returns>
    internal virtual Excel.Range GetRangeAt(Excel.Range rows, int index) => rows[index];

    /// <summary>
    /// Returns one cell of a row range by one-based column index.
    /// Test seam over the COM <c>Range.Cells[row, column]</c> dispatch.
    /// </summary>
    /// <remarks>
    /// <b>Why <c>.Cells</c> rather than the bare range indexer.</b> The
    /// two-argument form of <see cref="Excel.Range"/>'s own indexer
    /// (<c>row[1, columnIndex]</c>, i.e. the <c>Item</c>/<c>_Default</c>
    /// property) is <b>refused by the host</b> with
    /// <c>COMException 0x800A03EC</c> on every range shape this adapter
    /// addresses, while the identically-shaped <c>row.Cells[1, columnIndex]</c>
    /// is accepted — measured on a live Excel 16.0
    /// <c>tblGanttData</c> body row (<c>$A$5:$P$5</c>, 1 row x 16 columns),
    /// where the row is a one-row range obtained from
    /// <c>DataBodyRange.Rows[i]</c>. Because the two forms are
    /// indistinguishable by the shape of the range they address, only a
    /// live probe or a strict mock pins which one is correct; the
    /// contract harness overrides this seam and so never executed the
    /// production line, which is how the refusal reached the user as
    /// <see cref="ParentChildMutationRefusalReason.WriteFailed"/> on both
    /// verbs. <c>Range.Cells</c> is also the convention the live-proven
    /// <c>ExcelGanttRowIdentityRepairer.GetCell</c> already uses for the
    /// same <c>ParentId</c> write. See the source-shape guard in
    /// <c>ParentChildCellAccessorTests</c> and the live
    /// <c>ParentChildMutationIntegrationTests</c>.
    /// </remarks>
    /// <param name="row">The row range.</param>
    /// <param name="columnIndex">The one-based column index within the row.</param>
    /// <returns>The cell range.</returns>
    internal virtual Excel.Range GetCellAt(Excel.Range row, int columnIndex) => row.Cells[1, columnIndex];

    /// <summary>
    /// Returns the bulk <c>Value2</c> payload of a range. Test seam so
    /// contract tests inject <c>object[,]</c> without COM.
    /// </summary>
    /// <param name="range">The range to read.</param>
    /// <returns>The <c>Value2</c> payload.</returns>
    internal virtual object? GetRangeValue2(Excel.Range range) => range.Value2;

    /// <summary>
    /// Assigns a range's <c>Value2</c> payload. Test seam.
    /// </summary>
    /// <param name="range">The range to write.</param>
    /// <param name="payload">The payload.</param>
    internal virtual void SetRangeValues(Excel.Range range, object payload) => range.Value2 = payload;

    /// <summary>
    /// Returns the application's active cell. Test seam over the COM
    /// <c>Application.ActiveCell</c> property.
    /// </summary>
    /// <param name="application">The Excel application.</param>
    /// <returns>The active cell, or <see langword="null"/>.</returns>
    internal virtual Excel.Range? GetActiveCell(Application application) =>
        application.ActiveCell;

    /// <summary>
    /// Returns whether the table owns the selection. Test seam over the
    /// COM <c>ListObject.Active</c> property.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <returns><see langword="true"/> when the table is active.</returns>
    internal virtual bool GetTableActive(ListObject table) => table.Active;

    /// <summary>
    /// Returns the range's first worksheet row. Test seam over the COM
    /// <c>Range.Row</c> property.
    /// </summary>
    /// <param name="range">The range.</param>
    /// <returns>The one-based worksheet row.</returns>
    internal virtual int GetRangeRow(Excel.Range range) => range.Row;
}
