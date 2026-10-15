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
/// re-homed children's <c>ParentId</c> cells, which stay put. A
/// reorder move (R5.4 D5) rewrites both swapped windows' full width —
/// a content rotation (R5.4 D11-amended), not a worksheet-row move —
/// plus the resequenced
/// <c>SiblingOrder</c> cells outside their union, and a move at the
/// top or bottom of the selection's set writes nothing at all (R5.4
/// D4, silent success).
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
    IWorksheetProtectionGuard? protectionGuard = null,
    IOutlineGroupPort? outlineGroups = null) : IParentChildMutator
{
    // object? rather than the interop Application so the AddIn can pass
    // ExcelDnaUtil.Application without naming the interop type (the CS0433
    // duplicate-type hazard documented on IExcelApplicationAdapter) — the
    // same shape ExcelGanttRowInserter uses.
    private readonly Application? _application = application as Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    // R5.4 D12: the native Excel outline regroups with the mutation, not on
    // the next Refresh. Optional so a caller with no outline writer — and a
    // test that is not about grouping — does not have to construct one.
    private readonly IOutlineGroupPort? _outlineGroups = outlineGroups;

    /// <inheritdoc />
    public ParentChildMutationOutcome Promote() => Mutate(MutationAction.Promote);

    /// <inheritdoc />
    public ParentChildMutationOutcome Demote() => Mutate(MutationAction.Demote);

    /// <inheritdoc />
    public ParentChildMutationOutcome MoveUp() => Mutate(MutationAction.MoveUp);

    /// <inheritdoc />
    public ParentChildMutationOutcome MoveDown() => Mutate(MutationAction.MoveDown);

    /// <inheritdoc />
    public ParentChildMutationOutcome Remove() => Mutate(MutationAction.Remove);

    /// <summary>The mutation one <see cref="Mutate"/> run performs.</summary>
    private enum MutationAction
    {
        Promote = 0,
        Demote = 1,
        MoveUp = 2,
        MoveDown = 3,
        Remove = 4,
    }

    /// <summary>
    /// The shared pipeline: guard, locate, resolve the selection, read,
    /// plan, then apply. Every refusal happens before the first COM
    /// write, and every local (in-memory) mutation happens before the
    /// first COM write, so a failure anywhere leaves the workbook
    /// untouched.
    /// </summary>
    /// <param name="action">
    /// The promotion, demotion, or move the run performs.
    /// </param>
    /// <returns>The typed mutation outcome.</returns>
    private ParentChildMutationOutcome Mutate(MutationAction action)
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
        ParentChildMutationOutcome outcome = action switch
        {
            MutationAction.Promote => ApplyPromote(table, columnMap, body, events, selectedId, bodyIndex),
            MutationAction.Demote => ApplyDemote(table, columnMap, body, events, selectedId, bodyIndex),
            MutationAction.MoveUp or MutationAction.MoveDown => ApplyMove(
                table,
                worksheet,
                FirstBodyWorksheetRow(table),
                columnMap,
                body,
                events,
                selectedId,
                bodyIndex,
                moveUp: action == MutationAction.MoveUp),
            MutationAction.Remove => ApplyRemove(table, columnMap, body, events, selectedId, bodyIndex),
            _ => throw new ArgumentOutOfRangeException(nameof(action), action, "Unrecognised mutation action."),
        };

        return Finish(table, outcome);
    }

    /// <summary>
    /// Rebuilds the native Excel outline from the post-mutation table, so
    /// the <c>-</c>/<c>+</c> controls reflect the new hierarchy immediately
    /// rather than waiting for the next Refresh (R5.4 D12). Runs only on a
    /// successful mutation: refusals and boundary no-ops leave the outline
    /// untouched. An outline refusal degrades to the pre-D12 behaviour — the
    /// mutation stands and the next Refresh rebuilds the outline — rather
    /// than failing work that already landed.
    /// </summary>
    /// <param name="table">The mutated Gantt data table.</param>
    /// <param name="outcome">The mutation outcome to return.</param>
    /// <returns>The mutation outcome, unchanged.</returns>
    private ParentChildMutationOutcome Finish(
        ListObject table,
        ParentChildMutationOutcome outcome)
    {
        if (!outcome.Succeeded || outcome.CellsWritten == 0 || _outlineGroups is null)
        {
            return outcome;
        }

        if (!TryReadBody(table, out List<object?[]>? body, out var columnMap)
            || body is null || columnMap is null)
        {
            return outcome;
        }

        List<GanttEvent>? events = ProjectRows(
            body,
            columnMap,
            0,
            out ParentChildMutationRefusalReason? _);
        if (events is null)
        {
            return outcome;
        }

        _ = _outlineGroups.Apply(events);
        return outcome;
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
                if (ExcelGanttRowProjection.CellText(body[rowIndex], columnMap, "Id") == rehomedId.Value)
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

            if (plan.RehomedChildIds.Any(rehomedId => ExcelGanttRowProjection.CellText(body[index], columnMap, "Id") == rehomedId.Value))
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
    /// Applies a move plan (R5.4 D5): the selection's window and the
    /// neighbour's window exchange content in memory, both are rewritten
    /// full width (a content rotation — R5.4 D11-amended), and the
    /// resequenced <c>SiblingOrder</c> cells outside the swapped span
    /// are written one by one. A boundary plan (R5.4 D4) succeeds with
    /// zero writes and no COM call beyond the pipeline's reads.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="worksheet">The worksheet holding the table.</param>
    /// <param name="firstWorksheetRow">
    /// The worksheet row the table's first body row occupies, or <c>0</c> when
    /// it cannot be resolved.
    /// </param>
    /// <param name="columnMap">Schema-column to table-column index map.</param>
    /// <param name="body">The raw body matrix, in body order. Mutated in place.</param>
    /// <param name="events">The projected planner events, in body order.</param>
    /// <param name="selectedId">The identifier of the row being moved.</param>
    /// <param name="bodyIndex">The one-based body index of the selection.</param>
    /// <param name="moveUp">The move direction.</param>
    /// <returns>
    /// The typed mutation outcome, whose
    /// <see cref="ParentChildMutationOutcome.MovedToRowNumber"/> names
    /// the moved row's new body index on a successful swap (R5.4 D9).
    /// </returns>
    private ParentChildMutationOutcome ApplyMove(
        ListObject table,
        Worksheet worksheet,
        int firstWorksheetRow,
        int[] columnMap,
        List<object?[]> body,
        IReadOnlyList<GanttEvent> events,
        GanttRowId selectedId,
        int bodyIndex,
        bool moveUp)
    {
        MovePlan plan = CreateMovePlan(events, selectedId, moveUp);
        if (!plan.Succeeded)
        {
            return ParentChildMutationOutcome.PlanRefused(plan.Refusal!.Value, bodyIndex);
        }

        if (!plan.MovesAnything)
        {
            // R5.4 D4: the top or bottom of the set is a quiet success —
            // no dialog and no write.
            return ParentChildMutationOutcome.Ok(0, bodyIndex);
        }

        // R5.4 D9: the selection follows the moved row. The
        // selected row is the first row of its window -- a
        // group's own row or a child on its own. Move-up: the
        // selection's window is the lower one and its first row
        // lands on the neighbour's former span start. Move-down:
        // the neighbour's whole window moved up into the
        // selection's former span, so the first moved row lands
        // BELOW the neighbour's window. The move-down index is
        // the general form (the windows may be separated by the
        // gap rows of a malformed table, which stay between
        // them); it corrects the D9 formula SwapStart + 1, which
        // named the neighbour's last row rather than the moved
        // row whenever the two windows differ in length.
        var movedToBodyIndex = plan.WindowStart < plan.SwapStart
            ? plan.SwapStart + plan.SwapLength - plan.WindowLength + 1
            : plan.SwapStart + 1;

        // Guards against a crafted plan that disagrees with the table:
        // both windows must be non-empty, in range, and disjoint. The
        // real planners cannot emit such a plan; refusing is safer than
        // rotating the wrong window (the same discipline as the promote
        // guard above).
        var swapEnd = plan.SwapStart + plan.SwapLength;
        var windowEnd = plan.WindowStart + plan.WindowLength;
        if (plan.WindowLength <= 0
            || plan.SwapLength <= 0
            || plan.WindowStart < 0
            || plan.SwapStart < 0
            || windowEnd > body.Count
            || swapEnd > body.Count
            || (plan.WindowStart < swapEnd && plan.SwapStart < windowEnd))
        {
            return ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.TableInvalid);
        }

        // The content swap, applied to the in-memory body before any COM
        // write. The union of the two windows is contiguous, so it is
        // written as one span exactly like the promote window; any gap
        // rows of a malformed table between the windows are written back
        // with their own unchanged content.
        BlockSwap(body, plan.WindowStart, plan.WindowLength, plan.SwapStart, plan.SwapLength);
        var unionStart = Math.Min(plan.WindowStart, plan.SwapStart);
        var unionEnd = Math.Max(windowEnd, swapEnd) - 1;

        // R5.4 D13: the row-hidden flags rotate with the content. `Hidden`
        // is a property of the worksheet ROW, not of the content in it, and
        // a move is a content rotation -- so without this the parent of a
        // collapsed group lands on a row the host still reports hidden. The
        // Refresh that follows then measures that row at zero height,
        // `PanelCellGrid` accepts the zero deliberately (ADR-0034), and
        // `LaneRowAnchorResolver` refuses with `DegenerateRowHeight`, so the
        // whole Refresh fails (measured: scripts/probe-hidden-row-move.ps1).
        //
        // The flags go through the SAME `BlockSwap` permutation as the
        // content, so a parent stays visible and a collapsed child stays
        // hidden wherever the group lands. Unhiding the span outright would
        // expand the user's collapsed presentation, which `IOutlineGroupPort`
        // deliberately never does -- and it is not this adapter's state to
        // expand.
        //
        // Only the union span's flags are rotated: the gap rows of a
        // malformed table keep their positions, and every row outside the
        // union is untouched.
        List<bool> hiddenFlags = ReadHiddenFlags(worksheet, firstWorksheetRow, unionStart, unionEnd);
        BlockSwap(hiddenFlags, plan.WindowStart - unionStart, plan.WindowLength, plan.SwapStart - unionStart, plan.SwapLength);
        ParentChildMutationOutcome hidden = WriteHiddenFlags(worksheet, firstWorksheetRow, hiddenFlags, unionStart);
        if (!hidden.Succeeded)
        {
            return hidden;
        }

        // The resequenced SiblingOrder writes are keyed by row identifier
        // and resolve against the POST-swap positions, exactly as on
        // promote (R5.3 D4: the planner is the authority). The full-width
        // union write carries every write inside the union, so only the
        // cells beyond it are single-cell writes.
        Dictionary<GanttRowId, int> positions = IndexById(body, columnMap);
        List<(int RowIndex, string ColumnName, object? Value)> singles = [];
        foreach (SiblingOrderWrite write in plan.SiblingOrderWrites)
        {
            if (!positions.TryGetValue(write.RowId, out var position))
            {
                continue;
            }

            SetCell(body[position], columnMap, "SiblingOrder", write.SiblingOrder);
            if (position < unionStart || position > unionEnd)
            {
                // A row outside the written span keeps its row
                // position; only its SiblingOrder cell changes,
                // so it is written as a single cell.
                singles.Add((position, "SiblingOrder", write.SiblingOrder));
            }
        }

        // R5.4 D11 (amended 2026-10-12): the move is a content
        // rotation. The earlier worksheet-row insert+delete
        // relocation physically moved rows but host-deleted the
        // moved window's shapes (the lower window pays every
        // time), and a probe-measured whole-row Cut is a
        // destructive overwrite, not a swap — neither could keep
        // both windows' shapes. The union of the two windows is
        // contiguous, so it is written as one span; gap rows of a
        // malformed table (D5) are written back with their own
        // unchanged content. Shapes now change only when the
        // command's own Refresh rebuilds the chart.
        ParentChildMutationOutcome window = WriteRows(table, body, unionStart, unionEnd, bodyIndex);
        if (!window.Succeeded)
        {
            return window;
        }

        if (singles.Count == 0)
        {
            return window with { MovedToRowNumber = movedToBodyIndex };
        }

        ParentChildMutationOutcome outside = WriteCells(table, columnMap, singles, bodyIndex);
        return outside.Succeeded
            ? ParentChildMutationOutcome.Ok(window.CellsWritten + outside.CellsWritten, bodyIndex, movedToBodyIndex)
            : outside;
    }

    /// <summary>
    /// Computes the move plan. An internal seam so contract tests can
    /// drive the apply path with a crafted plan.
    /// </summary>
    /// <param name="rows">The projected events, in body order.</param>
    /// <param name="selectedId">The identifier of the row being moved.</param>
    /// <param name="moveUp">The move direction.</param>
    /// <returns>The move plan, or a typed refusal.</returns>
    internal virtual MovePlan CreateMovePlan(IReadOnlyList<GanttEvent> rows, GanttRowId selectedId, bool moveUp) =>
        HierarchyMutations.PlanMove(rows, selectedId, moveUp);

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
        List<GanttEvent>? events = ExcelGanttRowProjection.ProjectRows(body, columnMap, out var failedIndex);
        if (events is null)
        {
            refusal = failedIndex == selectedIndex
                ? ParentChildMutationRefusalReason.SelectionUnknown
                : ParentChildMutationRefusalReason.TableInvalid;
            return null;
        }

        return events;
    }

    /// <summary>
    /// Applies a remove plan (R5.4 D6): the selected row is
    /// deleted, its children are re-homed to its former parent
    /// (or top-level when it had none), and the sibling run it
    /// left is resequenced densely. Only the re-homed children's
    /// ParentId cells and the run's resequenced SiblingOrder
    /// cells are written; the delete then removes the whole
    /// worksheet row, so every row below -- written cells
    /// included -- shifts up with the shapes anchored to it
    /// (R5.4 D10, the owner's 2026-10-12 ruling; the shape
    /// anchored in the deleted row is removed by the host
    /// itself). The command performs no shape operation; the
    /// host moves the anchored shapes as its own row-delete
    /// side effect.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="columnMap">Schema-column to table-column index map.</param>
    /// <param name="body">The raw body matrix, in body order. Not mutated: the delete removes the selected row and shifts the survivors up.</param>
    /// <param name="events">The projected planner events, in body order.</param>
    /// <param name="selectedId">The identifier of the row being removed.</param>
    /// <param name="bodyIndex">The one-based body index of the selection.</param>
    /// <returns>The typed mutation outcome.</returns>
    private ParentChildMutationOutcome ApplyRemove(
        ListObject table,
        int[] columnMap,
        List<object?[]> body,
        IReadOnlyList<GanttEvent> events,
        GanttRowId selectedId,
        int bodyIndex)
    {
        // R5.4 D6: the planner is the authority for the surviving
        // hierarchy: which children move up, and to which parent.
        HierarchyPromotionPlan plan = HierarchyMutations.PlanPromotion(events, selectedId);

        // The run the deleted row belonged to is named by its former
        // parent: every survivor that shares that parent, plus the
        // promoted children once they inherit it, resequences densely.
        // A blank or unparseable ParentId cell is the top-level run,
        // the same convention the authoritative row reader applies.
        GanttRowId? runParent = plan.InheritedParentId;
        var fromIndex = bodyIndex - 1;
        var singles = new List<(int RowIndex, string ColumnName, object? Value)>();

        // The scan runs against the pre-delete body so every write
        // addresses the row's current position; the delete then
        // removes the worksheet row and carries each written cell
        // up with its row. The selected row itself is skipped: it
        // leaves with the delete. Existing SiblingOrder values are
        // ignored: the planner is the authority for the dense
        // resequence, exactly as the promote and demote planners
        // emit their runs' writes (R5.3 D4).
        var order = 0;
        for (var index = 0; index < body.Count; index++)
        {
            if (index == fromIndex)
            {
                continue;
            }

            GanttRowId? parentId = ParentIdOf(body[index], columnMap);
            if (parentId == selectedId)
            {
                // A child of the deleted row: it inherits the
                // deleted row's parent and takes the deleted
                // row's slot in the run, so its ParentId cell is
                // rewritten and its order is its dense position.
                singles.Add((index, "ParentId", plan.InheritedParentId?.Value));
                singles.Add((index, "SiblingOrder", order));
                order++;
            }
            else if (parentId == runParent)
            {
                // A surviving member of the run: it keeps its
                // place and renumbers densely.
                singles.Add((index, "SiblingOrder", order));
                order++;
            }
        }

        // The cell writes land on the rows' current positions.
        ParentChildMutationOutcome written = WriteCells(table, columnMap, singles, bodyIndex);
        if (!written.Succeeded)
        {
            return written;
        }

        // The delete goes last: a refused delete leaves the re-homed
        // children and the resequenced run in place -- still a valid
        // table -- where a delete-first failure would instead leave
        // children pointing at the deleted row's identifier.
        try
        {
            DeleteListRow(GetListRows(table), bodyIndex);
        }
        catch (COMException)
        {
            return ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.WriteFailed);
        }

        return ParentChildMutationOutcome.Ok(written.CellsWritten, bodyIndex);
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
                var tableIndex = columnMap[ExcelGanttRowProjection.SchemaIndexOf(columnName)];
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
    /// Swaps two disjoint contiguous blocks of a list in place. Rows in any
    /// gap between the blocks keep their positions.
    /// </summary>
    /// <typeparam name="T">The element type.</typeparam>
    /// <param name="body">The list to permute.</param>
    /// <param name="firstStart">The zero-based start of the first block.</param>
    /// <param name="firstLength">The length of the first block.</param>
    /// <param name="secondStart">The zero-based start of the second block.</param>
    /// <param name="secondLength">The length of the second block.</param>
    /// <remarks>
    /// Generic so the row-hidden flags can be rotated through <em>exactly</em>
    /// the permutation the content goes through, from one shared authority
    /// rather than a second hand-written copy that could drift (R5.4 D13).
    /// </remarks>
    private static void BlockSwap<T>(
        List<T> body,
        int firstStart,
        int firstLength,
        int secondStart,
        int secondLength)
    {
        var lowStart = Math.Min(firstStart, secondStart);
        var highStart = Math.Max(firstStart, secondStart);
        var lowLength = firstStart < secondStart ? firstLength : secondLength;
        var highLength = firstStart < secondStart ? secondLength : firstLength;

        List<T> span = body.GetRange(lowStart, highStart - lowStart + highLength);
        List<T> low = span.GetRange(0, lowLength);
        var secondOffset = highStart - lowStart;
        List<T> gap = span.GetRange(lowLength, secondOffset - lowLength);
        List<T> high = span.GetRange(secondOffset, highLength);

        var writeIndex = lowStart;
        foreach (T row in high)
        {
            body[writeIndex++] = row;
        }

        foreach (T row in gap)
        {
            body[writeIndex++] = row;
        }

        foreach (T row in low)
        {
            body[writeIndex++] = row;
        }
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
        object? value) => row[columnMap[ExcelGanttRowProjection.SchemaIndexOf(columnName)] - 1] = value;

    /// <summary>
    /// Reads a row's <c>ParentId</c> as a parsed identifier. A
    /// blank or unparseable cell is the top-level (null) parent --
    /// the same convention the authoritative row reader applies, so
    /// a top-level run is matched by its blank cells.
    /// </summary>
    /// <param name="row">The raw body row.</param>
    /// <param name="columnMap">Schema-column to table-column index map.</param>
    /// <returns>The parsed parent, or <see langword="null"/> for top-level.</returns>
    private static GanttRowId? ParentIdOf(object?[] row, int[] columnMap) =>
        GanttRowId.TryParse(ExcelGanttRowProjection.CellText(row, columnMap, "ParentId"), out GanttRowId? parent)
            ? parent
            : null;

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
            if (GanttRowId.TryParse(ExcelGanttRowProjection.CellText(body[index], columnMap, "Id"), out GanttRowId? id)
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
    /// Deletes the body row's whole worksheet row at the
    /// one-based index, so the table shrinks and every row
    /// below shifts up with its content and its anchored
    /// shapes (R5.4 D10). Test seam over the COM
    /// <c>ListRows</c> indexer, the <c>ListRow.Range</c> and
    /// <c>Range.EntireRow</c> properties and the
    /// <c>Range.Delete</c> method. The probes show
    /// <c>ListRow.Delete</c> alone only compacts the table in
    /// place -- the worksheet row is not deleted, no row below
    /// shifts and no anchored shape moves
    /// (scripts/probe-delete-row.ps1) -- while a true
    /// worksheet-row delete inside the ListObject range is
    /// accepted by the host, the symmetric counterpart of
    /// the insert ADR-0036 Q2 measured (both shift the
    /// sheet inside the ListObject range; the "refused
    /// insert" was ADR-0035's claim, withdrawn by ADR-0036
    /// itself), and is the row delete the owner's
    /// 2026-10-12 ruling requires
    /// (scripts/probe-delete-row-shift.ps1).
    /// </summary>
    /// <param name="rows">The table's body rows.</param>
    /// <param name="index">The one-based index of the row to delete.</param>
    internal virtual void DeleteListRow(ListRows rows, int index)
    {
        ListRow row = rows[index];
        Excel.Range rowRange = row.Range;
        Excel.Range entireRow = rowRange.EntireRow;
        entireRow.Delete();
    }

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
    /// Resolves the worksheet row the table's first body row occupies, so a
    /// body index can be addressed as a worksheet row (R5.4 D13). Test seam
    /// over <see cref="GetDataBodyRange"/> and <see cref="GetRangeRow"/>.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <returns>
    /// The one-based worksheet row of the first body row, or <c>0</c> when the
    /// body cannot be resolved. Zero addresses no row, so the caller's hidden
    /// writes are skipped rather than aimed at a guessed row.
    /// </returns>
    internal int FirstBodyWorksheetRow(ListObject table)
    {
        Excel.Range? body = GetDataBodyRange(table);
        return body is null ? 0 : GetRangeRow(body);
    }

    /// <summary>
    /// Reads the hidden flag of each body row in a contiguous span, in body
    /// order (R5.4 D13).
    /// </summary>
    /// <param name="worksheet">The worksheet holding the table.</param>
    /// <param name="firstWorksheetRow">
    /// The worksheet row of the first body row, or <c>0</c> when unresolved.
    /// </param>
    /// <param name="fromIndex">The zero-based body index the span starts at.</param>
    /// <param name="toIndex">The zero-based body index the span ends at.</param>
    /// <returns>
    /// One flag per span row, all <see langword="false"/> when the body row is
    /// unresolved. An unreadable row reads as visible: the conservative
    /// reading, which never leaves a row hidden by omission.
    /// </returns>
    internal List<bool> ReadHiddenFlags(Worksheet worksheet, int firstWorksheetRow, int fromIndex, int toIndex)
    {
        List<bool> flags = new(toIndex - fromIndex + 1);
        for (var index = fromIndex; index <= toIndex; index++)
        {
            Excel.Range? row = firstWorksheetRow > 0
                ? GetWorksheetRow(worksheet, firstWorksheetRow + index)
                : null;
            flags.Add(row is not null && ReadHiddenFlag(row));
        }

        return flags;
    }

    /// <summary>
    /// Writes the hidden flags back to their body rows, skipping any row
    /// whose flag already matches (R5.4 D13).
    /// </summary>
    /// <param name="worksheet">The worksheet holding the table.</param>
    /// <param name="firstWorksheetRow">
    /// The worksheet row of the first body row, or <c>0</c> when unresolved.
    /// </param>
    /// <param name="flags">The flags in body order, starting at <paramref name="fromIndex"/>.</param>
    /// <param name="fromIndex">The zero-based body index the span starts at.</param>
    /// <returns>The typed mutation outcome.</returns>
    /// <remarks>
    /// Only changed rows are written: a no-op rotation leaves every flag
    /// alone, so a collapsed group is never re-evaluated by an unnecessary
    /// write. A <see cref="ParentChildMutationRefusalReason.WriteFailed"/>
    /// reports a partially applied rotation, the same discipline
    /// <see cref="WriteRows"/> applies to its own host refusal.
    /// </remarks>
    internal ParentChildMutationOutcome WriteHiddenFlags(
        Worksheet worksheet,
        int firstWorksheetRow,
        List<bool> flags,
        int fromIndex)
    {
        if (firstWorksheetRow <= 0)
        {
            // The body row could not be resolved, so no row is addressed.
            // Refused rather than skipped: the rotation's content landed but
            // its presentation state did not, and that disagreement is the
            // defect this exists to remove.
            return ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.WriteFailed);
        }

        var written = 0;
        try
        {
            for (var offset = 0; offset < flags.Count; offset++)
            {
                var wanted = flags[offset];
                Excel.Range? row = GetWorksheetRow(worksheet, firstWorksheetRow + fromIndex + offset);
                if (row is null || ReadHiddenFlag(row) == wanted)
                {
                    continue;
                }

                SetRowHidden(row, wanted);
                written++;
            }
        }
        catch (COMException)
        {
            return ParentChildMutationOutcome.Refused(ParentChildMutationRefusalReason.WriteFailed);
        }

        return ParentChildMutationOutcome.Ok(written, fromIndex + 1);
    }

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

    /// <summary>
    /// Reads <c>Range.Hidden</c> as a real <see cref="bool"/>. Test seam over
    /// the COM <c>Range.Hidden</c> property.
    /// </summary>
    /// <param name="range">The row range.</param>
    /// <returns>The unwrapped visibility flag.</returns>
    /// <remarks>
    /// The PIA declares <c>Range.Hidden</c> as <see cref="object"/>, so it is
    /// unboxed for the same reason <c>ExcelColumnPresentationRestorer</c>
    /// unboxes it: comparing a boxed value against a <see cref="bool"/> would
    /// never be equal, so every row would be rewritten on every move. An
    /// unrecognised host value reads as <see langword="false"/> — visible, the
    /// conservative reading that never strands a row hidden by omission.
    /// </remarks>
    internal virtual bool ReadHiddenFlag(Excel.Range range) =>
        range.Hidden switch
        {
            bool value => value,
            int value => value != 0,
            double value => value != 0,
            _ => false,
        };

    /// <summary>
    /// Returns the whole worksheet row at the one-based index. Test seam over
    /// the COM <c>Worksheet.Rows</c> property and its indexer.
    /// </summary>
    /// <param name="worksheet">The worksheet holding the table.</param>
    /// <param name="rowIndex">The one-based worksheet row index.</param>
    /// <returns>The row range, or <see langword="null"/> when unreachable.</returns>
    internal virtual Excel.Range? GetWorksheetRow(Worksheet worksheet, int rowIndex) =>
        worksheet.Rows is { } rows ? rows[rowIndex] : null;

    /// <summary>
    /// Sets a row's hidden flag through the row's <c>EntireRow</c>. Test seam
    /// over the COM <c>Range.EntireRow</c> property and the
    /// <c>Range.Hidden</c> setter.
    /// </summary>
    /// <param name="row">The row range.</param>
    /// <param name="hidden">Whether the row is hidden.</param>
    /// <remarks>
    /// <b>The accessor is load-bearing.</b> The probe behind this decision
    /// measured <c>Range("A3:C3").Hidden = $true</c> as <b>refused</b> with
    /// COMException <c>0x800A03EC</c> while the identically shaped
    /// <c>Range("A3:C3").EntireRow.Hidden = $true</c> was accepted — the same
    /// accessor asymmetry ADR-0034 recorded for <c>OutlineLevel</c>. Writing
    /// through a plain table-sourced range would fail every hidden-flag write
    /// on a live host while passing every fake-based test.
    /// </remarks>
    internal virtual void SetRowHidden(Excel.Range row, bool hidden)
    {
        // Hold the EntireRow proxy in a local before setting
        // Hidden, the same discipline DeleteListRow applies to
        // its row delete: no chained COM expression (J).
        Excel.Range entireRow = row.EntireRow;
        entireRow.Hidden = hidden;
    }
}
