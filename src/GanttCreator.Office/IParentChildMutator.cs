using GanttCreator.Core;

namespace GanttCreator.Office;

/// <summary>
/// Why a parent/child mutation was refused before any cell was
/// written. Planner refusals are carried separately in
/// <see cref="ParentChildMutationOutcome.PlanRefusal"/> so the
/// translator can name the row and the specific rule.
/// </summary>
public enum ParentChildMutationRefusalReason
{
    /// <summary>The application object or active workbook was absent.</summary>
    NoActiveWorkbook = 0,

    /// <summary>The visible <c>tblGanttData</c> table was not found.</summary>
    TableMissing = 1,

    /// <summary>The target worksheet or workbook is protected.</summary>
    TargetProtected = 2,

    /// <summary>The active cell is outside the table body.</summary>
    SelectionOutsideTable = 3,

    /// <summary>The active cell does not name a table row with a valid identifier.</summary>
    SelectionUnknown = 4,

    /// <summary>
    /// The table body contains a row whose identifier or type cell
    /// cannot be parsed, so no hierarchy plan can be computed.
    /// </summary>
    TableInvalid = 5,

    /// <summary>The host refused a planned cell write.</summary>
    WriteFailed = 6,
}

/// <summary>
/// The typed result of one promote or demote mutation. Every refusal
/// leaves the workbook untouched.
/// </summary>
/// <param name="Succeeded">Whether the mutation completed.</param>
/// <param name="Refusal">
/// The infrastructure refusal reason, or <see langword="null"/> when the
/// pipeline reached the planner.
/// </param>
/// <param name="PlanRefusal">
/// The Core planner's refusal reason, or <see langword="null"/> when the
/// plan succeeded or the pipeline refused earlier.
/// </param>
/// <param name="SelectedRowNumber">
/// The one-based body index of the row the command targeted, when the
/// selection resolved to a row. This is the row number the translator
/// reports.
/// </param>
/// <param name="CellsWritten">How many cells the mutation wrote.</param>
/// <param name="MovedToRowNumber">
/// The one-based body index the targeted row occupies after a
/// successful move (R5.4 D9: the selection follows the moved
/// row), or <see langword="null"/> when the row did not move --
/// a boundary no-op, a promote, a demote, or a remove.
/// </param>
public sealed record ParentChildMutationOutcome(
    bool Succeeded,
    ParentChildMutationRefusalReason? Refusal,
    ParentChildRefusalReason? PlanRefusal,
    int? SelectedRowNumber,
    int CellsWritten,
    int? MovedToRowNumber = null)
{
    /// <summary>Creates a successful outcome.</summary>
    /// <param name="cellsWritten">How many cells the mutation wrote.</param>
    /// <param name="selectedRowNumber">The targeted row's one-based body index.</param>
    /// <param name="movedToRowNumber">The moved row's new one-based body index, when the row moved.</param>
    /// <returns>The successful outcome.</returns>
    public static ParentChildMutationOutcome Ok(
        int cellsWritten,
        int selectedRowNumber,
        int? movedToRowNumber = null) =>
        new(true, null, null, selectedRowNumber, cellsWritten, movedToRowNumber);

    /// <summary>Creates an infrastructure refusal. Nothing was mutated.</summary>
    /// <param name="refusal">Why no mutation occurred.</param>
    /// <returns>The refusal outcome.</returns>
    public static ParentChildMutationOutcome Refused(ParentChildMutationRefusalReason refusal) =>
        new(false, refusal, null, null, 0);

    /// <summary>Creates a planner refusal. Nothing was mutated.</summary>
    /// <param name="planRefusal">The Core planner's refusal reason.</param>
    /// <param name="selectedRowNumber">The targeted row's one-based body index.</param>
    /// <returns>The refusal outcome.</returns>
    public static ParentChildMutationOutcome PlanRefused(
        ParentChildRefusalReason planRefusal,
        int selectedRowNumber) =>
        new(false, null, planRefusal, selectedRowNumber, 0);
}

/// <summary>
/// The narrow port over the R5.3 parent/child mutation adapters. The
/// surface is deliberately primitive — no interop type crosses it — so
/// the AddIn compilation never names
/// <c>Microsoft.Office.Interop.Excel</c> (the CS0433 duplicate-type
/// hazard documented on <see cref="IExcelApplicationAdapter"/>).
/// </summary>
/// <remarks>
/// <para>
/// The port is mutating. Each method resolves the active cell to a
/// table body row, reads the table, asks the Core planner for the
/// change, and bulk-writes exactly the cells the plan names. The
/// protection guard runs first (ADR-0008 D4) and every refusal
/// mutates nothing.
/// </para>
/// <para>
/// <b>Why the outline regroups with the mutation (R5.4 D12).</b> Each
/// successful mutation also rebuilds the native Excel outline from the
/// post-mutation table, so the <c>-</c>/<c>+</c> controls reflect the
/// new hierarchy immediately rather than waiting for the next Refresh.
/// </para>
/// <para>
/// <b>Why parameterless methods.</b> The selection is the input, and
/// resolving it is Office work (the active-cell seam), so it belongs
/// in the adapter, not in the thin ribbon callback. This keeps the
/// AddIn command a pure outcome-to-message translation.
/// </para>
/// </remarks>
public interface IParentChildMutator
{
    /// <summary>
    /// Promotes the child row under the active cell to top-level
    /// (R5.3 D2): its content moves to the bottom of its sibling run,
    /// its <c>ParentId</c> is cleared, and the remaining run is
    /// resequenced densely.
    /// </summary>
    /// <returns>The typed mutation outcome.</returns>
    ParentChildMutationOutcome Promote();

    /// <summary>
    /// Demotes the top-level row under the active cell to a child of
    /// the nearest eligible parent above it (R5.3 D3): a
    /// <c>ParentId</c> rewrite plus the <c>SiblingOrder</c> tail
    /// assignment in the new run and the resequencing of the top-level
    /// run it leaves. No physical row move.
    /// </summary>
    /// <returns>The typed mutation outcome.</returns>
    ParentChildMutationOutcome Demote();

    /// <summary>
    /// Moves the group under the active cell up one position (R5.4
    /// D5): a top-level selection swaps its window (row plus child
    /// run) with the group above it, a child swaps with its previous
    /// sibling, and the top of the set is a silent success with zero
    /// cells written (R5.4 D4). Both affected windows are rewritten
    /// full width — a content rotation; the worksheet rows and their
    /// anchored shapes do not move (R5.4 D11-amended, the command's
    /// own Refresh rebuilds the chart). On success
    /// the outcome carries the moved row's new body index (R5.4 D9)
    /// so the caller can move the selection onto it.
    /// </summary>
    /// <returns>The typed mutation outcome.</returns>
    ParentChildMutationOutcome MoveUp();

    /// <summary>
    /// Moves the group under the active cell down one position (R5.4
    /// D5): the mirror of <see cref="MoveUp"/>, with the bottom of the
    /// set a silent success with zero cells written. On success the
    /// outcome carries the moved row's new body index (R5.4 D9) so
    /// the caller can move the selection onto it.
    /// </summary>
    /// <returns>The typed mutation outcome.</returns>
    ParentChildMutationOutcome MoveDown();

    /// <summary>
    /// Removes the row under the active cell (R5.4 D6): the row is
    /// deleted from <c>tblGanttData</c>, its children are re-homed to
    /// its parent (or top-level when it had none), and the sibling run
    /// it left is resequenced densely. Destructive under ADR-0008: the
    /// caller confirms before invoking, and No mutates nothing.
    /// </summary>
    /// <returns>The typed mutation outcome.</returns>
    ParentChildMutationOutcome Remove();
}
