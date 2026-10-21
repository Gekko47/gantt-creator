using GanttCreator.Core;

namespace GanttCreator.Office;

/// <summary>
/// Why an add-child was refused before any cell was written.
/// </summary>
public enum AddChildMutationRefusalReason
{
    /// <summary>The application object or active workbook was absent.</summary>
    NoActiveWorkbook = 0,

    /// <summary>The visible <c>tblGanttData</c> table was not found.</summary>
    TableMissing = 1,

    /// <summary>The target worksheet or workbook is protected.</summary>
    TargetProtected = 2,

    /// <summary>The active cell is outside the table body.</summary>
    SelectionOutsideTable = 3,

    /// <summary>The table body contains a row whose identifier or type cell cannot be parsed.</summary>
    TableInvalid = 4,

    /// <summary>The host refused the row insert.</summary>
    InsertRefused = 5,

    /// <summary>The host refused a planned cell write.</summary>
    WriteFailed = 6,
}

/// <summary>
/// The typed result of one add-child mutation. Every refusal before the
/// first write leaves the workbook untouched.
/// </summary>
/// <param name="Succeeded">Whether the mutation completed.</param>
/// <param name="Refusal">The infrastructure refusal reason, or null when the plan succeeded.</param>
/// <param name="PlanRefusal">The Core planner's refusal reason, or null when the plan succeeded.</param>
/// <param name="SelectedRowNumber">The one-based body index of the parent row, when the selection resolved to one.</param>
/// <param name="InsertedBodyIndex">The one-based body index the new row occupies, when it was inserted.</param>
/// <param name="CellsWritten">How many cells the mutation wrote.</param>
public sealed record AddChildMutationOutcome(
    bool Succeeded,
    AddChildMutationRefusalReason? Refusal,
    AddChildRefusalReason? PlanRefusal,
    int? SelectedRowNumber,
    int? InsertedBodyIndex,
    int CellsWritten)
{
    /// <summary>Creates a successful outcome.</summary>
    /// <param name="cellsWritten">How many cells the mutation wrote.</param>
    /// <param name="selectedRowNumber">The parent row's one-based body index.</param>
    /// <param name="insertedBodyIndex">The new row's one-based body index.</param>
    /// <returns>The successful outcome.</returns>
    public static AddChildMutationOutcome Ok(
        int cellsWritten,
        int selectedRowNumber,
        int insertedBodyIndex) =>
        new(true, null, null, selectedRowNumber, insertedBodyIndex, cellsWritten);

    /// <summary>Creates an infrastructure refusal. Nothing was mutated.</summary>
    /// <param name="refusal">Why no mutation occurred.</param>
    /// <returns>The refusal outcome.</returns>
    public static AddChildMutationOutcome Refused(AddChildMutationRefusalReason refusal) =>
        new(false, refusal, null, null, null, 0);

    /// <summary>Creates a planner refusal. Nothing was mutated.</summary>
    /// <param name="planRefusal">The Core planner's refusal reason.</param>
    /// <param name="selectedRowNumber">The parent row's one-based body index.</param>
    /// <returns>The refusal outcome.</returns>
    public static AddChildMutationOutcome PlanRefused(
        AddChildRefusalReason planRefusal,
        int selectedRowNumber) =>
        new(false, null, planRefusal, selectedRowNumber, null, 0);
}

/// <summary>
/// The narrow port over the R5.3 WI-E add-child adapter. The surface is
/// primitive — no interop type crosses it — so the AddIn never names
/// <c>Microsoft.Office.Interop.Excel</c>.
/// </summary>
/// <remarks>
/// <para>
/// The port is mutating. The method resolves the active cell to the PARENT
/// row, reads the table, asks the Core planner whether the requested type
/// may be added under it, inserts the new row, and writes the new row's
/// <c>ParentId</c> and <c>SiblingOrder</c> cells. The protection guard
/// runs first (ADR-0008 D4) and every refusal mutates nothing.
/// </para>
/// <para>
/// The selection is the input, and resolving it is Office work (the
/// active-cell seam), so it belongs in the adapter rather than in the thin
/// ribbon callback.
/// </para>
/// </remarks>
public interface IAddChildMutator
{
    /// <summary>
    /// Inserts one new row as a CHILD of the row under the active cell.
    /// </summary>
    /// <param name="childType">The type the new row will have.</param>
    /// <param name="nextId">The one-shot stable row-ID generator.</param>
    /// <returns>The typed mutation outcome.</returns>
    AddChildMutationOutcome AddChild(GanttEntityType childType, Func<GanttRowId> nextId);
}
