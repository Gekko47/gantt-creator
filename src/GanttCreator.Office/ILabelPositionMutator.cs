using GanttCreator.Core;

namespace GanttCreator.Office;

/// <summary>
/// Why a label-position mutation was refused before any cell was written.
/// Planner refusals name the catalogue rule; infrastructure refusals name the
/// host fact. Every refusal leaves the workbook untouched.
/// </summary>
public enum LabelPositionMutationRefusalReason
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
    /// The table body contains a row whose identifier or type cell cannot be
    /// parsed, so no plan can be computed.
    /// </summary>
    TableInvalid = 5,

    /// <summary>The selected row's type is unknown.</summary>
    TypeUnknown = 6,

    /// <summary>The position is not permitted for the row's type (or style).</summary>
    PositionNotAllowedForType = 7,

    /// <summary>The host refused the planned cell write.</summary>
    WriteFailed = 8,
}

/// <summary>
/// The typed result of one label-position mutation. A success writes exactly
/// one <c>LabelPosition</c> cell (blank for Default); every refusal writes
/// nothing.
/// </summary>
/// <param name="Succeeded">Whether the mutation completed.</param>
/// <param name="Refusal">The refusal reason, or null on success.</param>
/// <param name="SelectedRowNumber">The one-based body index targeted, when known.</param>
/// <param name="CellsWritten">How many cells the mutation wrote (0 or 1).</param>
public sealed record LabelPositionMutationOutcome(
    bool Succeeded,
    LabelPositionMutationRefusalReason? Refusal,
    int? SelectedRowNumber,
    int CellsWritten)
{
    /// <summary>Creates a successful outcome.</summary>
    public static LabelPositionMutationOutcome Ok(int cellsWritten, int selectedRowNumber) =>
        new(true, null, selectedRowNumber, cellsWritten);

    /// <summary>Creates a refusal. Nothing was mutated.</summary>
    public static LabelPositionMutationOutcome Refused(
        LabelPositionMutationRefusalReason refusal, int? selectedRowNumber = null) =>
        new(false, refusal, selectedRowNumber, 0);
}

/// <summary>
/// Sets one row's <c>LabelPosition</c> cell (R5.5 D2: table write, never shape
/// edit; bespoke per entity, never group-applied).
/// </summary>
/// <remarks>
/// The port mirrors <see cref="IParentChildMutator"/>: the selection is the
/// input and resolving it is Office work, so the parameterless method reads
/// the active cell while the indexed overload is the testable seam. The
/// protection guard runs first (ADR-0008 D4) and every refusal mutates
/// nothing. Refresh applies the scene effect; this port never renders.
/// </remarks>
public interface ILabelPositionMutator
{
    /// <summary>
    /// Sets the <c>LabelPosition</c> cell of the row under the active cell.
    /// A null position clears the cell (blank = Default per §22).
    /// </summary>
    /// <param name="position">The position to write, or null to clear.</param>
    /// <returns>The typed mutation outcome.</returns>
    LabelPositionMutationOutcome SetLabelPosition(GanttLabelPosition? position);
}
