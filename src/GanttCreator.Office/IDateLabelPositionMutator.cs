using GanttCreator.Core;

namespace GanttCreator.Office;

/// <summary>
/// Why a date-label mutation was refused before any cell was written.
/// Planner refusals name the catalogue rule; infrastructure refusals name the
/// host fact. Every refusal leaves the workbook untouched.
/// </summary>
public enum DateLabelMutationRefusalReason
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

    /// <summary>The selected row's type is unknown.</summary>
    TypeUnknown = 5,

    /// <summary>The choice is not permitted for the row's type (R5.14 D4).</summary>
    DateLabelNotAllowedForType = 6,

    /// <summary>The host refused the planned cell write.</summary>
    WriteFailed = 7,
}

/// <summary>
/// The typed result of one date-label mutation. A success writes exactly
/// one <c>DateLabelPosition</c> cell (blank for Auto); every refusal
/// writes nothing.
/// </summary>
/// <param name="Succeeded">Whether the mutation completed.</param>
/// <param name="Refusal">The refusal reason, or null on success.</param>
/// <param name="SelectedRowNumber">The one-based body index targeted, when known.</param>
/// <param name="CellsWritten">How many cells the mutation wrote (0 or 1).</param>
public sealed record DateLabelMutationOutcome(
    bool Succeeded,
    DateLabelMutationRefusalReason? Refusal,
    int? SelectedRowNumber,
    int CellsWritten)
{
    /// <summary>Creates a successful outcome.</summary>
    public static DateLabelMutationOutcome Ok(int cellsWritten, int selectedRowNumber) =>
        new(true, null, selectedRowNumber, cellsWritten);

    /// <summary>Creates a refusal. Nothing was mutated.</summary>
    public static DateLabelMutationOutcome Refused(
        DateLabelMutationRefusalReason refusal, int? selectedRowNumber = null) =>
        new(false, refusal, selectedRowNumber, 0);
}

/// <summary>
/// Sets one row's <c>DateLabelPosition</c> cell (R5.14: table write,
/// never shape edit; bespoke per entity, never group-applied).
/// </summary>
/// <remarks>
/// The port mirrors <see cref="ILabelPositionMutator"/>: the selection is
/// the input and resolving it is Office work, so the parameterless method
/// reads the active cell while the adapter's internal apply path is the
/// testable seam. The protection guard runs first (ADR-0008 D4) and every
/// refusal mutates nothing. Refresh applies the scene effect; this port
/// never renders. The capability is type-only (R5.14 D4): a
/// <c>Custom Activity</c> style carries no date mode, so it cannot narrow
/// the set the way it narrows <see cref="GanttLabelPosition"/>.
/// </remarks>
public interface IDateLabelPositionMutator
{
    /// <summary>
    /// Sets the <c>DateLabelPosition</c> cell of the row under the active cell.
    /// A null choice clears the cell (blank = Auto per R5.14 D3).
    /// </summary>
    /// <param name="position">The choice to write, or null to clear.</param>
    /// <returns>The typed mutation outcome.</returns>
    DateLabelMutationOutcome SetDateLabelPosition(GanttDateLabelPosition? position);
}