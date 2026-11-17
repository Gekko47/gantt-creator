namespace GanttCreator.Office;

/// <summary>
/// Why a style-key mutation was refused before any cell was written.
/// Planner refusals name the catalogue rule; infrastructure refusals name the
/// host fact. Every refusal leaves the workbook untouched.
/// </summary>
public enum StyleKeyMutationRefusalReason
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

    /// <summary>The selected row's type does not carry a StyleKey (not a style-key type).</summary>
    TypeNotCustomStyled = 7,

    /// <summary>The requested style key is blank or the registry does not know it.</summary>
    StyleUnknown = 8,

    /// <summary>The host refused the planned cell write.</summary>
    WriteFailed = 9,
}

/// <summary>
/// The typed result of one style-key mutation. A success writes exactly one
/// <c>StyleKey</c> cell (R5.6a: bespoke per entity, never group-applied);
/// every refusal writes nothing.
/// </summary>
/// <param name="Succeeded">Whether the mutation completed.</param>
/// <param name="Refusal">The refusal reason, or null on success.</param>
/// <param name="SelectedRowNumber">The one-based body index targeted, when known.</param>
/// <param name="CellsWritten">How many cells the mutation wrote (0 or 1).</param>
public sealed record StyleKeyMutationOutcome(
    bool Succeeded,
    StyleKeyMutationRefusalReason? Refusal,
    int? SelectedRowNumber,
    int CellsWritten)
{
    /// <summary>Creates a successful outcome.</summary>
    public static StyleKeyMutationOutcome Ok(int cellsWritten, int selectedRowNumber) =>
        new(true, null, selectedRowNumber, cellsWritten);

    /// <summary>Creates a refusal. Nothing was mutated.</summary>
    public static StyleKeyMutationOutcome Refused(
        StyleKeyMutationRefusalReason refusal, int? selectedRowNumber = null) =>
        new(false, refusal, selectedRowNumber, 0);
}

/// <summary>
/// Sets one row's <c>StyleKey</c> cell (R5.6a: table write, never shape edit;
/// bespoke per entity — ruling 1 — so the selection is the input and resolving
/// it is Office work).
/// </summary>
/// <remarks>
/// The port mirrors <see cref="ILabelPositionMutator"/>: the protection guard
/// runs first (ADR-0008 D4), the target row must be a style-key type
/// (Custom Activity / Custom Milestone), and the requested key must exist in
/// the configuration style registry. Refresh applies the scene effect; this
/// port never renders.
/// </remarks>
public interface IStyleKeyMutator
{
    /// <summary>
    /// Sets the <c>StyleKey</c> cell of the row under the active cell.
    /// </summary>
    /// <param name="styleKey">The exact style key to write.</param>
    /// <returns>The typed mutation outcome.</returns>
    StyleKeyMutationOutcome SetStyleKey(string styleKey);
}
