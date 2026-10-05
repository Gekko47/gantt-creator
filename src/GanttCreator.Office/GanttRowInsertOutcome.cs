namespace GanttCreator.Office;

/// <summary>Typed result of one guarded row insertion.</summary>
public enum GanttRowInsertRefusalReason
{
    /// <summary>No active Excel workbook exists.</summary>
    NoActiveWorkbook = 0,

    /// <summary>The visible <c>tblGanttData</c> table was not found.</summary>
    TableMissing = 1,

    /// <summary>The target worksheet or workbook is protected.</summary>
    TargetProtected = 2,

    /// <summary>Type validation could not be prepared after the row was added.</summary>
    TypeOptionsUnavailable = 3,

    /// <summary>
    /// The worksheet row was inserted, but the host refused a later call that writes
    /// the row's content, its height, or the list row itself.
    /// </summary>
    /// <remarks>
    /// This is deliberately NOT one of the "nothing happened" refusals above. By the
    /// time this is reported the row is already in the worksheet, so reporting the
    /// insert as failed would be a lie the user can disprove by looking at the sheet.
    /// The message says the row was added and not completed, which is the only honest
    /// description of the state.
    /// </remarks>
    RowWriteRefused = 4,

    /// <summary>
    /// The host refused the worksheet-row insert on the POSITIONAL branch, so no row
    /// was added and nothing was written.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This is deliberately separate from <see cref="RowWriteRefused"/>, and the
    /// distinction is the whole point. On the positional branch the adapter reads the
    /// row back by index (<c>GetListRowAt</c>), so continuing past a refused insert
    /// would aim the scaffold write at the user's <em>existing</em> row at that
    /// position — silent data loss reported as success. Refusing first is the only
    /// honest outcome, and it means "nothing happened", which is what the other
    /// refusals above already say.
    /// </para>
    /// <para>
    /// The append branch never produces this: its <c>AddRow</c> creates a genuine new
    /// row even when the push-down below the table failed, so the row exists and the
    /// lost padding row is reported through <c>PaddingRowReserved</c> instead.
    /// </para>
    /// </remarks>
    RowInsertRefused = 5,
}

/// <summary>The typed result of a row insertion.</summary>
/// <param name="Succeeded">Whether a row was appended.</param>
/// <param name="BodyIndex">The new row's one-based body index when successful.</param>
/// <param name="Refusal">The refusal reason when unsuccessful.</param>
/// <param name="PaddingRowReserved">
/// Whether the chart's bottom padding row survived the insert. False means the table
/// absorbed it, which is a visible defect rather than a cosmetic one.
/// </param>
/// <param name="ReservedRowsNormalised">
/// Whether the reserved rows below the body were restored to their tokens after the
/// insert (ADR-0038 D1). False means they still carry whatever height they had, which
/// leaves the chart's bottom margin the wrong size.
/// </param>
public sealed record GanttRowInsertOutcome(
    bool Succeeded,
    int? BodyIndex,
    GanttRowInsertRefusalReason? Refusal,
    bool PaddingRowReserved = true,
    bool ReservedRowsNormalised = true)
{
    /// <summary>Creates a successful insertion outcome.</summary>
    /// <param name="bodyIndex">The new row's one-based body index.</param>
    /// <param name="paddingRowReserved">
    /// Whether the chart's bottom padding row survived the insert. False means the
    /// table absorbed it, which is a visible defect rather than a cosmetic one, so
    /// it is reported instead of being passed over.
    /// </param>
    /// <param name="reservedRowsNormalised">
    /// Whether the reserved rows below the body were restored to their tokens.
    /// </param>
    /// <returns>A successful outcome.</returns>
    public static GanttRowInsertOutcome Ok(
        int bodyIndex,
        bool paddingRowReserved = true,
        bool reservedRowsNormalised = true) =>
        new(true, bodyIndex, null, paddingRowReserved, reservedRowsNormalised);

    /// <summary>Creates a refusal outcome.</summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <returns>A refusal outcome.</returns>
    public static GanttRowInsertOutcome Refused(GanttRowInsertRefusalReason refusal) =>
        new(false, null, refusal);
}
