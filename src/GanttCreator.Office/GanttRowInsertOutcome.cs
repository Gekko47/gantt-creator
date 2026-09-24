namespace GanttCreator.Office;

/// <summary>Typed result of one guarded row insertion.</summary>
public enum GanttRowInsertRefusalReason
{
    /// <summary>No active Excel workbook exists.</summary>
    NoActiveWorkbook = 0,

    /// <summary>The visible <c>tblGanttData</c> table was not found.</summary>
    TableMissing = 1,

    /// <summary>The active worksheet is protected.</summary>
    TargetProtected = 2,
}

/// <summary>The typed result of a row insertion.</summary>
/// <param name="Succeeded">Whether a row was appended.</param>
/// <param name="BodyIndex">The new row's one-based body index when successful.</param>
/// <param name="Refusal">The refusal reason when unsuccessful.</param>
public sealed record GanttRowInsertOutcome(
    bool Succeeded,
    int? BodyIndex,
    GanttRowInsertRefusalReason? Refusal)
{
    /// <summary>Creates a successful insertion outcome.</summary>
    /// <param name="bodyIndex">The new row's one-based body index.</param>
    /// <returns>A successful outcome.</returns>
    public static GanttRowInsertOutcome Ok(int bodyIndex) => new(true, bodyIndex, null);

    /// <summary>Creates a refusal outcome.</summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <returns>A refusal outcome.</returns>
    public static GanttRowInsertOutcome Refused(GanttRowInsertRefusalReason refusal) =>
        new(false, null, refusal);
}
