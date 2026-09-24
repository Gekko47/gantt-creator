namespace GanttCreator.Office;

/// <summary>Why row-identity repair refused to mutate.</summary>
public enum GanttRowIdentityRepairRefusalReason
{
    /// <summary>No active workbook exists.</summary>
    NoActiveWorkbook = 0,
    /// <summary>The Gantt table is missing.</summary>
    TableMissing = 1,
    /// <summary>The target is protected.</summary>
    TargetProtected = 2,
    /// <summary>An Id cell could not be written.</summary>
    WriteFailed = 3,
}

/// <summary>Typed result of an identity repair.</summary>
/// <param name="RepairedCount">The number of Id cells repaired.</param>
/// <param name="Refusal">The refusal reason on failure.</param>
public sealed record GanttRowIdentityRepairOutcome(
    int RepairedCount,
    GanttRowIdentityRepairRefusalReason? Refusal)
{
    /// <summary>Creates a successful outcome.</summary>
    /// <param name="repaired">The number of repaired cells.</param>
    /// <returns>The successful outcome.</returns>
    public static GanttRowIdentityRepairOutcome Ok(int repaired) => new(repaired, null);

    /// <summary>Creates a refusal.</summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <returns>The refusal outcome.</returns>
    public static GanttRowIdentityRepairOutcome Refused(GanttRowIdentityRepairRefusalReason refusal) =>
        new(0, refusal);
}

/// <summary>Repairs malformed, missing, and later duplicate row IDs only.</summary>
public interface IGanttRowIdentityRepairer
{
    /// <summary>Repairs only the Id column of the visible Gantt table.</summary>
    /// <returns>The typed repair result.</returns>
    GanttRowIdentityRepairOutcome Repair();
}
