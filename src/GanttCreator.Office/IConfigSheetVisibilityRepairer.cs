namespace GanttCreator.Office;

/// <summary>Why configuration-sheet visibility repair refused.</summary>
public enum ConfigSheetVisibilityRepairRefusalReason
{
    /// <summary>No active workbook exists.</summary>
    NoActiveWorkbook = 0,
    /// <summary>The configuration worksheet is missing.</summary>
    ConfigSheetMissing = 1,
    /// <summary>The workbook or sheet is protected.</summary>
    TargetProtected = 2,
}

/// <summary>Typed result of configuration-sheet visibility repair.</summary>
/// <param name="Succeeded">Whether visibility was repaired.</param>
/// <param name="Refusal">The refusal reason on failure.</param>
public sealed record ConfigSheetVisibilityRepairOutcome(
    bool Succeeded,
    ConfigSheetVisibilityRepairRefusalReason? Refusal)
{
    /// <summary>Creates a successful result.</summary>
    /// <returns>The successful result.</returns>
    public static ConfigSheetVisibilityRepairOutcome Ok() => new(true, null);

    /// <summary>Creates a refusal.</summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <returns>The refusal result.</returns>
    public static ConfigSheetVisibilityRepairOutcome Refused(ConfigSheetVisibilityRepairRefusalReason refusal) =>
        new(false, refusal);
}

/// <summary>Restores the configuration worksheet to VeryHidden.</summary>
public interface IConfigSheetVisibilityRepairer
{
    /// <summary>Repairs the configuration worksheet visibility.</summary>
    /// <returns>The typed result.</returns>
    ConfigSheetVisibilityRepairOutcome Repair();
}
