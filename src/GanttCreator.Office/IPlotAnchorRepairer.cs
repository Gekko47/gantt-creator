namespace GanttCreator.Office;

/// <summary>Why plot-anchor repair refused to mutate.</summary>
public enum PlotAnchorRepairRefusalReason
{
    /// <summary>No active workbook exists.</summary>
    NoActiveWorkbook = 0,
    /// <summary>The Gantt worksheet or its table is missing.</summary>
    TableMissing = 1,
    /// <summary>The target is protected.</summary>
    TargetProtected = 2,
    /// <summary>The defined name could not be written.</summary>
    NameWriteFailed = 3,
}

/// <summary>Typed result of a plot-anchor repair.</summary>
/// <param name="Succeeded">Whether the anchor was repaired.</param>
/// <param name="Refusal">The refusal reason on failure.</param>
public sealed record PlotAnchorRepairOutcome(
    bool Succeeded,
    PlotAnchorRepairRefusalReason? Refusal)
{
    /// <summary>Creates a successful outcome.</summary>
    /// <returns>The successful outcome.</returns>
    public static PlotAnchorRepairOutcome Ok() => new(true, null);

    /// <summary>Creates a refusal.</summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <returns>The refusal outcome.</returns>
    public static PlotAnchorRepairOutcome Refused(PlotAnchorRepairRefusalReason refusal) => new(false, refusal);
}

/// <summary>Repairs the stored plot-anchor defined name from the live Gantt table.</summary>
public interface IPlotAnchorRepairer
{
    /// <summary>Re-derives and writes the plot anchor.</summary>
    /// <returns>The typed repair result.</returns>
    PlotAnchorRepairOutcome Repair();
}
