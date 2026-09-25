using GanttCreator.Core.ConfigIntegrity;

namespace GanttCreator.Office;

/// <summary>Why a configuration repair attempt refused to mutate.</summary>
public enum ConfigRepairRefusalReason
{
    /// <summary>No active workbook was supplied.</summary>
    NoActiveWorkbook = 0,
    /// <summary>The workbook or target sheet is protected.</summary>
    TargetProtected = 1,
    /// <summary>The plan contains a refused finding.</summary>
    PlanRefused = 2,
    /// <summary>The configuration catalogue could not be written.</summary>
    CatalogueWriteRefused = 3,
    /// <summary>The TypeOptions materialisation failed after catalogue repair.</summary>
    TypeOptionsUnavailable = 4,
    /// <summary>Visibility repair failed.</summary>
    VisibilityRepairFailed = 5,
    /// <summary>Plot-anchor repair failed.</summary>
    PlotAnchorRepairFailed = 6,
    /// <summary>Row-identity repair failed.</summary>
    IdentityRepairFailed = 7,
}

/// <summary>Typed result of one approved configuration repair.</summary>
/// <param name="RepairedCount">Number of approved findings acted on.</param>
/// <param name="ConfirmationSkippedCount">Number of confirm-required findings declined.</param>
/// <param name="RefusedCount">Number of findings refused.</param>
/// <param name="QuarantinedCount">Number of user rows quarantined.</param>
/// <param name="Refusal">The typed refusal, if any.</param>
public sealed record ConfigRepairOutcome(
    int RepairedCount,
    int ConfirmationSkippedCount,
    int RefusedCount,
    int QuarantinedCount,
    ConfigRepairRefusalReason? Refusal)
{
    /// <summary>Creates a successful result.</summary>
    /// <param name="repaired">Repaired count.</param>
    /// <param name="skipped">Declined confirmation count.</param>
    /// <param name="quarantined">Quarantine count.</param>
    /// <returns>The successful result.</returns>
    public static ConfigRepairOutcome Ok(int repaired, int skipped, int quarantined) =>
        new(repaired, skipped, 0, quarantined, null);

    /// <summary>Creates a refusal result.</summary>
    /// <param name="refusal">The typed refusal.</param>
    /// <returns>The refusal result.</returns>
    public static ConfigRepairOutcome Refused(ConfigRepairRefusalReason refusal) =>
        new(0, 0, 0, 0, refusal);
}

/// <summary>Plan-constrained configuration repair operation.</summary>
public interface IConfigRepairer
{
    /// <summary>Repairs the workbook according to the supplied plan and confirmation decision.</summary>
    /// <param name="plan">The pure Core repair plan.</param>
    /// <param name="confirmRequiredChanges">Whether the user confirmed changes that may touch user-authored data.</param>
    /// <returns>The typed repair result.</returns>
    ConfigRepairOutcome Repair(ConfigIntegrityPlan plan, bool confirmRequiredChanges);
}
