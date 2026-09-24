using GanttCreator.Core.ConfigIntegrity;

namespace GanttCreator.Office;

/// <summary>Why a configuration integrity check could not inspect the workbook.</summary>
public enum ConfigIntegrityCheckRefusalReason
{
    /// <summary>No active workbook was supplied by Excel.</summary>
    NoActiveWorkbook = 0,
}

/// <summary>The typed result of one read-only configuration integrity check.</summary>
public sealed record ConfigIntegrityCheckOutcome(
    IReadOnlyList<ConfigIntegrityFinding> Findings,
    ConfigIntegrityCheckRefusalReason? Refusal)
{
    private static readonly IReadOnlyList<ConfigIntegrityFinding> _empty = [];

    /// <summary>Creates a successful check result.</summary>
    /// <param name="findings">All findings in deterministic order.</param>
    /// <returns>The successful result.</returns>
    public static ConfigIntegrityCheckOutcome Ok(IReadOnlyList<ConfigIntegrityFinding> findings) =>
        new(findings, null);

    /// <summary>Creates a no-active-workbook refusal.</summary>
    /// <returns>The typed refusal.</returns>
    public static ConfigIntegrityCheckOutcome NoActiveWorkbook() =>
        new(_empty, ConfigIntegrityCheckRefusalReason.NoActiveWorkbook);
}

/// <summary>Read-only configuration integrity checker over the active workbook.</summary>
public interface IConfigIntegrityChecker
{
    /// <summary>Checks the active workbook without mutating it.</summary>
    /// <returns>All detected findings or a typed refusal.</returns>
    ConfigIntegrityCheckOutcome Check();
}
