namespace GanttCreator.Core.ConfigIntegrity;

/// <summary>Why an additive settings upgrade could not be planned.</summary>
public enum AdditiveSchemaUpgradeRefusal
{
    /// <summary>The stored version is the current version; there is nothing to upgrade.</summary>
    NotOlder = 0,

    /// <summary>The stored version is newer than the running add-in.</summary>
    Newer = 1,

    /// <summary>
    /// The gap is more than one schema step, so a settings-only change cannot be
    /// proved and another catalogue may also have changed.
    /// </summary>
    MultiVersionGap = 2,

    /// <summary>The stored key list carries a key the approved catalogue does not.</summary>
    ExtraKey = 3,

    /// <summary>The stored settings table has no keys, so nothing can be preserved.</summary>
    EmptyKeySet = 4,

    /// <summary>The stored keys are not an exact prefix of the approved keys.</summary>
    KeyOrderMismatch = 5,
}

/// <summary>One approved additive upgrade: the version and the rows to append.</summary>
/// <param name="StoredVersion">The version on the workbook.</param>
/// <param name="MissingKeys">The approved keys absent from the stored table, in catalogue order.</param>
public sealed record AdditiveSchemaUpgradePlan(
    int StoredVersion,
    IReadOnlyList<string> MissingKeys);

/// <summary>The typed result of planning an additive settings upgrade.</summary>
/// <param name="Plan">The plan, or <see langword="null"/> on a refusal.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/> on success.</param>
public sealed record AdditiveSchemaUpgradeOutcome(
    AdditiveSchemaUpgradePlan? Plan,
    AdditiveSchemaUpgradeRefusal? Refusal)
{
    /// <summary>Gets whether an upgrade was planned.</summary>
    public bool Succeeded => Plan is not null;
}

/// <summary>
/// The single authority that decides whether a schema bump is a one-step,
/// settings-only upgrade that can be repaired in place (ADR-0073).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why the gate is strict.</b> An append-only settings change is provably
/// loss-free: the existing values are preserved and the new keys receive their
/// catalogue defaults. A reorder, an extra key, a missing middle key, or a
/// multi-version gap is not provably settings-only, so it falls through to the
/// confirmation-required structural repair.
/// </para>
/// <para>
/// <b>Why one version step.</b> The settings prefix proves the settings contract
/// changed by appending only. It cannot prove that <c>tblGanttTypes</c> or
/// <c>tblGanttMetrics</c> did not change in the same range, so a multi-step gap
/// is refused rather than risk a narrow rewrite that leaves half the catalogue
/// stale. The upgrade path is a version-range migration table.
/// </para>
/// </remarks>
public static class AdditiveSchemaUpgrade
{
    /// <summary>
    /// Attempts to plan an additive settings upgrade.
    /// </summary>
    /// <param name="storedVersion">The version stored on the workbook.</param>
    /// <param name="storedKeys">The stored settings keys, in body order.</param>
    /// <param name="currentVersion">The version the running add-in requires.</param>
    /// <param name="approvedKeys">The approved settings keys, in catalogue order.</param>
    /// <returns>The plan, or the typed refusal.</returns>
    /// <exception cref="ArgumentNullException">Thrown when a key list is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when either version is not positive.
    /// </exception>
    public static AdditiveSchemaUpgradeOutcome TryPlan(
        int storedVersion,
        IReadOnlyList<string> storedKeys,
        int currentVersion,
        IReadOnlyList<string> approvedKeys)
    {
        ArgumentNullException.ThrowIfNull(storedKeys);
        ArgumentNullException.ThrowIfNull(approvedKeys);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(storedVersion);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(currentVersion);

        if (storedVersion > currentVersion)
        {
            return Refused(AdditiveSchemaUpgradeRefusal.Newer);
        }

        if (storedVersion == currentVersion)
        {
            return Refused(AdditiveSchemaUpgradeRefusal.NotOlder);
        }

        if (currentVersion - storedVersion > 1)
        {
            return Refused(AdditiveSchemaUpgradeRefusal.MultiVersionGap);
        }

        // A stored key that the catalogue does not define is not an append-only
        // change: preserving the table would leave an unapproved row the reader
        // rejects, and dropping it would discard a value the user may own.
        if (storedKeys.Count > approvedKeys.Count)
        {
            return Refused(AdditiveSchemaUpgradeRefusal.ExtraKey);
        }

        if (storedKeys.Count == 0 && approvedKeys.Count > 0)
        {
            return Refused(AdditiveSchemaUpgradeRefusal.EmptyKeySet);
        }

        for (int index = 0; index < storedKeys.Count; index++)
        {
            if (!string.Equals(storedKeys[index], approvedKeys[index], StringComparison.Ordinal))
            {
                return Refused(AdditiveSchemaUpgradeRefusal.KeyOrderMismatch);
            }
        }

        string[] missing = [.. approvedKeys.Skip(storedKeys.Count)];
        return new AdditiveSchemaUpgradeOutcome(
            new AdditiveSchemaUpgradePlan(storedVersion, missing),
            null);
    }

    private static AdditiveSchemaUpgradeOutcome Refused(AdditiveSchemaUpgradeRefusal refusal) =>
        new(null, refusal);
}
