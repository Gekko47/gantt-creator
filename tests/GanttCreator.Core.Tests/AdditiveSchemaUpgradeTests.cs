using GanttCreator.Core;
using GanttCreator.Core.ConfigIntegrity;

namespace GanttCreator.Core.Tests;

/// <summary>
/// Tests for <see cref="AdditiveSchemaUpgrade"/>, the gate that decides whether a
/// schema bump is a one-step settings-only append (ADR-0073). Every refusal is
/// exercised positively: a bad input must produce the typed refusal, not a plan.
/// </summary>
public sealed class AdditiveSchemaUpgradeTests
{
    private static string[] ApprovedKeys =>
        [.. GanttCatalogues.Settings.Select(setting => setting.Key)];

    [Fact]
    public void A_prefix_of_the_approved_keys_plans_the_appended_rows()
    {
        string[] approved = ApprovedKeys;
        string[] stored = [.. approved.Take(approved.Length - 2)];
        int storedVersion = GanttSchemaVersion.CurrentSchemaVersion - 1;

        AdditiveSchemaUpgradeOutcome outcome = AdditiveSchemaUpgrade.TryPlan(
            storedVersion,
            stored,
            GanttSchemaVersion.CurrentSchemaVersion,
            approved);

        Assert.True(outcome.Succeeded);
        Assert.Equal(storedVersion, outcome.Plan!.StoredVersion);
        Assert.Equal(approved[^2..], outcome.Plan.MissingKeys);
    }

    [Fact]
    public void The_current_version_is_not_an_upgrade()
    {
        int version = GanttSchemaVersion.CurrentSchemaVersion;

        AdditiveSchemaUpgradeOutcome outcome = AdditiveSchemaUpgrade.TryPlan(
            version,
            ApprovedKeys,
            version,
            ApprovedKeys);

        Assert.Equal(AdditiveSchemaUpgradeRefusal.NotOlder, outcome.Refusal);
    }

    [Fact]
    public void A_newer_workbook_is_refused()
    {
        int current = GanttSchemaVersion.CurrentSchemaVersion;

        AdditiveSchemaUpgradeOutcome outcome = AdditiveSchemaUpgrade.TryPlan(
            current + 1,
            ApprovedKeys,
            current,
            ApprovedKeys);

        Assert.Equal(AdditiveSchemaUpgradeRefusal.Newer, outcome.Refusal);
    }

    [Fact]
    public void A_multi_version_gap_is_refused_because_other_catalogues_may_have_changed()
    {
        int current = GanttSchemaVersion.CurrentSchemaVersion;

        AdditiveSchemaUpgradeOutcome outcome = AdditiveSchemaUpgrade.TryPlan(
            current - 2,
            ApprovedKeys,
            current,
            ApprovedKeys);

        Assert.Equal(AdditiveSchemaUpgradeRefusal.MultiVersionGap, outcome.Refusal);
    }

    [Fact]
    public void An_extra_key_is_refused()
    {
        string[] approved = ApprovedKeys;
        string[] stored = [.. approved, "UnapprovedKey"];

        AdditiveSchemaUpgradeOutcome outcome = AdditiveSchemaUpgrade.TryPlan(
            GanttSchemaVersion.CurrentSchemaVersion - 1,
            stored,
            GanttSchemaVersion.CurrentSchemaVersion,
            approved);

        Assert.Equal(AdditiveSchemaUpgradeRefusal.ExtraKey, outcome.Refusal);
    }

    [Fact]
    public void An_empty_stored_set_is_refused()
    {
        AdditiveSchemaUpgradeOutcome outcome = AdditiveSchemaUpgrade.TryPlan(
            GanttSchemaVersion.CurrentSchemaVersion - 1,
            [],
            GanttSchemaVersion.CurrentSchemaVersion,
            ApprovedKeys);

        Assert.Equal(AdditiveSchemaUpgradeRefusal.EmptyKeySet, outcome.Refusal);
    }

    [Fact]
    public void A_reordered_prefix_is_refused()
    {
        string[] approved = ApprovedKeys;
        string[] stored = [.. approved.Take(approved.Length - 2)];
        (stored[0], stored[1]) = (stored[1], stored[0]);

        AdditiveSchemaUpgradeOutcome outcome = AdditiveSchemaUpgrade.TryPlan(
            GanttSchemaVersion.CurrentSchemaVersion - 1,
            stored,
            GanttSchemaVersion.CurrentSchemaVersion,
            approved);

        Assert.Equal(AdditiveSchemaUpgradeRefusal.KeyOrderMismatch, outcome.Refusal);
    }

    [Fact]
    public void A_missing_middle_key_is_refused_even_when_the_count_looks_like_an_append()
    {
        string[] approved = ApprovedKeys;

        // Replace a middle key with the last key, keeping the list shorter than the
        // approved set. A count-only check would call this an append; the prefix
        // check must see the displaced key.
        string[] stored = [.. approved.Take(approved.Length - 3), approved[^1]];

        AdditiveSchemaUpgradeOutcome outcome = AdditiveSchemaUpgrade.TryPlan(
            GanttSchemaVersion.CurrentSchemaVersion - 1,
            stored,
            GanttSchemaVersion.CurrentSchemaVersion,
            approved);

        Assert.Equal(AdditiveSchemaUpgradeRefusal.KeyOrderMismatch, outcome.Refusal);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void A_non_positive_stored_version_throws(int version)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            AdditiveSchemaUpgrade.TryPlan(
                version,
                ApprovedKeys,
                GanttSchemaVersion.CurrentSchemaVersion,
                ApprovedKeys));
    }

    [Fact]
    public void A_non_positive_current_version_throws()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            AdditiveSchemaUpgrade.TryPlan(
                GanttSchemaVersion.CurrentSchemaVersion - 1,
                ApprovedKeys,
                0,
                ApprovedKeys));
    }
}
