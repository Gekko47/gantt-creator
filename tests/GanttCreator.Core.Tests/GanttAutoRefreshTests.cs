using GanttCreator.Core;

namespace GanttCreator.Core.Tests;

/// <summary>
/// The R5.20 automatic-refresh switch: one persisted key, read fail-open, written in one form.
/// </summary>
/// <remarks>
/// The read is deliberately lenient and the write is deliberately single-form. A test that only
/// asserted the round trip would pass with either rule broken, so both halves are pinned
/// separately: the exact text <see cref="GanttAutoRefresh.Format"/> emits, and the value every
/// hand-written workbook may carry.
/// </remarks>
public sealed class GanttAutoRefreshTests
{
    [Theory]
    [InlineData(true, "TRUE")]
    [InlineData(false, "FALSE")]
    public void Format_emits_the_single_stored_form(bool enabled, string expected) =>
        Assert.Equal(expected, GanttAutoRefresh.Format(enabled));

    /// <summary>
    /// Absent, blank, and the enabled token all mean enabled. The first case is the one that
    /// matters: every workbook written before the key existed has no row, and it must behave
    /// exactly as it did rather than silently stopping the automatic refresh.
    /// </summary>
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("TRUE")]
    [InlineData("true")]
    public void An_absent_or_blank_value_reads_as_enabled(string? stored) =>
        Assert.True(GanttAutoRefresh.Parse(stored));

    /// <summary>The disabled token is the only value that switches the feature off.</summary>
    [Theory]
    [InlineData("FALSE")]
    [InlineData("false")]
    [InlineData(" FALSE ")]
    public void The_disabled_token_reads_as_disabled(string stored) =>
        Assert.False(GanttAutoRefresh.Parse(stored));

    /// <summary>
    /// An unreadable value is enabled, not disabled: the failure a user cannot see or undo from
    /// the ribbon would be every automatic pass stopping on a workbook nobody switched off.
    /// </summary>
    [Theory]
    [InlineData("yes")]
    [InlineData("0")]
    [InlineData("Off")]
    public void An_unreadable_value_falls_open(string stored) =>
        Assert.True(GanttAutoRefresh.Parse(stored));

    [Fact]
    public void The_round_trip_is_stable()
    {
        Assert.True(GanttAutoRefresh.Parse(GanttAutoRefresh.Format(true)));
        Assert.False(GanttAutoRefresh.Parse(GanttAutoRefresh.Format(false)));
    }

    /// <summary>
    /// The key is part of the catalogue, not a literal at the call sites: a writer that names its
    /// own string would keep working after a rename while the catalogue hash moved.
    /// </summary>
    [Fact]
    public void The_key_is_the_catalogue_row_with_the_enabled_default()
    {
        GanttSettingDefinition row = Assert.Single(
            GanttCatalogues.Settings,
            setting => string.Equals(setting.Key, GanttAutoRefresh.SettingKey, StringComparison.Ordinal));

        Assert.Equal(GanttAutoRefresh.EnabledValue, row.DefaultValue);
        Assert.True(GanttAutoRefresh.Parse(row.DefaultValue));
    }
}
