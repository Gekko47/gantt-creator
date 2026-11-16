using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The R5.20 automatic-refresh switch as the ribbon presents it: pressed state, wording, the
/// persisted write, the refused-write report, and the seed from the workbook's own setting.
/// </summary>
/// <remarks>
/// Shares the <c>diagnostics-service</c> collection because the switch lives on the session
/// singleton <see cref="AutoRefreshPreferences.Shared"/>; xUnit runs a collection's classes
/// sequentially, so these facts cannot race the host-lifecycle tests.
/// </remarks>
[Collection("diagnostics-service")]
public sealed class AutoRefreshRibbonTests : IDisposable
{
    public AutoRefreshRibbonTests() => RibbonStateService.Reset();

    public void Dispose()
    {
        RibbonStateService.Reset();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void Turning_the_switch_off_records_the_state_and_persists_false()
    {
        var writer = new Mock<IConfigCatalogueWriter>();
        IReadOnlyDictionary<string, string>? written = null;
        _ = writer
            .Setup(w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()))
            .Callback<IReadOnlyDictionary<string, string>>(values => written = values)
            .Returns(ConfigWriteOutcome.Ok());
        Mock<IRibbonUI> ribbon = new();
        var service = new RibbonStateService();
        service.SetCatalogueWriter(writer.Object);
        service.SetRibbon(ribbon.Object);

        service.SetAutoRefresh(false);

        Assert.False(AutoRefreshPreferences.Shared.AutoRefreshEnabled);
        Assert.NotNull(written);
        Assert.Equal(GanttAutoRefresh.DisabledValue, written![GanttAutoRefresh.SettingKey]);
        Assert.Single(written);
        ribbon.Verify(
            r => r.InvalidateControl(RibbonControlIds.ToggleAutoRefresh),
            Times.Once);
    }

    [Fact]
    public void Turning_the_switch_on_records_the_state_and_persists_true()
    {
        AutoRefreshPreferences.Shared.AutoRefreshEnabled = false;
        var writer = new Mock<IConfigCatalogueWriter>();
        IReadOnlyDictionary<string, string>? written = null;
        _ = writer
            .Setup(w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()))
            .Callback<IReadOnlyDictionary<string, string>>(values => written = values)
            .Returns(ConfigWriteOutcome.Ok());
        var service = new RibbonStateService();
        service.SetCatalogueWriter(writer.Object);

        service.SetAutoRefresh(true);

        Assert.True(AutoRefreshPreferences.Shared.AutoRefreshEnabled);
        Assert.NotNull(written);
        Assert.Equal(GanttAutoRefresh.EnabledValue, written![GanttAutoRefresh.SettingKey]);
    }

    /// <summary>
    /// The switch takes effect even when the workbook refuses to record it: the click must not
    /// silently do nothing, and the user must be told the choice will not survive reopening.
    /// </summary>
    [Fact]
    public void A_refused_persist_still_applies_the_switch_and_reports_it()
    {
        var writer = new Mock<IConfigCatalogueWriter>();
        _ = writer
            .Setup(w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()))
            .Returns(ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.HostRejected));
        List<string> warnings = [];
        var service = new RibbonStateService();
        service.SetCatalogueWriter(writer.Object);
        service.SetWarningPresenter(warnings.Add);

        service.SetAutoRefresh(false);

        Assert.False(AutoRefreshPreferences.Shared.AutoRefreshEnabled);
        string warning = Assert.Single(warnings);
        Assert.Contains("automatic refresh", warning, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// The workbook's value is adopted on activation and on a workbook change, through
    /// <see cref="GanttAutoRefresh.Parse"/>: an absent row means enabled, the disabled token
    /// means disabled, and an unreadable value stays enabled.
    /// </summary>
    [Fact]
    public void The_switch_is_seeded_from_the_workbook_settings()
    {
        AutoRefreshPreferences.Shared.AutoRefreshEnabled = true;

        RibbonStateService.SeedPersistedSwitches(new Dictionary<string, string>(StringComparer.Ordinal)
        {
            [GanttAutoRefresh.SettingKey] = GanttAutoRefresh.DisabledValue,
        });
        Assert.False(AutoRefreshPreferences.Shared.AutoRefreshEnabled);

        RibbonStateService.SeedPersistedSwitches(new Dictionary<string, string>(StringComparer.Ordinal)
        {
            [GanttAutoRefresh.SettingKey] = GanttAutoRefresh.EnabledValue,
        });
        Assert.True(AutoRefreshPreferences.Shared.AutoRefreshEnabled);

        RibbonStateService.SeedPersistedSwitches(new Dictionary<string, string>(StringComparer.Ordinal));
        Assert.True(AutoRefreshPreferences.Shared.AutoRefreshEnabled);
    }

    [Fact]
    public void The_seed_rejects_a_null_settings_map() =>
        Assert.Throws<ArgumentNullException>(() => RibbonStateService.SeedPersistedSwitches(null!));

    /// <summary>
    /// Pressed means enabled, read from the same object the scheduler reads, so the button cannot
    /// disagree with the pipeline.
    /// </summary>
    [Fact]
    public void The_pressed_state_reads_the_shared_switch()
    {
        Mock<IRibbonControl> control = new();

        AutoRefreshPreferences.Shared.AutoRefreshEnabled = true;
        Assert.True(RibbonStateService.GetToggleAutoRefreshPressed(control.Object));

        AutoRefreshPreferences.Shared.AutoRefreshEnabled = false;
        Assert.False(RibbonStateService.GetToggleAutoRefreshPressed(control.Object));
    }

    [Theory]
    [InlineData(true, "Auto refresh on")]
    [InlineData(false, "Auto refresh off")]
    public void The_label_states_the_current_state(bool enabled, string expected)
    {
        // The wording function is pinned once, and the ribbon getter is proved to read the same
        // switch rather than restating the state.
        Assert.Equal(expected, RibbonStateService.AutoRefreshLabel(enabled));

        AutoRefreshPreferences.Shared.AutoRefreshEnabled = enabled;
        Assert.Equal(expected, RibbonStateService.GetToggleAutoRefreshLabel(null));
    }
}
