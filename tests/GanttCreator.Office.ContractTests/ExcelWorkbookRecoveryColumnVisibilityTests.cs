using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contracts for the repair path's column-visibility restore (R5.18h slice 2b): a stored
/// choice is re-applied after a rebuild, and every unreadable or unparseable stored value
/// leaves the schema's presentation in place.
/// </summary>
/// <remarks>
/// A rebuild that recreates the worksheet starts with all five authoring columns visible,
/// because the schema pass deliberately does not touch a user-visible column's hidden flag.
/// If this step silently did nothing, the user's choice would be lost on the one command
/// they ran precisely to get a working sheet back.
/// </remarks>
public sealed class ExcelWorkbookRecoveryColumnVisibilityTests
{
    /// <summary>A port that records the sets it was asked to apply.</summary>
    private sealed class FakeColumnVisibilityPort : IColumnVisibilityPort
    {
        internal List<IReadOnlySet<string>> Applied { get; } = [];

        public ColumnVisibilityPortOutcome ReadHidden() =>
            ColumnVisibilityPortOutcome.Ok(new HashSet<string>(StringComparer.Ordinal));

        public ColumnVisibilityPortOutcome Apply(IReadOnlySet<string> hidden)
        {
            Applied.Add(hidden);
            return ColumnVisibilityPortOutcome.Ok(hidden);
        }
    }

    /// <summary>Builds a successful read outcome carrying the given settings.</summary>
    private static ConfigReadOutcome Settings(params (string Key, string Value)[] pairs)
    {
        Dictionary<string, string> settings = new(StringComparer.Ordinal);
        foreach ((string key, string value) in pairs)
        {
            settings[key] = value;
        }

        return ConfigReadOutcome.Ok(
            "workbook",
            settings,
            GanttStyleRegistry.Empty,
            new Dictionary<string, double>());
    }

    /// <summary>Builds a recovery whose reader and port are the supplied fakes.</summary>
    private static (ExcelWorkbookRecovery Recovery, FakeColumnVisibilityPort Port, List<string> Trace) Recovery(
        ConfigReadOutcome outcome)
    {
        Mock<IConfigCatalogueReader> reader = new();
        _ = reader.Setup(r => r.Read()).Returns(outcome);
        FakeColumnVisibilityPort port = new();
        List<string> trace = [];
        ExcelWorkbookRecovery recovery = new(
            application: null,
            catalogueReader: reader.Object,
            columnVisibilityPort: port,
            trace: trace.Add);
        return (recovery, port, trace);
    }

    /// <summary>The stored choice is applied, so a rebuilt sheet keeps it.</summary>
    [Fact]
    public void A_stored_choice_is_applied()
    {
        (ExcelWorkbookRecovery recovery, FakeColumnVisibilityPort port, _) =
            Recovery(Settings((GanttColumnVisibility.SettingKey, "Start,Type")));

        recovery.ApplyStoredColumnVisibility();

        IReadOnlySet<string> applied = Assert.Single(port.Applied);
        Assert.Equal(2, applied.Count);
        Assert.Contains("Type", applied);
        Assert.Contains("Start", applied);
    }

    /// <summary>An absent key means nothing was ever hidden, so nothing is written.</summary>
    [Fact]
    public void An_absent_key_applies_nothing()
    {
        (ExcelWorkbookRecovery recovery, FakeColumnVisibilityPort port, List<string> trace) =
            Recovery(Settings(("TimeScale", "Month")));

        recovery.ApplyStoredColumnVisibility();

        Assert.Empty(port.Applied);
        Assert.Empty(trace);
    }

    /// <summary>A blank value is "nothing hidden", which the schema already shows.</summary>
    [Fact]
    public void A_blank_key_applies_nothing()
    {
        (ExcelWorkbookRecovery recovery, FakeColumnVisibilityPort port, _) =
            Recovery(Settings((GanttColumnVisibility.SettingKey, "  ")));

        recovery.ApplyStoredColumnVisibility();

        Assert.Empty(port.Applied);
    }

    /// <summary>
    /// A value that does not parse is refused and traced rather than read as "nothing
    /// hidden": guessing would un-hide a column the user hid.
    /// </summary>
    [Fact]
    public void A_malformed_value_is_refused_and_traced()
    {
        (ExcelWorkbookRecovery recovery, FakeColumnVisibilityPort port, List<string> trace) =
            Recovery(Settings((GanttColumnVisibility.SettingKey, "Type,,Start")));

        recovery.ApplyStoredColumnVisibility();

        Assert.Empty(port.Applied);
        Assert.Contains(trace, line => line.Contains("did not", StringComparison.Ordinal));
    }

    /// <summary>An unreadable settings table leaves the schema's presentation alone.</summary>
    [Fact]
    public void An_unreadable_catalogue_applies_nothing()
    {
        (ExcelWorkbookRecovery recovery, FakeColumnVisibilityPort port, List<string> trace) =
            Recovery(ConfigReadOutcome.Refused(ConfigReadRefusalReason.ConfigSheetMissing));

        recovery.ApplyStoredColumnVisibility();

        Assert.Empty(port.Applied);
        Assert.Contains(trace, line => line.Contains("could not be read", StringComparison.Ordinal));
    }
}
