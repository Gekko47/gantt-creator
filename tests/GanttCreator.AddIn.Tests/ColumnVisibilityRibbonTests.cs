using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contracts for the column-visibility dropdown's projection and action (R5.18h):
/// the items, the state each label states, the port and key the toggle writes, and the
/// refusal that keeps the table authorable.
/// </summary>
/// <remarks>
/// These run against the session singleton because that is what the RibbonX callbacks
/// resolve, so the collection attribute keeps them serial with the other singleton tests
/// and <see cref="Dispose"/> resets it.
/// </remarks>
[Collection("diagnostics-service")]
public sealed class ColumnVisibilityRibbonTests : IDisposable
{
    public void Dispose() => RibbonStateService.Reset();

    /// <summary>A port that records what was applied and can refuse the write.</summary>
    private sealed class FakeColumnVisibilityPort(params string[] hidden) : IColumnVisibilityPort
    {
        private readonly ColumnVisibilityPortOutcome _read =
            ColumnVisibilityPortOutcome.Ok(new HashSet<string>(hidden, StringComparer.Ordinal));

        internal List<IReadOnlySet<string>> Applied { get; } = [];

        internal bool RefuseApply { get; init; }

        public ColumnVisibilityPortOutcome ReadHidden() => _read;

        public ColumnVisibilityPortOutcome Apply(IReadOnlySet<string> hidden)
        {
            if (RefuseApply)
            {
                return ColumnVisibilityPortOutcome.Refused(ColumnVisibilityRefusalReason.TargetProtected);
            }

            Applied.Add(hidden);
            return ColumnVisibilityPortOutcome.Ok(hidden);
        }
    }

    /// <summary>Arms the singleton with a port and a writer, returning the recorded writes.</summary>
    private static (RibbonStateService Service, Mock<IConfigCatalogueWriter> Writer) Arm(
        IColumnVisibilityPort port)
    {
        RibbonStateService service = RibbonStateService.Instance;
        Mock<IConfigCatalogueWriter> writer = new();
        service.SetCatalogueWriter(writer.Object);
        service.SetColumnVisibilityPort(port);
        return (service, writer);
    }

    /// <summary>The last settings payload the writer was handed.</summary>
    private static Dictionary<string, string>? LastWritten(Mock<IConfigCatalogueWriter> writer)
    {
        foreach (IInvocation invocation in writer.Invocations.Reverse())
        {
            if (!string.Equals(
                invocation.Method.Name,
                nameof(IConfigCatalogueWriter.WriteSettings),
                StringComparison.Ordinal))
            {
                continue;
            }

            return invocation.Arguments[0] as Dictionary<string, string>;
        }

        return null;
    }

    /// <summary>The dropdown offers exactly the columns the Core contract names.</summary>
    [Fact]
    public void The_dropdown_offers_one_item_per_offered_column()
    {
        Assert.Equal(GanttColumnVisibility.ToggleableColumnNames.Count, RibbonStateService.GetColumnVisibilityCount());
        Assert.Equal("colType", RibbonStateService.GetColumnVisibilityItemID(0));
        Assert.Equal("colDuration", RibbonStateService.GetColumnVisibilityItemID(4));
    }

    /// <summary>An out-of-range index yields an empty item rather than throwing.</summary>
    [Fact]
    public void An_out_of_range_index_yields_an_empty_item()
    {
        Assert.Equal(string.Empty, RibbonStateService.GetColumnVisibilityItemID(5));
        Assert.Equal(string.Empty, RibbonStateService.GetColumnVisibilityItemID(-1));
    }

    /// <summary>Each label states the live visibility the port reports.</summary>
    [Fact]
    public void The_item_labels_state_the_live_visibility()
    {
        (RibbonStateService service, _) = Arm(new FakeColumnVisibilityPort("Start"));

        Assert.Equal("Type (shown)", service.GetColumnVisibilityItemLabel(0));
        Assert.Equal("Start (hidden)", service.GetColumnVisibilityItemLabel(2));
        Assert.Equal(string.Empty, service.GetColumnVisibilityItemLabel(9));
    }

    /// <summary>
    /// Choosing an item hides the column on the worksheet, records the choice in the
    /// settings key, and the label states the new visibility.
    /// </summary>
    [Fact]
    public void Toggling_hides_applies_persists_and_states_the_result()
    {
        FakeColumnVisibilityPort port = new();
        (RibbonStateService service, Mock<IConfigCatalogueWriter> writer) = Arm(port);

        service.ToggleColumnVisibility("colType");

        Assert.Equal("Type", Assert.Single(Assert.Single(port.Applied)));
        Assert.Equal(
            "Type",
            LastWritten(writer)?[GanttColumnVisibility.SettingKey]);
        Assert.Equal("Type (hidden)", service.GetColumnVisibilityItemLabel(0));
        Assert.Equal("Description (shown)", service.GetColumnVisibilityItemLabel(1));
    }

    /// <summary>Choosing the same item again shows the column and clears the stored name.</summary>
    [Fact]
    public void Toggling_twice_restores_the_column_and_clears_the_key()
    {
        FakeColumnVisibilityPort port = new();
        (RibbonStateService service, Mock<IConfigCatalogueWriter> writer) = Arm(port);

        service.ToggleColumnVisibility("colStart");
        service.ToggleColumnVisibility("colStart");

        Assert.Equal(string.Empty, LastWritten(writer)?[GanttColumnVisibility.SettingKey]);
        Assert.Equal("Start (shown)", service.GetColumnVisibilityItemLabel(2));
        Assert.Empty(port.Applied[1]);
    }

    /// <summary>
    /// Hiding the last authoring column is refused: nothing is applied, nothing is
    /// written, and the states do not move.
    /// </summary>
    [Fact]
    public void Hiding_the_last_authoring_column_changes_nothing()
    {
        FakeColumnVisibilityPort port = new();
        (RibbonStateService service, Mock<IConfigCatalogueWriter> writer) = Arm(port);

        service.ToggleColumnVisibility("colType");
        service.ToggleColumnVisibility("colDescription");
        service.ToggleColumnVisibility("colStart");
        int appliedAfterThree = port.Applied.Count;

        service.ToggleColumnVisibility("colFinish");

        Assert.Equal(appliedAfterThree, port.Applied.Count);
        Assert.Equal("Type,Description,Start", LastWritten(writer)?[GanttColumnVisibility.SettingKey]);
        Assert.Equal("Finish (shown)", service.GetColumnVisibilityItemLabel(3));
    }

    /// <summary>An item id the dropdown never issued is ignored, not guessed at.</summary>
    [Fact]
    public void An_unknown_item_id_is_ignored()
    {
        FakeColumnVisibilityPort port = new();
        (RibbonStateService service, Mock<IConfigCatalogueWriter> writer) = Arm(port);

        service.ToggleColumnVisibility("colId");
        service.ToggleColumnVisibility("Type");
        service.ToggleColumnVisibility(null);

        Assert.Empty(port.Applied);
        Assert.Null(LastWritten(writer));
    }

    /// <summary>
    /// A port refusal is reported and nothing is recorded: the setting is only written
    /// after the worksheet accepted the change.
    /// </summary>
    [Fact]
    public void A_refused_apply_writes_no_setting()
    {
        FakeColumnVisibilityPort port = new() { RefuseApply = true };
        (RibbonStateService service, Mock<IConfigCatalogueWriter> writer) = Arm(port);

        service.ToggleColumnVisibility("colType");

        Assert.Null(LastWritten(writer));
        Assert.Equal("Type (shown)", service.GetColumnVisibilityItemLabel(0));
    }
}
