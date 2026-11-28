using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contracts for the column-visibility split button's projection and action (R5.18h,
/// R5.22): the per-column rows, the all/none primary, the port and key the toggle writes,
/// and the compensation path that keeps the sheet and the setting together.
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

    /// <summary>A panel-grid port that records the visible set the width display asks for.</summary>
    private sealed class RecordingPanelGridMeasurementPort(PanelCellGrid grid) : IPanelGridMeasurementPort
    {
        internal IReadOnlyList<string> LastColumns { get; private set; } = [];

        public PanelGridOutcome Measure(
            IReadOnlyList<string> includedColumns,
            IReadOnlyList<double>? bodyRowHeightsPt = null)
        {
            LastColumns = [.. includedColumns];
            _ = bodyRowHeightsPt;
            return PanelGridOutcome.Ok(grid);
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

    /// <summary>
    /// Every row the menu declares names a column the Core contract offers, and the row's
    /// label states that column's live visibility. An id the menu never issued — an engine
    /// column, or none at all — yields nothing rather than a guessed label.
    /// </summary>
    [Fact]
    public void The_menu_rows_name_every_offered_column()
    {
        (RibbonStateService service, _) = Arm(new FakeColumnVisibilityPort());

        foreach (string name in GanttColumnVisibility.ToggleableColumnNames)
        {
            Assert.Equal(
                name + " (shown)",
                service.GetColumnVisibilityLabel(
                    RibbonControlIds.ColumnItemPrefix + GanttColumnVisibility.RibbonItemId(name)));
        }

        // The positive refusal cases: an engine column (never offered) and a null id.
        Assert.Equal(string.Empty, service.GetColumnVisibilityLabel("colId"));
        Assert.Equal(string.Empty, service.GetColumnVisibilityLabel(null));
    }

    /// <summary>Each label states the live visibility the port reports.</summary>
    [Fact]
    public void The_item_labels_state_the_live_visibility()
    {
        (RibbonStateService service, _) = Arm(new FakeColumnVisibilityPort("Start"));

        Assert.Equal("Type (shown)", service.GetColumnVisibilityLabel("colType"));
        Assert.Equal("Start (hidden)", service.GetColumnVisibilityLabel("colStart"));
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
        Assert.Equal("Type (hidden)", service.GetColumnVisibilityLabel("colType"));
        Assert.Equal("Description (shown)", service.GetColumnVisibilityLabel("colDescription"));
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
        Assert.Equal("Start (shown)", service.GetColumnVisibilityLabel("colStart"));
        Assert.Empty(port.Applied[1]);
    }

    /// <summary>
    /// Hiding every column is allowed and persisted (ADR-0065 D1): the old last-column
    /// refusal is gone, so the sixth toggle applies and records the full set.
    /// </summary>
    [Fact]
    public void Hiding_every_column_is_applied_persisted_and_stated()
    {
        FakeColumnVisibilityPort port = new();
        (RibbonStateService service, Mock<IConfigCatalogueWriter> writer) = Arm(port);

        foreach (string itemId in new[]
                 {
                     "colType", "colActivityID", "colDescription",
                     "colStart", "colFinish", "colDuration",
                 })
        {
            service.ToggleColumnVisibility(itemId);
        }

        Assert.Equal(6, port.Applied.Count);
        Assert.Equal(
            string.Join(',', GanttColumnVisibility.ToggleableColumnNames),
            LastWritten(writer)?[GanttColumnVisibility.SettingKey]);
        Assert.Equal("Duration (hidden)", service.GetColumnVisibilityLabel("colDuration"));
        Assert.Equal("Show all columns", service.GetColumnVisibilityAllLabel());
    }

    /// <summary>
    /// The primary hides every column from an all-visible state and persists the full list
    /// (ADR-0065 D3).
    /// </summary>
    [Fact]
    public void Toggling_all_hides_every_column_and_persists_the_full_list()
    {
        FakeColumnVisibilityPort port = new();
        (RibbonStateService service, Mock<IConfigCatalogueWriter> writer) = Arm(port);

        service.ToggleAllColumnVisibility();

        Assert.Equal(GanttColumnVisibility.ToggleableColumnNames.Count, port.Applied[0].Count);
        Assert.Equal(
            string.Join(',', GanttColumnVisibility.ToggleableColumnNames),
            LastWritten(writer)?[GanttColumnVisibility.SettingKey]);
        Assert.Equal("Show all columns", service.GetColumnVisibilityAllLabel());
    }

    /// <summary>
    /// A partial state shows all rather than completing the hide, so the primary is a
    /// forgiving reset and the menu keeps precise hiding (ADR-0065 D3).
    /// </summary>
    [Fact]
    public void Toggling_all_from_a_partial_state_shows_every_column()
    {
        FakeColumnVisibilityPort port = new("Start");
        (RibbonStateService service, _) = Arm(port);

        service.ToggleAllColumnVisibility();

        Assert.Empty(port.Applied[0]);
        Assert.Equal("Hide all columns", service.GetColumnVisibilityAllLabel());
    }

    /// <summary>
    /// A refused persist re-applies the previous set, so the sheet and the setting cannot
    /// describe different choices (ADR-0065 D5).
    /// </summary>
    [Fact]
    public void A_refused_persist_restores_the_previous_state()
    {
        FakeColumnVisibilityPort port = new();
        (RibbonStateService service, Mock<IConfigCatalogueWriter> writer) = Arm(port);
        writer.Setup(w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()))
            .Returns(ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.HostRejected));

        service.ToggleAllColumnVisibility();

        Assert.Equal(2, port.Applied.Count);
        Assert.Equal(GanttColumnVisibility.ToggleableColumnNames.Count, port.Applied[0].Count);
        Assert.Empty(port.Applied[1]);
        Assert.Equal("Hide all columns", service.GetColumnVisibilityAllLabel());
    }

    /// <summary>
    /// The width readout measures the candidate visible set, so an all-hidden toggle
    /// measures zero columns and still reports the plot width the chart will use.
    /// </summary>
    [Fact]
    public void Toggling_all_recomputes_the_width_from_the_empty_visible_set()
    {
        FakeColumnVisibilityPort port = new();
        (RibbonStateService service, _) = Arm(port);
        // SetPreset persists the layout; a loose writer mock would return null and the
        // service would dereference it. The width fact is about the measurement, not the write.
        service.SetCatalogueWriter(null);
        PanelCellGrid grid = PanelCellGrid.TryCreate(
            [new PanelColumn("Type", 300.0)],
            [12.0],
            12.0,
            []).Grid ?? throw new InvalidOperationException("Fixture grid should be valid.");
        RecordingPanelGridMeasurementPort measurement = new(grid);
        service.SetPanelGridMeasurementPort(measurement);
        service.SetPreset(SizePresets.A4Landscape);

        service.ToggleAllColumnVisibility();

        Assert.Empty(measurement.LastColumns);
        Assert.False(string.IsNullOrEmpty(service.GetPlotWidth()));
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
        Assert.Equal("Type (shown)", service.GetColumnVisibilityLabel("colType"));
    }

    /// <summary>
    /// R5.24: a column toggle writes its setting in place and does not move the active tab, so it
    /// must not fire an unconditional <c>ActivateTab</c>. That restore was the reported flicker:
    /// it refocused a tab that had never lost focus (owner report 2026-11-28).
    /// </summary>
    [Fact]
    public void A_column_toggle_does_not_activate_the_tab()
    {
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
        var ribbon = new Mock<IRibbonUI>();
        RibbonStateService service = RibbonStateTestSupport.Arm(adapter, ribbon: ribbon.Object);
        var reader = new Mock<IGanttTableReader>(MockBehavior.Strict);
        _ = reader.Setup(r => r.Read())
            .Returns(GanttTableReadOutcome.Ok([], configurationPresent: true));
        service.SetTableReader(reader.Object);
        service.SetColumnVisibilityPort(new FakeColumnVisibilityPort());
        var writer = new Mock<IConfigCatalogueWriter>();
        _ = writer.Setup(w => w.WriteSettings(It.IsAny<IReadOnlyDictionary<string, string>>()))
            .Returns(ConfigWriteOutcome.Ok());
        service.SetCatalogueWriter(writer.Object);
        service.Refresh();

        // The refresh that brings the surface to Ready activates the tab itself.
        ribbon.Invocations.Clear();

        service.ToggleAllColumnVisibility();

        ribbon.Verify(r => r.ActivateTab(It.IsAny<string>()), Times.Never);
    }
}
