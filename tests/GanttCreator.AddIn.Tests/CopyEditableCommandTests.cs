using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// R6.8: the Copy Editable command is the pipeline's owner and nothing more. These facts pin
/// the stage order, the export-selection rule, the verified copy wiring, the D1 message (with
/// the owner's partial-export clause), every refusal line, and presenter degradation.
/// </summary>
/// <remarks>
/// The command holds no Office dependency of its own — every host interaction rides the
/// supplied seams — so these tests need no live Excel. The tagged
/// <c>CopyEditableIntegrationTests</c> body drives the real adapters.
/// </remarks>
public class CopyEditableCommandTests
{
    private static readonly DateOnly _start = new(2026, 1, 5);
    private static readonly DateOnly _finish = new(2026, 1, 9);

    [Fact]
    public void Run_pins_the_pipeline_order()
    {
        Harness harness = new();

        _ = harness.Run();

        Assert.Equal(["read", "config", "measure", "export", "resolve", "lifecycle"], harness.Order);
    }

    [Fact]
    public void Run_feeds_the_validated_events_to_the_export_request()
    {
        Harness harness = new();
        harness.Rows = [ValidSpan(1), ValidSpan(2)];

        _ = harness.Run();

        Assert.NotNull(harness.SeenEvents);
        Assert.Equal(2, harness.SeenEvents.Count);
    }

    [Fact]
    public void Run_sends_table_plus_chart_and_the_panel_theme_by_default()
    {
        Harness harness = new();

        _ = harness.Run();

        Assert.Equal(ExportSelection.TablePlusChart, harness.SeenSelection);
        Assert.NotNull(harness.SeenPanelTheme);
    }

    [Fact]
    public void Run_sends_chart_only_and_no_panel_theme_for_an_explicit_false()
    {
        Harness harness = new();
        harness.Settings = new Dictionary<string, string>
        {
            [CopyEditableCommand.ExportIncludeDataPanelSettingKey] = "FALSE",
        };

        _ = harness.Run();

        Assert.Equal(ExportSelection.ChartOnly, harness.SeenSelection);
        Assert.Null(harness.SeenPanelTheme);
    }

    [Fact]
    public void Run_treats_a_malformed_inclusion_setting_as_table_plus_chart()
    {
        Harness harness = new();
        harness.Settings = new Dictionary<string, string>
        {
            [CopyEditableCommand.ExportIncludeDataPanelSettingKey] = "not-a-boolean",
        };

        _ = harness.Run();

        Assert.Equal(ExportSelection.TablePlusChart, harness.SeenSelection);
    }

    [Fact]
    public void Run_wires_the_verified_copy_seam_to_the_supplied_ports()
    {
        Harness harness = new();
        _ = harness.Copier
            .Setup(copier => copier.CopyStaged(It.IsAny<string>()))
            .Returns(ShapeWriteOutcome.Ok());
        _ = harness.Clipboard
            .Setup(reader => reader.Read())
            .Returns(ClipboardFormatRead.ForSuccess(new ClipboardFormatSnapshot(
                7,
                [.. StagedShapeCopier.RequiredDrawingFormatNames])));
        _ = harness.Run();

        StagedGroupOutcome group = new(
            "GanttCreatorStaging:export-group",
            ["panel-cell:0"],
            new RectD(0, 0, 100, 50),
            null,
            null);
        StagingCopyResult result = harness.SeenLifecycleRequest!.Copy!(group);

        Assert.True(result.Succeeded);
        harness.Copier.Verify(
            copier => copier.CopyStaged("GanttCreatorStaging:export-group"),
            Times.Once);
    }

    [Fact]
    public void Run_states_the_staged_shape_count_in_the_success_message()
    {
        Harness harness = new();
        harness.NextLifecycle = SuccessfulLifecycle(3);

        CopyEditableOutcome outcome = harness.Run();

        Assert.True(outcome.Succeeded);
        Assert.Equal(3, outcome.StagedShapeCount);
        Assert.Equal(0, outcome.DroppedRowCount);
        Assert.Null(outcome.Refusal);
        Assert.Equal(
            "Copied the editable Gantt (3 shapes) to the clipboard.",
            outcome.Message);
        Assert.Equal([outcome.Message], harness.Successes);
        Assert.Empty(harness.Failures);
    }

    [Fact]
    public void Run_uses_the_singular_shape_word_for_one_shape()
    {
        Harness harness = new();
        harness.NextLifecycle = SuccessfulLifecycle(1);

        CopyEditableOutcome outcome = harness.Run();

        Assert.Equal(
            "Copied the editable Gantt (1 shape) to the clipboard.",
            outcome.Message);
    }

    [Fact]
    public void Run_names_the_dropped_error_rows_in_the_message()
    {
        Harness harness = new();
        harness.Rows = [ValidSpan(1), MissingStart(2)];

        CopyEditableOutcome outcome = harness.Run();

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.DroppedRowCount);
        Assert.Equal(
            "Copied the editable Gantt (3 shapes) to the clipboard. "
                + "1 row(s) with errors were not copied — fix them and copy again.",
            outcome.Message);
        Assert.Contains(harness.Records, record => record.Contains("rows-dropped=1", StringComparison.Ordinal));
    }

    [Fact]
    public void Run_refuses_when_the_table_cannot_be_read()
    {
        Harness harness = new();
        harness.ReadOutcome = GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.TableMissing);

        CopyEditableOutcome outcome = harness.Run();

        Assert.False(outcome.Succeeded);
        Assert.Equal(CopyEditableRefusal.TableReadRefused, outcome.Refusal);
        Assert.Contains("Initialise sheet", Assert.Single(harness.Failures), StringComparison.Ordinal);
        Assert.Empty(harness.Successes);
        Assert.Equal(["read"], harness.Order);
    }

    [Fact]
    public void Run_refuses_when_the_configuration_cannot_be_read()
    {
        Harness harness = new();
        harness.ConfigOutcome = ConfigReadOutcome.Refused(ConfigReadRefusalReason.ConfigSheetMissing);

        CopyEditableOutcome outcome = harness.Run();

        Assert.False(outcome.Succeeded);
        Assert.Equal(CopyEditableRefusal.ConfigReadRefused, outcome.Refusal);
        Assert.Contains("Initialise sheet", Assert.Single(harness.Failures), StringComparison.Ordinal);
        Assert.Equal(["read", "config"], harness.Order);
    }

    [Fact]
    public void Run_refuses_when_the_panel_cannot_be_measured()
    {
        Harness harness = new();
        harness.GridOutcome = PanelGridOutcome.Refused(PanelGridRefusalReason.InvalidMeasurement);

        CopyEditableOutcome outcome = harness.Run();

        Assert.False(outcome.Succeeded);
        Assert.Equal(CopyEditableRefusal.MeasurementRefused, outcome.Refusal);
        Assert.Contains("could not be measured", Assert.Single(harness.Failures), StringComparison.Ordinal);
        Assert.Equal(["read", "config", "measure"], harness.Order);
    }

    [Fact]
    public void Run_presents_the_export_request_refusal_message()
    {
        Harness harness = new();
        harness.NextExport = SceneBuildRequestOutcome.Refused(
            SceneBuildRequestRefusal.NoPlotRange,
            "The table has no dates to plot.");

        CopyEditableOutcome outcome = harness.Run();

        Assert.False(outcome.Succeeded);
        Assert.Equal(CopyEditableRefusal.ExportRequestRefused, outcome.Refusal);
        Assert.Equal("The table has no dates to plot.", Assert.Single(harness.Failures));
    }

    [Fact]
    public void Run_presents_the_empty_visible_table_message_verbatim()
    {
        Harness harness = new();
        harness.NextBounds = new ExportBoundsOutcome(null, ExportBoundsRefusal.EmptyVisibleTable, null);

        CopyEditableOutcome outcome = harness.Run();

        Assert.False(outcome.Succeeded);
        Assert.Equal(CopyEditableRefusal.BoundsRefused, outcome.Refusal);
        Assert.Equal(ExportBoundsRefusalMessages.EmptyVisibleTable, Assert.Single(harness.Failures));
    }

    [Fact]
    public void Run_presents_the_verified_copys_own_detail_for_a_copy_refusal()
    {
        Harness harness = new();
        const string detail =
            "The clipboard was busy or unavailable (clipboard-unavailable:0x00000005). "
                + "Close the application holding the clipboard and try again.";
        harness.NextLifecycle = new StagingLifecycleOutcome(
            null,
            null,
            new StagingCopyResult(false, detail, StagedCopyRefusal.ClipboardUnavailable),
            StagingLifecycleRefusal.CopyRefused,
            []);

        CopyEditableOutcome outcome = harness.Run();

        Assert.False(outcome.Succeeded);
        Assert.Equal(CopyEditableRefusal.CopyRefused, outcome.Refusal);
        Assert.Equal(detail, Assert.Single(harness.Failures));
    }

    [Fact]
    public void Run_reports_cleanup_residue_as_a_failure()
    {
        Harness harness = new();
        harness.NextLifecycle = new StagingLifecycleOutcome(
            null,
            null,
            null,
            StagingLifecycleRefusal.CleanupResidue,
            ["panel-cell:0"]);

        CopyEditableOutcome outcome = harness.Run();

        Assert.False(outcome.Succeeded);
        Assert.Equal(CopyEditableRefusal.CleanupResidue, outcome.Refusal);
        Assert.Contains("temporary shapes", Assert.Single(harness.Failures), StringComparison.Ordinal);
        Assert.Empty(harness.Successes);
    }

    [Fact]
    public void Run_does_not_report_success_when_the_copy_seam_never_ran()
    {
        Harness harness = new();
        harness.NextLifecycle = new StagingLifecycleOutcome(null, null, null, null, []);

        CopyEditableOutcome outcome = harness.Run();

        Assert.False(outcome.Succeeded);
        Assert.Equal(CopyEditableRefusal.NothingCopied, outcome.Refusal);
        Assert.Empty(harness.Successes);
    }

    [Fact]
    public void Run_never_lets_a_throwing_presenter_escape()
    {
        Harness harness = new();

        CopyEditableOutcome outcome = CopyEditableCommand.Run(
            harness.Build(),
            _ => throw new InvalidOperationException("presenter failed"),
            _ => { });

        Assert.True(outcome.Succeeded);
    }

    [Theory]
    [InlineData("TableReader", nameof(CopyEditableRefusal.MissingTableReader))]
    [InlineData("ConfigReader", nameof(CopyEditableRefusal.MissingConfigReader))]
    [InlineData("GridMeasurement", nameof(CopyEditableRefusal.MissingGridMeasurement))]
    [InlineData("ExportRequestFactory", nameof(CopyEditableRefusal.MissingExportRequestFactory))]
    [InlineData("BoundsResolver", nameof(CopyEditableRefusal.MissingBoundsResolver))]
    [InlineData("Lifecycle", nameof(CopyEditableRefusal.MissingLifecycle))]
    [InlineData("StateScope", nameof(CopyEditableRefusal.MissingStateScope))]
    [InlineData("StagingWriter", nameof(CopyEditableRefusal.MissingStagingWriter))]
    [InlineData("GroupingPort", nameof(CopyEditableRefusal.MissingGroupingPort))]
    [InlineData("CopyPort", nameof(CopyEditableRefusal.MissingCopyPort))]
    [InlineData("ClipboardReader", nameof(CopyEditableRefusal.MissingClipboardReader))]
    public void Run_refuses_each_missing_dependency(string member, string expectedRefusal)
    {
        CopyEditableRefusal expected = Enum.Parse<CopyEditableRefusal>(expectedRefusal);
        Harness harness = new();
        CopyEditableRequest request = member switch
        {
            "TableReader" => harness.Build() with { TableReader = null },
            "ConfigReader" => harness.Build() with { ConfigReader = null },
            "GridMeasurement" => harness.Build() with { GridMeasurement = null },
            "ExportRequestFactory" => harness.Build() with { ExportRequestFactory = null },
            "BoundsResolver" => harness.Build() with { BoundsResolver = null },
            "Lifecycle" => harness.Build() with { Lifecycle = null },
            "StateScope" => harness.Build() with { StateScope = null },
            "StagingWriter" => harness.Build() with { StagingWriter = null },
            "GroupingPort" => harness.Build() with { GroupingPort = null },
            "CopyPort" => harness.Build() with { CopyPort = null },
            "ClipboardReader" => harness.Build() with { ClipboardReader = null },
            _ => throw new ArgumentOutOfRangeException(nameof(member), member, "Unknown dependency."),
        };

        CopyEditableOutcome outcome = CopyEditableCommand.Run(
            request,
            harness.Successes.Add,
            harness.Failures.Add);

        Assert.False(outcome.Succeeded);
        Assert.Equal(expected, outcome.Refusal);
        Assert.Equal(CopyEditableCommand.ConstructionMessage, Assert.Single(harness.Failures));
        Assert.Empty(harness.Successes);
    }

    [Fact]
    public void Run_refuses_a_null_request()
    {
        List<string> failures = [];

        CopyEditableOutcome outcome = CopyEditableCommand.Run(null, _ => { }, failures.Add);

        Assert.False(outcome.Succeeded);
        Assert.Equal(CopyEditableRefusal.NullRequest, outcome.Refusal);
        Assert.Equal(CopyEditableCommand.ConstructionMessage, Assert.Single(failures));
    }

    [Fact]
    public void SelectSuccessPresenter_uses_the_status_bar_when_visible()
    {
        List<string> statusBar = [];
        List<string> dialog = [];

        Action<string> presenter = CopyEditableCommand.SelectSuccessPresenter(
            true,
            statusBar.Add,
            dialog.Add);
        presenter("copied");

        Assert.Equal(["copied"], statusBar);
        Assert.Empty(dialog);
    }

    [Fact]
    public void SelectSuccessPresenter_uses_the_dialog_when_the_status_bar_is_hidden()
    {
        List<string> statusBar = [];
        List<string> dialog = [];

        Action<string> presenter = CopyEditableCommand.SelectSuccessPresenter(
            false,
            statusBar.Add,
            dialog.Add);
        presenter("copied");

        Assert.Equal(["copied"], dialog);
        Assert.Empty(statusBar);
    }

    [Fact]
    public void SelectSuccessPresenter_treats_unknown_visibility_as_visible()
    {
        List<string> statusBar = [];
        List<string> dialog = [];

        Action<string> presenter = CopyEditableCommand.SelectSuccessPresenter(
            null,
            statusBar.Add,
            dialog.Add);
        presenter("copied");

        Assert.Equal(["copied"], statusBar);
        Assert.Empty(dialog);
    }

    private static GanttRowDto ValidSpan(int rowNumber) =>
        new(
            rowNumber,
            GanttRowId.New().Value,
            null,
            null,
            "As-Planned Activity",
            "Build",
            _start,
            _finish,
            null,
            null,
            null,
            null,
            null,
            null,
            null);

    private static GanttRowDto MissingStart(int rowNumber) =>
        new(
            rowNumber,
            GanttRowId.New().Value,
            null,
            null,
            "As-Planned Activity",
            "Broken",
            null,
            _finish,
            null,
            null,
            null,
            null,
            null,
            null,
            null);

    private static PanelCellGrid Grid() =>
        PanelCellGrid.TryCreate(
            [new PanelColumn("Type", 60.0)],
            [15.0],
            15.0,
            null).Grid!;

    private static ExportBoundsOutcome OkBounds() =>
        new(new ExportBounds(new RectD(0, 0, 100, 50)), null, null);

    private static StagingLifecycleOutcome SuccessfulLifecycle(int stagedCount)
    {
        string[] ids = [.. Enumerable.Range(0, stagedCount).Select(index => $"panel-cell:{index}")];
        RectD bounds = new(0, 0, 100, 50);
        return new StagingLifecycleOutcome(
            new StagingCompositionOutcome(ids, bounds, null, null, null),
            new StagedGroupOutcome("GanttCreatorStaging:export-group", ids, bounds, null, null),
            StagingCopyResult.Success,
            null,
            []);
    }

    /// <summary>
    /// The deterministic fakes and captured observations one fact drives. <see cref="Build"/>
    /// wires every seam in pipeline order, so the order the command actually calls them is the
    /// order <see cref="Order"/> records.
    /// </summary>
    private sealed class Harness
    {
        public List<string> Order { get; } = [];

        public List<string> Successes { get; } = [];

        public List<string> Failures { get; } = [];

        public List<string> Records { get; } = [];

        public Mock<IGanttTableReader> TableReader { get; } = new();

        public Mock<IConfigCatalogueReader> ConfigReader { get; } = new();

        public Mock<IPanelGridMeasurementPort> Measurement { get; } = new();

        public Mock<IApplicationStateScope> Scope { get; } = new();

        public Mock<IShapeWritePort> Writer { get; } = new();

        public Mock<IShapeGroupingPort> Grouper { get; } = new();

        public Mock<IShapeCopyPort> Copier { get; } = new();

        public Mock<IClipboardFormatReader> Clipboard { get; } = new();

        public IReadOnlyList<GanttRowDto> Rows { get; set; } = [ValidSpan(1)];

        public IReadOnlyDictionary<string, string> Settings { get; set; } =
            new Dictionary<string, string>();

        public GanttTableReadOutcome? ReadOutcome { get; set; }

        public ConfigReadOutcome? ConfigOutcome { get; set; }

        public PanelGridOutcome? GridOutcome { get; set; }

        public SceneBuildRequestOutcome NextExport { get; set; } =
            SceneBuildRequestOutcome.Ok(new SceneBuildRequest());

        public ExportBoundsOutcome NextBounds { get; set; } = OkBounds();

        public StagingLifecycleOutcome NextLifecycle { get; set; } = SuccessfulLifecycle(3);

        public IReadOnlyList<GanttEvent>? SeenEvents { get; private set; }

        public ExportSelection? SeenSelection { get; private set; }

        public PanelTheme? SeenPanelTheme { get; private set; }

        public StagingLifecycleRequest? SeenLifecycleRequest { get; private set; }

        public CopyEditableRequest Build()
        {
            _ = TableReader
                .Setup(reader => reader.Read())
                .Callback(() => Order.Add("read"))
                .Returns(() => ReadOutcome ?? GanttTableReadOutcome.Ok(Rows));
            _ = ConfigReader
                .Setup(reader => reader.Read())
                .Callback(() => Order.Add("config"))
                .Returns(() => ConfigOutcome ?? ConfigReadOutcome.Ok(
                    "workbook-id",
                    Settings,
                    GanttStyleRegistry.Empty,
                    new Dictionary<string, double>()));
            _ = Measurement
                .Setup(port => port.Measure(
                    It.IsAny<IReadOnlyList<string>>(),
                    It.IsAny<IReadOnlyList<double>?>()))
                .Callback(() => Order.Add("measure"))
                .Returns(() => GridOutcome ?? PanelGridOutcome.Ok(Grid()));

            return new CopyEditableRequest(
                TableReader.Object,
                ConfigReader.Object,
                Measurement.Object,
                (events, settings, registry, grid, metrics) =>
                {
                    Order.Add("export");
                    SeenEvents = events;
                    return NextExport;
                },
                (request, panelTheme, selection) =>
                {
                    Order.Add("resolve");
                    SeenSelection = selection;
                    SeenPanelTheme = panelTheme;
                    return NextBounds;
                },
                request =>
                {
                    Order.Add("lifecycle");
                    SeenLifecycleRequest = request;
                    return NextLifecycle;
                },
                Scope.Object,
                Writer.Object,
                Grouper.Object,
                Copier.Object,
                Clipboard.Object,
                Records.Add);
        }

        public CopyEditableOutcome Run(CopyEditableRequest? request = null) =>
            CopyEditableCommand.Run(request ?? Build(), Successes.Add, Failures.Add);
    }
}
