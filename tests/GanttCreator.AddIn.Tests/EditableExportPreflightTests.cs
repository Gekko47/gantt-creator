using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// R7.6 D4: the shared editable-export pre-flight is the one authority for the
/// read → configuration → validate → measure → export request → selection → bounds chain,
/// including the post-resolution guard. These facts pin its order, its typed refusals, the
/// guard positives, and the moved export-selection and inset rules.
/// </summary>
public class EditableExportPreflightTests
{
    private static readonly DateOnly _start = new(2026, 1, 5);
    private static readonly DateOnly _finish = new(2026, 1, 9);

    [Fact]
    public void Run_pins_the_read_config_measure_export_resolve_order()
    {
        Harness harness = new();

        _ = harness.Run();

        Assert.Equal(["read", "config", "measure", "export", "resolve"], harness.Order);
    }

    /// <summary>
    /// The R7.6 Step 0 guard: a resolved bounds with a zero extent is refused with the
    /// existing <see cref="ExportBoundsRefusal.InvalidBounds"/> before any staging shape is
    /// created. <see cref="RectD"/> permits a zero extent, which is the reachable arm.
    /// </summary>
    [Fact]
    public void Run_refuses_a_zero_width_resolved_bounds_with_InvalidBounds()
    {
        Harness harness = new()
        {
            NextExport = SceneBuildRequestOutcome.Ok(new SceneBuildRequest()),
            NextBounds = new ExportBoundsOutcome(new ExportBounds(new RectD(0, 0, 0, 50)), null, null),
        };

        EditableExportPreflightOutcome outcome = harness.Run();

        Assert.False(outcome.Succeeded);
        Assert.Equal(EditableExportPreflightRefusal.BoundsRefused, outcome.Refusal);
        Assert.NotNull(outcome.BoundsOutcome);
        Assert.Equal(ExportBoundsRefusal.InvalidBounds, outcome.BoundsOutcome.Refusal);
    }

    [Fact]
    public void Run_refuses_a_zero_height_resolved_bounds_with_InvalidBounds()
    {
        Harness harness = new()
        {
            NextBounds = new ExportBoundsOutcome(new ExportBounds(new RectD(0, 0, 100, 0)), null, null),
        };

        EditableExportPreflightOutcome outcome = harness.Run();

        Assert.Equal(EditableExportPreflightRefusal.BoundsRefused, outcome.Refusal);
        Assert.Equal(ExportBoundsRefusal.InvalidBounds, outcome.BoundsOutcome?.Refusal);
    }

    /// <summary>The positive control: a usable extent passes the guard untouched.</summary>
    [Fact]
    public void Run_accepts_a_usable_resolved_bounds()
    {
        Harness harness = new();

        EditableExportPreflightOutcome outcome = harness.Run();

        Assert.True(outcome.Succeeded);
        Assert.Equal(100, outcome.Preparation!.Resolution.Bounds!.Bounds.Width);
        Assert.Equal(50, outcome.Preparation.Resolution.Bounds.Bounds.Height);
    }

    [Theory]
    [InlineData("TableReader", nameof(EditableExportPreflightRefusal.MissingTableReader))]
    [InlineData("ConfigReader", nameof(EditableExportPreflightRefusal.MissingConfigReader))]
    [InlineData("GridMeasurement", nameof(EditableExportPreflightRefusal.MissingGridMeasurement))]
    [InlineData("ExportRequestFactory", nameof(EditableExportPreflightRefusal.MissingExportRequestFactory))]
    [InlineData("BoundsResolver", nameof(EditableExportPreflightRefusal.MissingBoundsResolver))]
    public void Run_refuses_each_missing_dependency(string member, string expectedRefusal)
    {
        EditableExportPreflightRefusal expected = Enum.Parse<EditableExportPreflightRefusal>(expectedRefusal);
        Harness harness = new();
        EditableExportPreflightRequest request = member switch
        {
            "TableReader" => harness.Build() with { TableReader = null },
            "ConfigReader" => harness.Build() with { ConfigReader = null },
            "GridMeasurement" => harness.Build() with { GridMeasurement = null },
            "ExportRequestFactory" => harness.Build() with { ExportRequestFactory = null },
            "BoundsResolver" => harness.Build() with { BoundsResolver = null },
            _ => throw new ArgumentOutOfRangeException(nameof(member), member, "Unknown dependency."),
        };

        EditableExportPreflightOutcome outcome = EditableExportPreflight.Run(request);

        Assert.False(outcome.Succeeded);
        Assert.Equal(expected, outcome.Refusal);
    }

    [Fact]
    public void Run_refuses_a_null_request()
    {
        EditableExportPreflightOutcome outcome = EditableExportPreflight.Run(null);

        Assert.Equal(EditableExportPreflightRefusal.NullRequest, outcome.Refusal);
    }

    [Fact]
    public void Run_carries_the_table_read_refusal_reason()
    {
        Harness harness = new()
        {
            ReadOutcome = GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.TableMissing),
        };

        EditableExportPreflightOutcome outcome = harness.Run();

        Assert.Equal(EditableExportPreflightRefusal.TableReadRefused, outcome.Refusal);
        Assert.Equal(GanttTableReadRefusalReason.TableMissing, outcome.TableReadRefusal);
        Assert.Equal(["read"], harness.Order);
    }

    [Fact]
    public void Run_carries_the_configuration_refusal_reason()
    {
        Harness harness = new()
        {
            ConfigOutcome = ConfigReadOutcome.Refused(ConfigReadRefusalReason.ConfigSheetMissing),
        };

        EditableExportPreflightOutcome outcome = harness.Run();

        Assert.Equal(EditableExportPreflightRefusal.ConfigReadRefused, outcome.Refusal);
        Assert.Equal(ConfigReadRefusalReason.ConfigSheetMissing, outcome.ConfigRefusal);
        Assert.Equal(["read", "config"], harness.Order);
    }

    [Fact]
    public void Run_carries_the_measurement_refusal_reason_and_the_dropped_rows()
    {
        Harness harness = new()
        {
            Rows = [ValidSpan(1), MissingStart(2)],
            GridOutcome = PanelGridOutcome.Refused(PanelGridRefusalReason.InvalidMeasurement),
        };

        EditableExportPreflightOutcome outcome = harness.Run();

        Assert.Equal(EditableExportPreflightRefusal.MeasurementRefused, outcome.Refusal);
        Assert.Equal(PanelGridRefusalReason.InvalidMeasurement, outcome.MeasurementRefusal);
        Assert.Equal(1, outcome.DroppedRowCount);
        Assert.Contains(harness.Records, record => record.Contains("rows-dropped=1", StringComparison.Ordinal));
    }

    [Fact]
    public void Run_carries_the_export_request_message()
    {
        Harness harness = new()
        {
            NextExport = SceneBuildRequestOutcome.Refused(
                SceneBuildRequestRefusal.NoPlotRange,
                "The table has no dates to plot."),
        };

        EditableExportPreflightOutcome outcome = harness.Run();

        Assert.Equal(EditableExportPreflightRefusal.ExportRequestRefused, outcome.Refusal);
        Assert.Equal("The table has no dates to plot.", outcome.ExportRequestMessage);
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
        Harness harness = new()
        {
            Settings = new Dictionary<string, string>
            {
                [EditableExportPreflight.ExportIncludeDataPanelSettingKey] = "FALSE",
            },
        };

        _ = harness.Run();

        Assert.Equal(ExportSelection.ChartOnly, harness.SeenSelection);
        Assert.Null(harness.SeenPanelTheme);
    }

    [Fact]
    public void Run_resolves_the_inset_from_the_outer_padding_token()
    {
        Harness harness = new()
        {
            Metrics = new Dictionary<string, double>
            {
                [EditableExportPreflight.OuterPaddingToken] = 7.5,
            },
        };

        EditableExportPreflightOutcome outcome = harness.Run();

        Assert.Equal(7.5, outcome.Preparation!.InsetPt);
    }

    [Fact]
    public void Run_falls_back_to_the_catalogue_inset_when_the_token_is_absent()
    {
        Harness harness = new();

        EditableExportPreflightOutcome outcome = harness.Run();

        Assert.Equal(
            GanttCatalogues.MetricDefault(EditableExportPreflight.OuterPaddingToken),
            outcome.Preparation!.InsetPt);
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

    private sealed class Harness
    {
        public List<string> Order { get; } = [];

        public List<string> Records { get; } = [];

        public IReadOnlyList<GanttRowDto> Rows { get; set; } = [ValidSpan(1)];

        public IReadOnlyDictionary<string, string> Settings { get; set; } =
            new Dictionary<string, string>();

        public IReadOnlyDictionary<string, double> Metrics { get; set; } =
            new Dictionary<string, double>();

        public GanttTableReadOutcome? ReadOutcome { get; set; }

        public ConfigReadOutcome? ConfigOutcome { get; set; }

        public PanelGridOutcome? GridOutcome { get; set; }

        public SceneBuildRequestOutcome NextExport { get; set; } =
            SceneBuildRequestOutcome.Ok(new SceneBuildRequest());

        public ExportBoundsOutcome NextBounds { get; set; } =
            new(new ExportBounds(new RectD(0, 0, 100, 50)), null, null);

        public ExportSelection? SeenSelection { get; private set; }

        public PanelTheme? SeenPanelTheme { get; private set; }

        public EditableExportPreflightRequest Build()
        {
            FakeTableReader reader = new(this);
            FakeConfigReader config = new(this);
            FakeMeasurement measurement = new(this);

            return new EditableExportPreflightRequest(
                reader,
                config,
                measurement,
                (events, settings, registry, grid, metrics) =>
                {
                    Order.Add("export");
                    return NextExport;
                },
                (request, panelTheme, selection) =>
                {
                    Order.Add("resolve");
                    SeenSelection = selection;
                    SeenPanelTheme = panelTheme;
                    return NextBounds;
                },
                Records.Add);
        }

        public EditableExportPreflightOutcome Run() => EditableExportPreflight.Run(Build());

        private sealed class FakeTableReader(Harness harness) : IGanttTableReader
        {
            public GanttTableReadOutcome Read()
            {
                harness.Order.Add("read");
                return harness.ReadOutcome ?? GanttTableReadOutcome.Ok(harness.Rows);
            }

            public GanttTableReadOutcome Read(IReadOnlyList<int> rowNumbers) => Read();
        }

        private sealed class FakeConfigReader(Harness harness) : IConfigCatalogueReader
        {
            public ConfigSchemaObservation? ObserveSchema() => null;

            public ConfigReadOutcome Read()
            {
                harness.Order.Add("config");
                return harness.ConfigOutcome ?? ConfigReadOutcome.Ok(
                    "workbook-id",
                    harness.Settings,
                    GanttStyleRegistry.Empty,
                    harness.Metrics);
            }
        }

        private sealed class FakeMeasurement(Harness harness) : IPanelGridMeasurementPort
        {
            public PanelGridOutcome Measure(
                IReadOnlyList<string> columns,
                IReadOnlyList<double>? rowHeights = null)
            {
                harness.Order.Add("measure");
                return harness.GridOutcome ?? PanelGridOutcome.Ok(Grid());
            }
        }
    }
}
