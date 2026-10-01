using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R4.8A D5: the scene-request factory is the single authority for every scene
/// input, and the profile it emits is the one D4 requires.
/// </summary>
/// <remarks>
/// No Excel: the factory takes four values and returns a request, so every claim here
/// is about the translation itself rather than about a host.
/// </remarks>
public class ExcelSceneBuildRequestFactoryTests
{
    private static GanttEvent Event(DateOnly start, DateOnly finish) =>
        new(
            RowNumber: 1,
            Id: GanttRowId.New(),
            LaneId: GanttRowId.New(),
            StackIndex: 0,
            Type: GanttEntityType.AsPlannedActivity,
            Description: "Design",
            Start: start,
            Finish: finish,
            ParentId: null,
            StyleKey: "AsPlannedActivity",
            LabelPosition: null,
            FillColour: null,
            StrokeColour: null,
            Visible: true,
            SortOrder: 0);

    /// <summary>
    /// A measured grid narrow enough to leave a readable A4-landscape plot.
    /// </summary>
    /// <param name="widthPt">The panel's total measured width.</param>
    /// <param name="originTopPt">The first body row's absolute top edge (ADR-0030 D3).</param>
    /// <param name="rowHeightsPt">The measured body row heights.</param>
    private static PanelCellGrid Grid(
        double widthPt = 120,
        double originTopPt = 58,
        double[]? rowHeightsPt = null) =>
        PanelCellGrid.TryCreate(
            [new PanelColumn("Id", widthPt)],
            rowHeightsPt ?? [15.0],
            15,
            ["Id"],
            originTopPt).Grid!;

    private static FakeTextMetrics Metrics() => new();

    /// <summary>
    /// The measured set is every VISIBLE schema column, and it is derived from the
    /// schema rather than restated in the factory.
    /// </summary>
    /// <remarks>
    /// This is the regression test for the live "The worksheet columns could not be
    /// measured" failure. <c>MeasuredColumns</c> was the literal
    /// <c>["Id", "Type", "Description"]</c> and <c>Id</c> is an <c>EngineHidden</c>
    /// column: step 4 of the refresh hides it, Excel reports a hidden column's
    /// <c>Range.Width</c> as <c>0</c>, and <c>PanelCellGrid.TryCreate</c> refuses a
    /// non-positive width. Every live refresh therefore refused at
    /// <c>MeasurementRefused</c>.
    /// </remarks>
    [Fact]
    public void Measured_columns_are_exactly_the_visible_schema_columns_in_schema_order()
    {
        Assert.Equal(
            GanttTableSchema.Default.Columns
                .Where(column => !column.IsHidden)
                .Select(column => column.Name),
            ExcelSceneBuildRequestFactory.MeasuredColumns);
    }

    /// <summary>
    /// The positive assertion of the fix, and the one that would have caught it:
    /// <b>no measured column may be a hidden column.</b>
    /// </summary>
    /// <remarks>
    /// Asserting the exact list alone would still pass if the schema were changed to
    /// hide <c>Type</c>, so the invariant is asserted directly as well. This is the
    /// pairing the failure needed: the check that a hidden column is measured at all,
    /// which is the condition that produces a zero width and the refusal.
    /// </remarks>
    [Fact]
    public void Measured_columns_contain_no_hidden_column_because_a_hidden_one_measures_zero()
    {
        IReadOnlyList<string> measured = ExcelSceneBuildRequestFactory.MeasuredColumns;

        List<string> hidden = [.. GanttTableSchema.Default.Columns.Where(c => c.IsHidden).Select(c => c.Name)];

        Assert.DoesNotContain(measured, name => hidden.Contains(name));
    }

    /// <summary>
    /// The measured set is not empty, so the measurement is never refused for want of
    /// columns.
    /// </summary>
    /// <remarks>
    /// <c>Measure</c> refuses an empty column list outright, so a schema change that
    /// hid every column would otherwise surface as the same user-facing
    /// "columns could not be measured" message with a different cause.
    /// </remarks>
    [Fact]
    public void Measured_columns_are_not_empty_because_an_empty_set_is_refused_outright()
    {
        Assert.NotEmpty(ExcelSceneBuildRequestFactory.MeasuredColumns);
    }

    /// <summary>
    /// The measured panel width is the sum of the visible columns, so entity guide
    /// section 3's "the panel right edge touches the plot left edge" holds. The plot's
    /// left edge is <c>textPanelWidthPt + chrome</c>.
    /// </summary>
    /// <remarks>
    /// Measuring only the label columns would leave the plot drawn on top of the
    /// still-visible <c>Start</c>, <c>Finish</c>, and <c>Duration</c> columns, so this
    /// asserts the placement rather than merely the list.
    /// </remarks>
    [Fact]
    public void The_plot_left_edge_starts_after_the_whole_measured_panel_plus_chrome()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string>(),
            StyleRegistry(),
            Grid(widthPt: 120));

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);

        SceneBuildRequest request = Assert.IsType<SceneBuildRequest>(outcome.Request);
        double chrome = GanttCatalogues.MetricDefault("ChartOuterPaddingPt");

        // Both `Grid` and `PlotBounds` are nullable on the request and are unwrapped
        // with the `!.Value`/`!` form the rest of this file already uses. The factory
        // only emits a request whose grid and plot both resolved, so a null here is a
        // defect the equality assertion reports as a mismatch rather than a null crash.
        Assert.Equal(request.Grid!.TotalWidthPt + chrome, request.PlotBounds!.Value.X, precision: 6);
    }

    /// <summary>
    /// The plot's TOP is the first body row's measured top edge (ADR-0030 D1).
    /// </summary>
    /// <remarks>
    /// <b>This is the test for the reported defect.</b> The plot's top used to be a
    /// fixed token sum — <c>TitleBand + YearBand + PeriodBand</c> = 58pt — regardless
    /// of the worksheet, which put lane 0 roughly 43pt (2.4 rows) below its own row.
    /// The value here is deliberately neither 58 nor zero, so neither the old page
    /// origin nor a default can satisfy it.
    /// </remarks>
    /// <summary>
    /// Builds a request over one event and the supplied measured grid.
    /// </summary>
    /// <param name="grid">The measured grid the scene is composed against.</param>
    /// <returns>The factory outcome.</returns>
    private static SceneBuildRequestOutcome Create(PanelCellGrid grid)
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());
        return factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string>(),
            StyleRegistry(),
            grid);
    }

    [Fact]
    public void The_plot_top_is_the_first_body_rows_measured_top()
    {
        SceneBuildRequestOutcome outcome = Create(
            Grid(originTopPt: 137.5, rowHeightsPt: [15.0, 15.0, 15.0]));

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);

        SceneBuildRequest request = Assert.IsType<SceneBuildRequest>(outcome.Request);
        Assert.Equal(137.5, request.PlotBounds!.Value.Y, precision: 6);
    }

    /// <summary>
    /// The plot's HEIGHT is the measured body total, so the chart ends with the table
    /// (ADR-0030 D2).
    /// </summary>
    /// <remarks>
    /// Previously the height was the remaining <em>page</em> height (~531pt on A4
    /// landscape), so a 3-row chart ran ~400pt past the last row. The sum here is
    /// 45pt, and it is asserted rather than any page figure.
    /// </remarks>
    [Fact]
    public void The_plot_height_is_the_measured_body_total_not_the_page_height()
    {
        SceneBuildRequestOutcome outcome = Create(
            Grid(rowHeightsPt: [15.0, 15.0, 15.0]));

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);

        SceneBuildRequest request = Assert.IsType<SceneBuildRequest>(outcome.Request);
        RectD plot = request.PlotBounds!.Value;

        Assert.Equal(45.0, plot.Height, precision: 6);
        Assert.Equal(45.0, request.Grid!.TotalRowHeightPt, precision: 6);
    }

    /// <summary>
    /// The plot's bottom edge is the last row's bottom edge.
    /// </summary>
    /// <remarks>
    /// Stated as a relationship rather than a number because it is the property the
    /// user sees: the chart must stop where the table stops. Before ADR-0030 the chart
    /// ran far past it, so nothing asserted this.
    /// </remarks>
    [Fact]
    public void The_plot_bottom_coincides_with_the_last_rows_bottom()
    {
        SceneBuildRequestOutcome outcome = Create(
            Grid(originTopPt: 64, rowHeightsPt: [18.0, 18.0, 24.0, 18.0]));

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);

        SceneBuildRequest request = Assert.IsType<SceneBuildRequest>(outcome.Request);
        RectD plot = request.PlotBounds!.Value;

        Assert.Equal(64.0 + 18.0 + 18.0 + 24.0 + 18.0, plot.Y + plot.Height, precision: 6);
    }

    /// <summary>
    /// A table taller than the preset's page still renders (ADR-0030 D7).
    /// </summary>
    /// <remarks>
    /// <b>The counterweight to the height test above.</b> With the page still bounding
    /// the vertical extent, a 40-row schedule on A4 landscape (29 rows tall) would be
    /// refused outright — the same defect as lane auto-growth, wearing a page budget
    /// instead of a row height. The width budget is untouched, so an over-wide panel
    /// still refuses.
    /// </remarks>
    [Fact]
    public void A_table_taller_than_the_page_is_not_refused()
    {
        double[] fortyRows = [.. Enumerable.Repeat(18.0, 40)];

        SceneBuildRequestOutcome outcome = Create(Grid(rowHeightsPt: fortyRows));

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);

        SceneBuildRequest request = Assert.IsType<SceneBuildRequest>(outcome.Request);
        Assert.Equal(40 * 18.0, request.PlotBounds!.Value.Height, precision: 6);
    }

    /// <summary>
    /// A panel too wide to leave a readable plot is still refused (R4.7H D4).
    /// </summary>
    /// <remarks>
    /// The counterweight that D7 did not loosen: dropping the page bound is about the
    /// <em>vertical</em> extent only. The width budget still refuses a panel that
    /// would leave an unreadable plot, so "the page no longer bounds it" cannot be
    /// read as "the page no longer constrains anything".
    /// </remarks>
    [Fact]
    public void An_over_wide_panel_is_still_refused()
    {
        // The default preset is A4 landscape (841.89pt wide); at 830 the panel plus
        // both chromes leaves under 1pt of plot, below MinimumPlotWidthPt (24).
        SceneBuildRequestOutcome outcome = Create(Grid(widthPt: 830));

        Assert.False(outcome.Succeeded);
    }

    [Fact]
    public void The_live_profile_is_emitted_and_no_data_panel_is_supplied()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string>(),
            StyleRegistry(),
            Grid());

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);

        // D4. Chosen in code, and the panel is null because the worksheet's own cells
        // ARE the panel. A drawn replica would cover the user's rows.
        Assert.Equal(SceneCompositionProfile.LiveExcel, outcome.Request!.Profile);
        Assert.Null(outcome.Request.Panel);
        Assert.False(SceneCompositionProfiles.DrawsDataPanel(outcome.Request.Profile));
    }

    /// <summary>
    /// The request the factory produces is accepted by the real scene builder.
    /// </summary>
    /// <remarks>
    /// This is the test that makes the factory useful rather than merely plausible. A
    /// factory that returned a request missing Metrics or LaneMetrics would satisfy
    /// every field-level assertion above and still be unusable in production, because
    /// the scene builder would refuse it.
    /// </remarks>
    [Fact]
    public void The_request_it_produces_is_accepted_by_the_scene_builder()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string>(),
            StyleRegistry(),
            Grid());

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);

        SceneBuildOutcome built = SceneBuilder.TryBuild(outcome.Request);

        Assert.True(built.Succeeded, "the produced request was refused: " + built.Refusal);
        Assert.NotEmpty(built.Result!.Scene.Primitives);
    }

    /// <summary>The plot bounds come from the resolver, not from a second derivation.</summary>
    [Fact]
    public void The_plot_bounds_are_the_resolver_output_for_the_preset()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());
        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string>(),
            StyleRegistry(),
            Grid(widthPt: 200));

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        Assert.Equal(SizePresets.Default, outcome.Request!.Preset);

        // A wider panel must leave a NARROWER plot on the same page, which is only
        // true if the plot is the remainder after the panel rather than a constant.
        // The resolver's rule is PlotWidth = presetWidth - panelWidth - chrome, so
        // the 200pt panel's plot is the narrower of the two.
        SceneBuildRequestOutcome widePanel = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string>(),
            StyleRegistry(),
            Grid(widthPt: 100));

        Assert.True(widePanel.Succeeded);
        Assert.True(
            outcome.Request!.PlotBounds!.Value.Width < widePanel.Request!.PlotBounds!.Value.Width,
            "the 200pt panel's plot ("
            + outcome.Request.PlotBounds.Value.Width
            + "pt) must be narrower than the 100pt panel's plot ("
            + widePanel.Request.PlotBounds.Value.Width
            + "pt)");
    }

    /// <summary>An unknown preset key refuses rather than silently falling back to A4.</summary>
    [Fact]
    public void An_unknown_size_preset_refuses_and_names_the_alternatives()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string> { ["SizePreset"] = "A3Gigantic" },
            StyleRegistry(),
            Grid());

        Assert.False(outcome.Succeeded);
        Assert.Equal(SceneBuildRequestRefusal.UnknownSizePreset, outcome.Refusal);

        // The message must be actionable: it names the key that was rejected and the
        // keys that would work.
        Assert.Contains("A3Gigantic", outcome.Message!, StringComparison.Ordinal);
        Assert.Contains("A4Portrait", outcome.Message!, StringComparison.Ordinal);
    }

    /// <summary>
    /// A known preset key is honoured <em>when the factory is handed one</em>.
    /// </summary>
    /// <remarks>
    /// The name used to read "a known preset key is honoured", which overstated it. This
    /// test injects the key into a hand-built dictionary, so it proves the factory
    /// parses a preset it is <em>given</em> — it does not prove the workbook can give it
    /// one, and today it cannot: <c>SizePreset</c> is absent from
    /// <c>GanttCatalogues.Settings</c>, so <c>ValidateSettings</c> never returns it. See
    /// <see cref="The_size_preset_keys_are_still_absent_from_the_catalogue"/> for that
    /// gap, which needs a schema bump and is tracked separately.
    /// </remarks>
    [Theory]
    [InlineData("A4Portrait")]
    [InlineData("A4Landscape")]
    [InlineData("Presentation16x9")]
    [InlineData("Presentation4x3")]
    public void A_known_preset_key_is_honoured(string key)
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string> { ["SizePreset"] = key },
            StyleRegistry(),
            Grid());

        Assert.True(outcome.Succeeded, key + " refused: " + outcome.Message);
        Assert.Equal(key, outcome.Request!.Preset!.Key);
    }

    /// <summary>
    /// The <c>SizePreset</c> key exists in the catalogue, and a value set against the
    /// catalogue's own key set reaches the factory.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This is the reachability test the injected-dictionary version of
    /// <c>A_known_preset_key_is_honoured</c> could never provide. That test proved the
    /// factory parses a preset it is <em>handed</em>; this one proves the catalogue
    /// actually hands it one, which is what schema version 7 fixed.
    /// </para>
    /// <para>
    /// The settings map is built from <c>GanttCatalogues.Settings</c> and only the
    /// <c>SizePreset</c> value is overridden per case — the rest stay at their catalogue
    /// defaults, exactly as a live workbook would read them. Before version 7 the key
    /// was absent, so every one of these cases was unreachable in production while the
    /// injected-dictionary test stayed green.
    /// </para>
    /// </remarks>
    [Theory]
    [InlineData("A4Portrait")]
    [InlineData("A4Landscape")]
    [InlineData("Presentation16x9")]
    [InlineData("Presentation4x3")]
    public void A_preset_selected_in_a_catalogue_shaped_map_reaches_the_factory(string key)
    {
        Dictionary<string, string> fromCatalogue = GanttCatalogues.Settings.ToDictionary(
            static setting => setting.Key,
            static setting => setting.DefaultValue,
            StringComparer.Ordinal);

        // The catalogue must actually offer this key, or the rest of the test is
        // vacuous — this is the assertion that would have caught the original gap.
        Assert.Contains("SizePreset", fromCatalogue.Keys);

        fromCatalogue["SizePreset"] = key;

        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            fromCatalogue,
            StyleRegistry(),
            Grid());

        Assert.True(outcome.Succeeded, key + " refused: " + outcome.Message);
        Assert.Equal(key, outcome.Request!.Preset!.Key);
    }

    /// <summary>
    /// The catalogue's own default preset is the one a freshly initialised workbook
    /// renders with.
    /// </summary>
    /// <remarks>
    /// Before schema 7 there was no such row, so every live chart was A4-portrait
    /// regardless of anything the user did. This pins the default now that a default
    /// exists, so retuning it is a deliberate change.
    /// </remarks>
    [Fact]
    public void The_catalogue_default_preset_reaches_the_factory()
    {
        Dictionary<string, string> fromCatalogue = GanttCatalogues.Settings.ToDictionary(
            static setting => setting.Key,
            static setting => setting.DefaultValue,
            StringComparer.Ordinal);

        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            fromCatalogue,
            StyleRegistry(),
            Grid());

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        Assert.Equal(
            fromCatalogue["SizePreset"],
            outcome.Request!.Preset!.Key);
    }

    /// <summary>A text panel too wide for the page refuses with the shortfall named.</summary>
    [Fact]
    public void A_text_panel_too_wide_for_the_preset_refuses_with_the_shortfall()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string>(),
            StyleRegistry(),
            Grid(widthPt: 5000));

        Assert.False(outcome.Succeeded);
        Assert.Equal(SceneBuildRequestRefusal.PlotGeometryRefused, outcome.Refusal);
        Assert.Contains("too wide", outcome.Message!, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// The plot range covers the events themselves, padded, rather than a stored
    /// range that could survive the data changing.
    /// </summary>
    [Fact]
    public void The_plot_range_is_derived_from_the_events_and_padded()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 3, 1), new DateOnly(2024, 3, 31))],
            new Dictionary<string, string> { ["RangePaddingDays"] = "2" },
            StyleRegistry(),
            Grid());

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        Assert.Equal(new DateOnly(2024, 2, 28), outcome.Request!.PlotStart);
        Assert.Equal(new DateOnly(2024, 4, 2), outcome.Request.PlotFinish);
    }

    /// <summary>
    /// Rows with no start date refuse rather than producing a chart over a range
    /// invented from the epoch.
    /// </summary>
    [Fact]
    public void No_event_with_a_date_refuses_rather_than_inventing_a_range()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [],
            new Dictionary<string, string>(),
            StyleRegistry(),
            Grid());

        Assert.False(outcome.Succeeded);
        Assert.Equal(SceneBuildRequestRefusal.NoPlotRange, outcome.Refusal);
    }

    /// <summary>A single-day span is widened rather than refused as unreadable.</summary>
    [Fact]
    public void A_one_day_span_is_widened_to_two_days_rather_than_refused()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 3, 1), new DateOnly(2024, 3, 1))],
            new Dictionary<string, string> { ["RangePaddingDays"] = "0" },
            StyleRegistry(),
            Grid());

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        Assert.True(outcome.Request!.PlotFinish > outcome.Request.PlotStart);
    }

    /// <summary>A malformed numeric setting falls back rather than refusing the refresh.</summary>
    /// <summary>
    /// A metric-looking key in the settings map cannot corrupt a metric, because
    /// metrics are resolved from the catalogue and never from the settings map.
    /// </summary>
    /// <remarks>
    /// This test previously asserted that a malformed <c>GridLinePt</c> setting fell
    /// back to 0.5. That contract is withdrawn: the key was never read in the first
    /// place, so the test was passing over dead code and gave false assurance that the
    /// settings path was live. The behaviour now asserted is the real one — a garbage
    /// value under a metric name is ignored, and the catalogue default survives.
    /// </remarks>
    [Fact]
    public void A_malformed_metric_setting_is_ignored_rather_than_applied()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string> { ["GridLinePt"] = "not-a-number" },
            StyleRegistry(),
            Grid());

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        Assert.Equal(GanttCatalogues.MetricDefault("GridLinePt"), outcome.Request!.GridLinePt);
    }

    /// <summary>
    /// Every metric the factory emits equals the code-owned catalogue default.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>This is the test that would have caught the dead-metric-token defect.</b> The
    /// factory previously read its metrics with <c>ReadDouble(settings, tokenName,
    /// fallback)</c>, but the settings map carries only the approved setting keys and
    /// no metric name is one of them, so every lookup missed and the divergent literal
    /// fallback was used instead — six of eight disagreed with the catalogue. Asserting
    /// the request against <see cref="GanttCatalogues.Metrics"/> states the contract
    /// that matters: the rendered chart uses the tokens the workbook publishes.
    /// </para>
    /// <para>
    /// It is written as a whole-catalogue sweep rather than as eight hand-picked
    /// assertions, so a token added to the request later is covered without anyone
    /// remembering to extend this list.
    /// </para>
    /// </remarks>
    [Fact]
    public void Every_emitted_metric_equals_the_catalogue_default()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string>(),
            StyleRegistry(),
            Grid());

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        SceneBuildRequest request = outcome.Request!;

        (string Token, double Actual)[] emitted =
        [
            ("ChartOuterPaddingPt", request.ChartOuterPaddingPt),
            ("TitleBandHeightPt", request.TitleBandHeightPt),
            ("YearBandHeightPt", request.YearBandHeightPt),
            ("PeriodBandHeightPt", request.PeriodBandHeightPt),
            ("MinimumHeaderLabelWidthPt", request.MinimumHeaderLabelWidthPt),
            ("GridLinePt", request.GridLinePt),
            ("MajorBoundaryPt", request.MajorBoundaryPt),
            ("MilestoneSizePt", request.MilestoneSizePt),
            ("DelineatorLinePt", request.DelineatorLinePt),
            ("LabelGapPt", request.LabelGapPt),
            ("LabelHeightPt", request.LabelHeightPt),
            ("LanePaddingTopPt", request.LaneMetrics!.LanePaddingTopPt),
            ("LanePaddingBottomPt", request.LaneMetrics.LanePaddingBottomPt),
            ("StackGapPt", request.LaneMetrics.StackGapPt),
            ("SplitterHeightPt", request.LaneMetrics.SplitterHeightPt),
            ("SpacerHeightPt", request.LaneMetrics.SpacerHeightPt),
        ];

        foreach ((string token, double actual) in emitted)
        {
            double expected = GanttCatalogues.MetricDefault(token);
            Assert.True(
                Math.Abs(expected - actual) < 1e-9,
                $"Metric '{token}' was emitted as {actual} but the catalogue default is {expected}.");
        }
    }

    /// <summary>
    /// A metric name is never read from the settings map, so editing a settings cell
    /// cannot move a metric the catalogue owns.
    /// </summary>
    /// <remarks>
    /// The defect was a lookup against the wrong table: the key was always absent, so
    /// the branch that consumed it was dead. This asserts the absence directly — the
    /// setting is supplied, and the emitted value is still the catalogue default. Under
    /// the old code this returned 0.5 from the fallback and would still have passed for
    /// <c>GridLinePt</c>, so the token chosen here is one whose old fallback differed
    /// from its catalogue default (<c>TitleBandHeightPt</c>: 14 against 24).
    /// </remarks>
    [Fact]
    public void A_metric_setting_is_not_read_from_the_settings_map()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string>
            {
                ["TitleBandHeightPt"] = "999",
                ["ChartOuterPaddingPt"] = "999",
            },
            StyleRegistry(),
            Grid());

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        Assert.Equal(GanttCatalogues.MetricDefault("TitleBandHeightPt"), outcome.Request!.TitleBandHeightPt);
        Assert.Equal(GanttCatalogues.MetricDefault("ChartOuterPaddingPt"), outcome.Request.ChartOuterPaddingPt);
    }

    /// <summary>
    /// The catalogue accessor itself refuses an unknown token name rather than
    /// inventing a value.
    /// </summary>
    /// <remarks>
    /// The positive test for the validator: a bad token name must produce the error
    /// path. Without it, a typo in one of the factory's token constants would resolve
    /// to nothing and the metric would silently disappear from the chart.
    /// </remarks>
    [Fact]
    public void An_unknown_metric_token_is_refused_by_the_catalogue()
    {
        Assert.Throws<ArgumentException>(() => GanttCatalogues.MetricDefault("NoSuchTokenPt"));
        Assert.False(GanttCatalogues.IsMetricToken("NoSuchTokenPt"));
        Assert.False(GanttCatalogues.IsMetricToken(null));
        Assert.True(GanttCatalogues.IsMetricToken("TitleBandHeightPt"));
    }

    /// <summary>
    /// Every setting key the factory reads is a key the configuration reader can
    /// actually return.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>This is the test that would have caught the key-name mismatches.</b> The
    /// factory read <c>"DateFormat"</c> while the catalogue key is
    /// <c>DateDisplayFormat</c>, and read <c>"DelineatorStackGapPt"</c> while the
    /// metric token is <c>StackGapPt</c>. Both lookups always missed. Their tests
    /// passed because each injected the key straight into a hand-built dictionary,
    /// so the suite proved the factory honours a key it was handed — never that the
    /// workbook can hand it one.
    /// </para>
    /// <para>
    /// Asserting against <see cref="GanttCatalogues.Settings"/> closes that gap: a key
    /// that is not in the approved set is now a failing test rather than a silently
    /// dead branch. The <c>SizePreset</c> and <c>RangePaddingDays</c> keys are
    /// currently <em>absent</em> from that set, which is a known gap recorded in
    /// STATUS.md and pending a schema bump — they are therefore not asserted here,
    /// and instead named explicitly below so the gap stays visible.
    /// </para>
    /// </remarks>
    [Fact]
    public void Every_setting_key_the_factory_reads_exists_in_the_catalogue()
    {
        HashSet<string> approved = new(
            GanttCatalogues.Settings.Select(static setting => setting.Key),
            StringComparer.Ordinal);

        // The keys the factory reads as *settings*. Each was read by literal string in
        // the pre-fix code; DateFormat is listed because it is the name the factory
        // used and the catalogue does NOT carry it, so this assertion is what turns
        // the rename into a test failure rather than a silent behaviour change.
        string[] read =
        [
            "ChartTitle",
            "AlternateBanding",
            "ShowMinorGrid",
            "ShowMajorGrid",
            "TimeScale",
            "PeriodLabelFormat",
            "DateDisplayFormat",
        ];

        foreach (string key in read)
        {
            Assert.True(
                approved.Contains(key),
                $"The factory reads setting '{key}', which the catalogue does not define.");
        }

        // The retired spelling must be gone, or the guard above would still pass on a
        // rename that left the old name in place as a second, dead read.
        Assert.DoesNotContain("DateFormat", approved);
    }

    /// <summary>
    /// The two settings R4.7H and the plot-range padding need are now real catalogue
    /// keys, so the factory's reads of them can actually be satisfied.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This test previously asserted the <em>opposite</em>: that <c>SizePreset</c> and
    /// <c>RangePaddingDays</c> were absent from the catalogue, which recorded a known
    /// gap so it could not be closed silently. The owner's ruling permitted the schema
    /// bump that adding a key requires, so the gap is closed and the assertion is
    /// inverted — the keys must now be present, or the factory reads them from a map
    /// that can never supply them and the presets fall back silently again.
    /// </para>
    /// <para>
    /// Both defaults are asserted, not just membership: a present key carrying a value
    /// the factory cannot parse would pass a membership check and still be dead.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_size_preset_keys_are_now_catalogue_settings()
    {
        HashSet<string> approved = new(
            GanttCatalogues.Settings.Select(static setting => setting.Key),
            StringComparer.Ordinal);

        Assert.Contains("SizePreset", approved);
        Assert.Contains("RangePaddingDays", approved);

        // The stored defaults must be values the factory accepts: a preset key that
        // `SizePresets.ByKey` rejects would be rejected at render time on a freshly
        // initialised workbook.
        GanttSettingDefinition preset = GanttCatalogues.Settings.First(s => s.Key == "SizePreset");
        Assert.NotNull(SizePresets.ByKey(preset.DefaultValue));

        GanttSettingDefinition padding = GanttCatalogues.Settings.First(s => s.Key == "RangePaddingDays");
        Assert.True(
            int.TryParse(padding.DefaultValue, CultureInfo.InvariantCulture, out int days) && days >= 0,
            $"RangePaddingDays default '{padding.DefaultValue}' is not a non-negative integer.");
    }

    /// <summary>A null collaborator is refused rather than dereferenced.</summary>
    [Fact]
    public void Null_inputs_are_refused()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        Assert.Throws<ArgumentNullException>(() => factory.Create(
            null!,
            new Dictionary<string, string>(),
            StyleRegistry(),
            Grid()));
    }

    /// <summary>A registry with the one style the fixture events reference.</summary>
    private static GanttStyleRegistry StyleRegistry() =>
        new(
            [
                new GanttStyleDefinition(
                    "AsPlannedActivity",
                    new HashSet<GanttLabelPosition> { GanttLabelPosition.Inside },
                    EntityColourCapability.Fill | EntityColourCapability.Stroke,
                    DefaultLabelPosition: GanttLabelPosition.Inside,
                    FillColour: "#DDEBF7",
                    StrokeColour: "#2E75B6",
                    TextColour: "#000000",
                    HatchPattern: GanttHatchPattern.None,
                    StandardOutlinePt: 1,
                    ActivityHeightPt: 12),
            ]);
}
