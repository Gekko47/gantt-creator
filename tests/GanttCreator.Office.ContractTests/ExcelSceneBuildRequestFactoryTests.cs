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

    /// <summary>A measured grid narrow enough to leave a readable A4-landscape plot.</summary>
    private static PanelCellGrid Grid(double widthPt = 120) =>
        PanelCellGrid.TryCreate([new PanelColumn("Id", widthPt)], [15.0], 15, ["Id"]).Grid!;

    private static FakeTextMetrics Metrics() => new();

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
    /// Records the two keys R4.7H introduced that the settings catalogue does not yet
    /// carry, so the gap is asserted rather than remembered.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <c>SizePreset</c> and <c>RangePaddingDays</c> are read by the factory but are
    /// not in <c>GanttCatalogues.Settings</c>, so <c>ValidateSettings</c> cannot return
    /// them and both always fall back. R4.7H's size presets are therefore unreachable
    /// in production while <c>A_known_preset_key_is_honoured</c> passes on an injected
    /// dictionary.
    /// </para>
    /// <para>
    /// Adding either key changes the <c>tblGanttSettings</c> key/value contract, which
    /// <see cref="GanttSchemaVersion"/> requires a version bump for — so this is
    /// deliberately <b>not</b> fixed in the same change as the renames. The test
    /// asserts the gap still exists, and its failure message names the consequence,
    /// so closing it cannot be forgotten and cannot happen silently.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_size_preset_keys_are_still_absent_from_the_catalogue()
    {
        HashSet<string> approved = new(
            GanttCatalogues.Settings.Select(static setting => setting.Key),
            StringComparer.Ordinal);

        Assert.DoesNotContain("SizePreset", approved);
        Assert.DoesNotContain("RangePaddingDays", approved);
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
