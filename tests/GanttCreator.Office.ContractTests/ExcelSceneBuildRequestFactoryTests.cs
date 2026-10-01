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

    /// <summary>A known preset key is honoured.</summary>
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
    [Fact]
    public void A_malformed_numeric_setting_falls_back_to_its_default()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string> { ["GridLinePt"] = "not-a-number" },
            StyleRegistry(),
            Grid());

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        Assert.Equal(0.5, outcome.Request!.GridLinePt);
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
