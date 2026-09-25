using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

public sealed class FrameBandsBuilderTests
{
    private static readonly FrameBandsTheme _theme = new(
        new SceneStyle("Background"),
        new SceneStyle("AlternateBand"),
        new SceneStyle("MinorGrid", strokeColour: ColourHex.Parse("#D9D9D9"), outlineWidthPt: 0.5),
        new SceneStyle("MajorGrid", strokeColour: ColourHex.Parse("#000000"), outlineWidthPt: 1),
        new SceneStyle("YearHeader"),
        new SceneStyle("PeriodHeader"),
        new SceneStyle("Title")
    );

    [Fact]
    public void Builds_exact_bounds_and_chart_owned_primitive_layers()
    {
        FrameBandsCreationOutcome outcome = Build(CreateRequest());

        Assert.True(outcome.Succeeded);
        ChartFrameGeometry geometry = outcome.Result!.Geometry;
        Assert.Equal(new RectD(-6, 36, 412, 170), geometry.ChartBounds);
        Assert.Equal(new RectD(0, 66, 400, 134), geometry.ContentBounds);
        Assert.Equal(new RectD(0, 42, 400, 24), geometry.TitleBounds);
        Assert.All(outcome.Result.Primitives, primitive => Assert.Equal(SceneOwnerId.Chart, primitive.OwnerId));
        Assert.Contains(outcome.Result.Primitives, primitive => primitive.ZLayer == ZLayer.Background);
        Assert.Contains(outcome.Result.Primitives, primitive => primitive.ZLayer == ZLayer.AlternateBand);
        Assert.Contains(outcome.Result.Primitives, primitive => primitive.ZLayer == ZLayer.Grid);
        Assert.Contains(outcome.Result.Primitives, primitive => primitive.ZLayer == ZLayer.Frame);
        Assert.Contains(outcome.Result.Primitives, primitive => primitive.ZLayer == ZLayer.Title);
    }

    [Fact]
    public void Emits_clipped_year_and_period_headers_with_alternating_bands()
    {
        FrameBandsResult result = Build(CreateRequest()).Result!;

        Assert.Equal(
            1,
            result.Primitives.OfType<SceneRect>().Count(rect => rect.PrimitiveId.StartsWith("chart:year:", StringComparison.Ordinal))
        );
        Assert.Equal(
            3,
            result.Primitives.OfType<SceneRect>().Count(rect => rect.PrimitiveId.StartsWith("chart:period:", StringComparison.Ordinal))
        );
        Assert.Single(result.Primitives.OfType<SceneRect>(), rect => rect.PrimitiveId.StartsWith("chart:band:", StringComparison.Ordinal));
        Assert.Equal(3, result.Primitives.OfType<SceneText>().Count(text => text.PrimitiveId.Contains(":label", StringComparison.Ordinal)));
        Assert.Equal(
            "Jan",
            result.Primitives.OfType<SceneText>().Single(text => text.PrimitiveId.Contains("2024-01-01", StringComparison.Ordinal)).Text
        );
    }

    [Fact]
    public void ShowTitle_false_omits_title_band_and_text()
    {
        FrameBandsResult result = Build(CreateRequest() with { ShowTitle = false }).Result!;

        Assert.DoesNotContain(result.Primitives, primitive => primitive.PrimitiveId.StartsWith("chart:title", StringComparison.Ordinal));
        Assert.Null(result.Geometry.TitleBounds);
    }

    [Fact]
    public void Title_overflow_uses_ellipsis_and_warning_without_changing_font()
    {
        FixedMeasurer measurer = new(_ => 10_000);
        FrameBandsCreationOutcome outcome = FrameBandsBuilder.TryBuild(CreateRequest(), measurer);

        Assert.True(outcome.Succeeded);
        SceneText title = Assert.Single(outcome.Result!.Primitives.OfType<SceneText>(), text => text.PrimitiveId == "chart:title-text");
        Assert.EndsWith("…", title.Text, StringComparison.Ordinal);
        Assert.Contains(outcome.Result.Warnings, warning => warning.Code == "TitleOverflow");
        Assert.Equal(_theme.Title, title.Style);
    }

    [Fact]
    public void Narrow_periods_suppress_labels_and_warn()
    {
        FrameBandsResult result = Build(CreateRequest() with { MinimumHeaderLabelWidthPt = 1_000 }).Result!;

        Assert.All(
            result.Primitives.OfType<SceneText>().Where(text => text.PrimitiveId.Contains(":label", StringComparison.Ordinal)),
            text => Assert.False(text.PrimitiveId.Contains("period", StringComparison.Ordinal))
        );
        Assert.Contains(result.Warnings, warning => warning.Code == "AllPeriodLabelsSuppressed");
    }

    [Fact]
    public void Invalid_requests_return_typed_refusals()
    {
        Assert.Equal(FrameBandsRefusal.NullRequest, FrameBandsBuilder.TryBuild(null, new FixedMeasurer(_ => 1)).Refusal);
        Assert.Equal(
            FrameBandsRefusal.InvalidTimeScale,
            FrameBandsBuilder.TryBuild(CreateRequest() with { TimeScale = null! }, new FixedMeasurer(_ => 1)).Refusal
        );
        Assert.Equal(
            FrameBandsRefusal.InvalidTheme,
            FrameBandsBuilder.TryBuild(CreateRequest() with { Theme = null! }, new FixedMeasurer(_ => 1)).Refusal
        );
        Assert.Equal(
            FrameBandsRefusal.IncompatibleSettings,
            FrameBandsBuilder
                .TryBuild(
                    CreateRequest() with
                    {
                        Scale = GanttTimeScale.Quarter,
                        PeriodLabelFormat = GanttPeriodLabelFormat.MMM,
                    },
                    new FixedMeasurer(_ => 1)
                )
                .Refusal
        );
        Assert.Equal(
            FrameBandsRefusal.InvalidGeometry,
            FrameBandsBuilder.TryBuild(CreateRequest() with { PlotBounds = default }, new FixedMeasurer(_ => 1)).Refusal
        );
        Assert.Equal(
            FrameBandsRefusal.BlankVisibleTitle,
            FrameBandsBuilder.TryBuild(CreateRequest() with { ChartTitle = " " }, new FixedMeasurer(_ => 1)).Refusal
        );
        Assert.Equal(FrameBandsRefusal.TextMeasurementUnavailable, FrameBandsBuilder.TryBuild(CreateRequest(), null).Refusal);
        Assert.Equal(
            FrameBandsRefusal.TextMeasurementUnavailable,
            FrameBandsBuilder.TryBuild(CreateRequest(), new FixedMeasurer(_ => double.NaN)).Refusal
        );
    }

    [Fact]
    public void Clips_first_and_last_periods_across_a_year_boundary()
    {
        TimeScale scale = TimeScale.TryCreate(new DateOnly(2023, 12, 15), new DateOnly(2024, 2, 10), 100, 400).Scale!;
        FrameBandsResult result = Build(
            CreateRequest() with
            {
                TimeScale = scale,
                Scale = GanttTimeScale.Month,
                PeriodLabelFormat = GanttPeriodLabelFormat.MMM,
            }
        ).Result!;

        SceneText[] periodLabels = result
            .Primitives.OfType<SceneText>()
            .Where(text => text.PrimitiveId.StartsWith("chart:period:", StringComparison.Ordinal))
            .ToArray();
        Assert.Equal(["Dec", "Jan", "Feb"], periodLabels.Select(text => text.Text));
        Assert.Equal(
            2,
            result.Primitives.OfType<SceneRect>().Count(rect => rect.PrimitiveId.StartsWith("chart:year:", StringComparison.Ordinal))
        );
        double yearBoundaryX = scale.DateToX(new DateOnly(2024, 1, 1));
        Assert.Contains(result.Primitives.OfType<SceneLine>(), line => line.From.X == yearBoundaryX && line.Style.StyleKey == "MajorGrid");
    }

    [Theory]
    [InlineData(GanttTimeScale.Month, GanttPeriodLabelFormat.MM, "01")]
    [InlineData(GanttTimeScale.Month, GanttPeriodLabelFormat.MMM, "Jan")]
    [InlineData(GanttTimeScale.Quarter, GanttPeriodLabelFormat.Quarter, "Q1")]
    [InlineData(GanttTimeScale.Year, GanttPeriodLabelFormat.Year, "2024")]
    public void Supports_each_closed_scale_and_format_pair(GanttTimeScale scale, GanttPeriodLabelFormat format, string expectedLabel)
    {
        FrameBandsResult result = Build(CreateRequest() with { Scale = scale, PeriodLabelFormat = format }).Result!;

        Assert.Contains(result.Primitives.OfType<SceneText>(), text => text.Text == expectedLabel);
    }

    [Fact]
    public void Chart_owner_snapshot_round_trips_and_deduplicates_grid_boundaries()
    {
        FrameBandsResult result = Build(CreateRequest()).Result!;
        GanttScene scene = GanttScene
            .TryCreate(result.Geometry.ChartBounds, result.Geometry.ChartBounds, result.Primitives, result.Warnings)
            .Scene!;
        string json = SceneSnapshot.Serialize(scene);
        GanttScene roundTripped = SceneSnapshot.Deserialize(json);

        Assert.Equal(json, SceneSnapshot.Serialize(roundTripped));
        Assert.All(roundTripped.Primitives, primitive => Assert.Equal(SceneOwnerId.Chart, primitive.OwnerId));
        Assert.Equal(
            4,
            roundTripped.Primitives.OfType<SceneLine>().Count(line => line.PrimitiveId.StartsWith("chart:grid:", StringComparison.Ordinal))
        );
    }

    private static FrameBandsCreationOutcome Build(FrameBandsRequest request) =>
        FrameBandsBuilder.TryBuild(request, new FixedMeasurer(text => text.Length * 6.0));

    private static FrameBandsRequest CreateRequest()
    {
        TimeScale scale = TimeScale.TryCreate(new DateOnly(2024, 1, 1), new DateOnly(2024, 3, 31), 100, 400).Scale!;
        return new FrameBandsRequest(
            scale,
            GanttTimeScale.Month,
            GanttPeriodLabelFormat.MMM,
            new RectD(0, 100, 100, 100),
            new RectD(100, 100, 300, 100),
            6,
            24,
            18,
            16,
            18,
            0.5,
            1,
            true,
            "Gantt Chart",
            true,
            true,
            true,
            _theme
        );
    }

    private sealed class FixedMeasurer(Func<string, double> measure) : ITextWidthMeasurer
    {
        public bool TryMeasure(string text, out double widthPt)
        {
            widthPt = measure(text);
            return double.IsFinite(widthPt) && widthPt >= 0;
        }
    }
}
