using System.Globalization;
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

    /// <summary>
    /// Each side of the frame margin is applied independently (ADR-0031 D1).
    /// </summary>
    /// <remarks>
    /// <b>This is the test the single scalar could not have.</b> The old request
    /// carried one <c>ChartOuterPaddingPt</c> for all four sides, so there was no way
    /// to express — let alone assert — a live chart with no left margin and row-sized
    /// top and bottom margins. Checking only that "the bounds are the content plus
    /// 6pt all round" would pass against a builder that ignored the per-side values
    /// entirely and applied a constant.
    /// </remarks>
    [Fact]
    public void Each_side_of_the_padding_is_applied_to_its_own_edge()
    {
        FrameBandsCreationOutcome outcome = Build(
            CreateRequest() with
            {
                // The live arrangement: no left margin, tall top and bottom rows, and
                // the chrome token on the right.
                Padding = new ChartPaddingPt(LeftPt: 0, TopPt: 30, RightPt: 6, BottomPt: 40),
            });

        Assert.True(outcome.Succeeded);
        ChartFrameGeometry geometry = outcome.Result!.Geometry;
        RectD chart = geometry.ChartBounds;
        RectD union = UnionOf(geometry);

        Assert.Equal(union.Left, chart.Left, precision: 6);
        Assert.Equal(union.Top - 30, chart.Top, precision: 6);
        Assert.Equal(union.Right + 6, chart.Right, precision: 6);
        Assert.Equal(union.Bottom + 40, chart.Bottom, precision: 6);
    }

    /// <summary>
    /// A uniform padding reproduces the old single-scalar behaviour exactly, which is
    /// what every export profile passes.
    /// </summary>
    /// <remarks>
    /// The counterweight to the per-side test: without it, a "simplification" back to
    /// one scalar would still pass the live assertion and silently reintroduce the
    /// left margin on every exported chart.
    /// </remarks>
    [Fact]
    public void A_uniform_padding_matches_the_content_plus_that_margin_on_all_sides()
    {
        const double padding = 6;
        FrameBandsCreationOutcome outcome = Build(
            CreateRequest() with { Padding = ChartPaddingPt.Uniform(padding) });

        Assert.True(outcome.Succeeded);
        ChartFrameGeometry geometry = outcome.Result!.Geometry;
        RectD chart = geometry.ChartBounds;
        RectD union = UnionOf(geometry);

        Assert.Equal(union.Left - padding, chart.Left, precision: 6);
        Assert.Equal(union.Top - padding, chart.Top, precision: 6);
        Assert.Equal(union.Right + padding, chart.Right, precision: 6);
        Assert.Equal(union.Bottom + padding, chart.Bottom, precision: 6);
    }

    /// <summary>
    /// A negative or non-finite margin on ANY side is refused, not just the first.
    /// </summary>
    /// <remarks>
    /// <b>The positive test for the per-side validator.</b> A check that only
    /// inspected one member would accept a request whose other three sides were
    /// sound, producing chart bounds that extended past their own content on the one
    /// unchecked side — a rectangle no caller asked for and nothing downstream would
    /// report. Every side is a separate case here for that reason.
    /// </remarks>
    [Theory]
    [InlineData(-1d, 6d, 6d, 6d)]
    [InlineData(6d, -1d, 6d, 6d)]
    [InlineData(6d, 6d, -1d, 6d)]
    [InlineData(6d, 6d, 6d, -1d)]
    [InlineData(double.NaN, 6d, 6d, 6d)]
    [InlineData(6d, 6d, double.PositiveInfinity, 6d)]
    public void A_bad_margin_on_any_side_is_refused(
        double left,
        double top,
        double right,
        double bottom)
    {
        FrameBandsCreationOutcome outcome = Build(
            CreateRequest() with { Padding = new ChartPaddingPt(left, top, right, bottom) });

        Assert.False(outcome.Succeeded);
        Assert.Equal(FrameBandsRefusal.InvalidGeometry, outcome.Refusal);
    }

    /// <summary>
    /// Zero on every side is ACCEPTED: a chart that hugs its content is a legitimate
    /// request, not a degenerate one.
    /// </summary>
    /// <remarks>
    /// The counterweight to the refusal theory above. Validating "non-negative" with
    /// a strictly-positive test would refuse the live profile's own left margin, which
    /// is exactly the value ADR-0031 D2 sets to zero.
    /// </remarks>
    [Fact]
    public void A_zero_margin_on_every_side_is_accepted()
    {
        FrameBandsCreationOutcome outcome = Build(
            CreateRequest() with { Padding = ChartPaddingPt.None });

        Assert.True(outcome.Succeeded);
        Assert.Equal(UnionOf(outcome.Result!.Geometry), outcome.Result.Geometry.ChartBounds);
    }

    /// <summary>
    /// The rectangle the frame is derived from: the content bounds widened to include
    /// the title band, when one is present.
    /// </summary>
    /// <param name="geometry">The resolved frame geometry.</param>
    /// <returns>The union the chart bounds are padded from.</returns>
    /// <remarks>
    /// The frame wraps the union, not <c>ContentBounds</c> alone: entity guide §1
    /// defines the chart as the title, panel, headers and plot plus the margin, so a
    /// test that compared against the content rectangle would be asserting the wrong
    /// edge and would fail for a correct builder.
    /// </remarks>
    private static RectD UnionOf(ChartFrameGeometry geometry)
    {
        RectD content = Assert.IsType<RectD>(geometry.ContentBounds);
        if (geometry.TitleBounds is not { } title)
        {
            return content;
        }

        double left = Math.Min(content.Left, title.Left);
        double top = Math.Min(content.Top, title.Top);
        double right = Math.Max(content.Right, title.Right);
        double bottom = Math.Max(content.Bottom, title.Bottom);
        return new RectD(left, top, right - left, bottom - top);
    }

    /// <summary>
    /// The period band sits BELOW the year band, directly above the plot
    /// (entity guide §5/§6).
    /// </summary>
    /// <remarks>
    /// <b>This is the test the inversion was missing.</b> The bands were built
    /// year-then-period upward from the plot's top, which put the year band directly
    /// above the plot and pushed the period band above it — the reverse of guide §6
    /// ("one clipped cell per period below the year band"). On a live sheet that drew
    /// the year band over the header row. No test asserted the RELATIONSHIP: each
    /// band's own bounds were checked, never which one sat above the other, so the
    /// suites stayed green through a live chart showing the mistake.
    /// </remarks>
    [Fact]
    public void The_period_band_is_below_the_year_band_and_above_the_plot()
    {
        FrameBandsCreationOutcome outcome = Build(CreateRequest());

        Assert.True(outcome.Succeeded);
        ChartFrameGeometry geometry = outcome.Result!.Geometry;

        // Stacked with no gap and no overlap: the period band's top is the year's
        // bottom, and its bottom is the plot's top.
        Assert.Equal(geometry.YearBounds.Bottom, geometry.PeriodBounds.Y, precision: 6);
        Assert.True(
            geometry.YearBounds.Y < geometry.PeriodBounds.Y,
            $"Year band ({geometry.YearBounds.Y}) must sit above the period band ({geometry.PeriodBounds.Y}).");
    }

    /// <summary>
    /// The emitted year and period header PRIMITIVES sit where the returned geometry
    /// says they do.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This is the assertion the inversion was missing. The sibling test above checks
    /// that <c>geometry.YearBounds</c> is above <c>geometry.PeriodBounds</c>, and it
    /// passed throughout - because <c>TryBuild</c> built that geometry correctly. But
    /// <c>AddHeaders</c> recomputed both bands from the plot's top edge and the two
    /// band heights, in the opposite order, so the primitives that were actually drawn
    /// were mirrored relative to the geometry the same method returned. Checking the
    /// geometry alone therefore could never see the defect.
    /// </para>
    /// <para>
    /// The live symptom was month labels in the upper row and the year in the lower
    /// one, which is entity guide §5/§6 inverted on screen.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_emitted_header_primitives_agree_with_the_returned_geometry()
    {
        FrameBandsRequest request = CreateRequest();
        FrameBandsCreationOutcome outcome = Build(request);

        Assert.True(outcome.Succeeded);
        ChartFrameGeometry geometry = outcome.Result!.Geometry;

        SceneRect yearBand = Assert.Single(
            outcome.Result.Primitives.OfType<SceneRect>(),
            rect => rect.PrimitiveId.StartsWith("chart:year:", StringComparison.Ordinal));

        // One rectangle per period, so there is more than one; every one of them must
        // share the period band's vertical placement.
        SceneRect[] periodBands =
        [
            .. outcome.Result.Primitives
                .OfType<SceneRect>()
                .Where(rect => rect.PrimitiveId.StartsWith("chart:period:", StringComparison.Ordinal)),
        ];
        Assert.NotEmpty(periodBands);

        // Vertical placement and height come from the geometry; only the horizontal
        // extent differs per band, so X and Width are deliberately not compared.
        Assert.Equal(geometry.YearBounds.Y, yearBand.Bounds.Y, precision: 6);
        Assert.Equal(geometry.YearBounds.Height, yearBand.Bounds.Height, precision: 6);
        Assert.All(periodBands, band => Assert.Equal(geometry.PeriodBounds.Y, band.Bounds.Y, precision: 6));
        Assert.All(periodBands, band => Assert.Equal(geometry.PeriodBounds.Height, band.Bounds.Height, precision: 6));

        // The relationship itself, on the drawn primitives rather than the geometry.
        Assert.True(
            yearBand.Bounds.Y < periodBands[0].Bounds.Y,
            $"The drawn year band ({yearBand.Bounds.Y}) must sit above the drawn period band ({periodBands[0].Bounds.Y}).");
        Assert.All(periodBands, band => Assert.Equal(request.PlotBounds.Top, band.Bounds.Bottom, precision: 6));
    }

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
    public void The_outer_padding_is_inside_the_chart_bounds_on_all_four_sides()
    {
        // Entity guide §1: ChartBounds is the union of the title, data panel, time
        // headers, and plot "plus ChartOuterPaddingPt". The product decision recorded
        // there is that the padding is INSIDE the bounds, not a renderer-side margin -
        // which matters because §1 also says empty whitespace outside the bounds is
        // not exported, so a renderer that treated the padding as outside would export
        // a chart missing exactly ChartOuterPaddingPt of margin.
        //
        // Checked against the content and title rectangles the builder reports, not
        // against a restatement of its own arithmetic, so this is a real cross-check
        // of the derived bounds rather than a tautology.
        const double padding = 6;
        ChartFrameGeometry geometry = Build(CreateRequest() with { Padding = ChartPaddingPt.Uniform(padding) }).Result!.Geometry;

        RectD content = Assert.IsType<RectD>(geometry.ContentBounds);
        RectD title = Assert.IsType<RectD>(geometry.TitleBounds);

        // The union is content plus the title, because the title band sits above the
        // content and is part of the chart.
        double unionLeft = Math.Min(content.Left, title.Left);
        double unionTop = Math.Min(content.Top, title.Top);
        double unionRight = Math.Max(content.Right, title.Right);
        double unionBottom = Math.Max(content.Bottom, title.Bottom);

        Assert.Equal(unionLeft - padding, geometry.ChartBounds.Left, precision: 9);
        Assert.Equal(unionTop - padding, geometry.ChartBounds.Top, precision: 9);
        Assert.Equal(unionRight + padding, geometry.ChartBounds.Right, precision: 9);
        Assert.Equal(unionBottom + padding, geometry.ChartBounds.Bottom, precision: 9);

        // The negative origin is a real, reachable outcome of that rule, not a
        // hypothetical: Excel cannot express a negative shape offset, which is why the
        // renderer translation rule exists.
        Assert.True(
            geometry.ChartBounds.Left < 0,
            $"Expected the derived origin to be negative with content at the origin, got {geometry.ChartBounds}.");
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

        // Four header labels: one year plus three periods. Before R3.17 the year
        // label used a "chart:year-label:" prefix, so this count silently covered
        // only the periods and the year label was unasserted; the convention fix
        // made it visible.
        Assert.Equal(4, result.Primitives.OfType<SceneText>().Count(text => text.PrimitiveId.EndsWith(":label", StringComparison.Ordinal)));
        Assert.Equal(
            "Jan",
            result.Primitives.OfType<SceneText>().Single(text => text.PrimitiveId.Contains("2024-01-01", StringComparison.Ordinal)).Text
        );
    }

    [Fact]
    public void Band_and_title_text_is_centred_in_its_resolved_bounds()
    {
        // §2 fixes the title as "horizontally centred and vertically
        // middle-aligned", and the year/period band labels are centred in their
        // band. Before ADR-0018 these carried GanttLabelPosition.Auto, which is
        // a label *position* and meaningless as a text alignment, so nothing
        // pinned the centring. This is the positive pin for that behaviour.
        FrameBandsResult result = Build(CreateRequest()).Result!;

        SceneText[] texts = [.. result.Primitives.OfType<SceneText>()];
        Assert.NotEmpty(texts);
        Assert.All(texts, text => Assert.Equal(GanttTextAlignment.Centre, text.Alignment));
        Assert.Equal(GanttTextAlignment.Centre, result.Primitives.OfType<SceneText>().Single(text => text.PrimitiveId == "chart:title-text").Alignment);
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
    public void Mismatched_time_scale_plot_edges_are_refused()
    {
        var leftMismatch = CreateRequest() with
        {
            TimeScale = TimeScale
                .TryCreate(new DateOnly(2024, 1, 1), new DateOnly(2024, 3, 31), 100 + (GeometryMath.Epsilon * 2), 400)
                .Scale!,
        };
        var rightMismatch = CreateRequest() with
        {
            TimeScale = TimeScale
                .TryCreate(new DateOnly(2024, 1, 1), new DateOnly(2024, 3, 31), 100, 400 - (GeometryMath.Epsilon * 2))
                .Scale!,
        };

        Assert.Equal(FrameBandsRefusal.InvalidGeometry, Build(leftMismatch).Refusal);
        Assert.Equal(FrameBandsRefusal.InvalidGeometry, Build(rightMismatch).Refusal);
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

        Assert.Contains(
            result.Primitives.OfType<SceneText>(),
            text => text.PrimitiveId.StartsWith("chart:period:", StringComparison.Ordinal) && text.Text == expectedLabel
        );
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

    [Fact]
    public void Snapshot_is_identical_under_de_de_and_th_th()
    {
        string germanSnapshot = BuildSnapshot("de-DE");
        string thaiSnapshot = BuildSnapshot("th-TH");

        Assert.Equal(germanSnapshot, thaiSnapshot);
    }

    private static string BuildSnapshot(string cultureName)
    {
        CultureInfo originalCulture = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo(cultureName);
            FrameBandsResult result = Build(CreateRequest()).Result!;
            GanttScene scene = GanttScene
                .TryCreate(result.Geometry.ChartBounds, result.Geometry.ChartBounds, result.Primitives, result.Warnings)
                .Scene!;
            return SceneSnapshot.Serialize(scene);
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
        }
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
            ChartPaddingPt.Uniform(6),
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

    [Fact]
    public void Every_header_label_is_named_after_its_own_band_primitive()
    {
        // A header label names its parent by appending ":label" to the parent's
        // identifier. The year header used to use a "chart:year-label:" prefix
        // while the period header appended ":label", so the two header kinds
        // disagreed about how a child names its parent; a renderer reconciles on
        // this exact text, so the disagreement was a real reconciliation hazard
        // (R3.17). This is convention-agnostic: it pins the relationship, not a
        // literal spelling, so a future rename cannot silently break it.
        FrameBandsResult result = Build(CreateRequest()).Result!;

        HashSet<string> rectIds =
        [
            .. result.Primitives.OfType<SceneRect>().Select(rect => rect.PrimitiveId),
        ];

        SceneText[] labels = [.. result.Primitives.OfType<SceneText>().Where(text => text.PrimitiveId.EndsWith(":label", StringComparison.Ordinal))];

        // The period header emits three labels and the year header one, so a
        // non-empty set is required: an empty one would make All vacuous.
        Assert.NotEmpty(labels);
        Assert.All(labels, label =>
        {
            string parentId = label.PrimitiveId[..^":label".Length];
            Assert.Contains(parentId, rectIds);
        });
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
