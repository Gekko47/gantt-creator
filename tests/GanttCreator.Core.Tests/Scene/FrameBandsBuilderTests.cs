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
    /// The plot-spanning shapes are lifted into the header row by the configured
    /// overlap, and the TOP is the only edge that moves up (ADR-0037 D1).
    /// </summary>
    /// <remarks>
    /// <para>
    /// The mechanism exists because Excel resizes a shape on a row insert only when
    /// the insert is strictly below the shape's cell anchor. With the plot top exactly
    /// on the header/body boundary, a row added at the top slid the shape down instead
    /// of stretching it and the plot stayed unpainted there (measured 2026-10-03,
    /// <c>scripts/probe-frame-stretch.ps1</c>). Lifting the top resolves the anchor to
    /// the header row.
    /// </para>
    /// <para>
    /// <b>The height is now the sum of BOTH overlaps</b>, because ADR-0038 extends the
    /// bottom through the anchor row as well. This test isolates the TOP by passing a
    /// zero anchor row, and
    /// <c>The_plot_spanning_shapes_extend_down_through_the_anchor_row_and_keep_their_top</c>
    /// isolates the bottom. Neither alone would catch a builder that applied one edge's
    /// adjustment to both.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_plot_spanning_shapes_lift_into_the_header_row_and_keep_their_bottom()
    {
        // Zero anchor row: this test is about the TOP, and ADR-0038's bottom extension
        // is asserted separately.
        FrameBandsCreationOutcome outcome = Build(CreateRequest(0.5, 0));
        Assert.True(outcome.Succeeded, $"Refused: {outcome.Refusal}");

        FrameBandsRequest request = CreateRequest(0.5, 0);
        RectD plot = request.PlotBounds;

        var bands = outcome.Result!.Primitives.OfType<SceneRect>()
            .Where(rect => rect.PrimitiveId.StartsWith("chart:band:", StringComparison.Ordinal))
            .ToList();
        Assert.NotEmpty(bands);

        foreach (SceneRect band in bands)
        {
            Assert.Equal(plot.Y - 0.5, band.Bounds.Y, 6);
            Assert.Equal(plot.Height + 0.5, band.Bounds.Height, 6);
            Assert.Equal(plot.Bottom, band.Bounds.Bottom, 6);
        }

        // The vertical grid lines share the same span, so they stretch with the bands.
        var gridLines = outcome.Result!.Primitives.OfType<SceneLine>()
            .Where(line => line.PrimitiveId.StartsWith("chart:grid", StringComparison.Ordinal))
            .ToList();
        Assert.NotEmpty(gridLines);

        foreach (SceneLine line in gridLines)
        {
            Assert.Equal(plot.Y - 0.5, line.From!.Y, 6);
            Assert.Equal(plot.Bottom, line.To!.Y, 6);
        }
    }

    /// <summary>
    /// Zero anchor row reproduces the previous geometry EXACTLY, so the mechanism is
    /// retunable rather than baked in (ADR-0038 D1).
    /// </summary>
    /// <remarks>
    /// <para>
    /// The positive test for the retunable claim, and now the COUNTERWEIGHT to the
    /// bottom-extension test: that test asserts the bands end <em>past</em> the plot's
    /// bottom, and this one asserts they end <em>on</em> it when the anchor row is
    /// zero. Without the pair, a builder that ignored the request and always applied
    /// the catalogue default would pass the first while making the token a lie.
    /// </para>
    /// <para>
    /// Both edges are zeroed together so the whole plot span is the plot rectangle:
    /// asserting them independently would let a builder that honoured the top and
    /// ignored the bottom still pass.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_zero_anchor_row_reproduces_the_plot_bounds_exactly()
    {
        FrameBandsCreationOutcome outcome = Build(CreateRequest(0, 0));
        Assert.True(outcome.Succeeded, $"Refused: {outcome.Refusal}");

        RectD plot = CreateRequest(0, 0).PlotBounds;

        foreach (SceneRect band in outcome.Result!.Primitives.OfType<SceneRect>()
            .Where(rect => rect.PrimitiveId.StartsWith("chart:band:", StringComparison.Ordinal)))
        {
            Assert.Equal(plot.Y, band.Bounds.Y, 9);
            Assert.Equal(plot.Height, band.Bounds.Height, 9);
        }

        foreach (SceneLine line in outcome.Result!.Primitives.OfType<SceneLine>()
            .Where(line => line.PrimitiveId.StartsWith("chart:grid", StringComparison.Ordinal)))
        {
            Assert.Equal(plot.Y, line.From!.Y, 9);
            Assert.Equal(plot.Bottom, line.To!.Y, 9);
        }
    }

    /// <summary>
    /// The plot-spanning shapes extend DOWN through the anchor row, and the closing
    /// line terminates them there (ADR-0038 D2/D3).
    /// </summary>
    /// <remarks>
    /// <para>
    /// The extension is <c>anchorHeightPt + MajorBoundaryPt / 2</c>: at the fixture's
    /// 0.25pt anchor row and 1pt closing line that is <b>0.75pt</b>. Asserted as that
    /// computed figure rather than a literal, so the arithmetic is pinned rather than
    /// the constant that happens to satisfy it today.
    /// </para>
    /// <para>
    /// The top is asserted unchanged in the same test. The two edges are the asymmetry
    /// this change encodes -- top is a sub-row overlap into the header, bottom is a
    /// reserved row -- and a test that only checked the bottom would pass a builder that
    /// had accidentally started lifting the top edge too.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_plot_spanning_shapes_extend_down_through_the_anchor_row_and_keep_their_top()
    {
        const double anchorPt = 0.25;
        FrameBandsCreationOutcome outcome = Build(CreateRequest(0.5, anchorPt));
        Assert.True(outcome.Succeeded, $"Refused: {outcome.Refusal}");

        FrameBandsRequest request = CreateRequest(0.5, anchorPt);
        RectD plot = request.PlotBounds;
        double extensionPt = anchorPt + (request.MajorBoundaryPt / 2);

        var bands = outcome.Result!.Primitives.OfType<SceneRect>()
            .Where(rect => rect.PrimitiveId.StartsWith("chart:band:", StringComparison.Ordinal))
            .ToList();
        Assert.NotEmpty(bands);

        foreach (SceneRect band in bands)
        {
            Assert.Equal(plot.Y - 0.5, band.Bounds.Y, 6);
            Assert.Equal(plot.Bottom + extensionPt, band.Bounds.Bottom, 6);
        }

        var gridLines = outcome.Result!.Primitives.OfType<SceneLine>()
            .Where(line => line.PrimitiveId.StartsWith("chart:grid", StringComparison.Ordinal))
            .ToList();
        Assert.NotEmpty(gridLines);

        foreach (SceneLine line in gridLines)
        {
            Assert.Equal(plot.Y - 0.5, line.From!.Y, 6);
            Assert.Equal(plot.Bottom + extensionPt, line.To!.Y, 6);
        }
    }

    /// <summary>
    /// The chart BACKGROUND keeps covering the whole chart, headers included.
    /// </summary>
    /// <remarks>
    /// The counterweight to the overlap test. An earlier draft applied the plot span
    /// to the background too, which shrank it to the plot and would have left the
    /// title and header rows unpainted. The background's top already sits above the
    /// header row, so it needs no overlap at all.
    /// </remarks>
    [Fact]
    public void The_background_still_spans_the_whole_chart_and_is_not_lifted()
    {
        FrameBandsCreationOutcome outcome = Build(CreateRequest(0.5));
        Assert.True(outcome.Succeeded, $"Refused: {outcome.Refusal}");

        SceneRect background = outcome.Result!.Primitives.OfType<SceneRect>()
            .Single(rect => rect.PrimitiveId == "chart:background");

        Assert.Equal(outcome.Result.Geometry.ChartBounds.Top, background.Bounds.Top, 6);
        Assert.Equal(outcome.Result.Geometry.ChartBounds.Height, background.Bounds.Height, 6);
    }

    /// <summary>
    /// The closing line sits exactly on the bands' bottom edge and covers the overhang
    /// (ADR-0038 D3).
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>Coverage is ARITHMETIC, never measurement.</b> An Excel <c>AddLine</c> has a
    /// degenerate bounding box -- its <c>Top + Height</c> is a point, not the stroke --
    /// so a test asking whether the band was "covered" by reading the line's bounds
    /// reports <c>False</c> for a perfectly drawn chart. The live probe
    /// (<c>scripts/probe-anchor-row-height.ps1</c>) printed <c>covered=False</c> for that
    /// reason and it is an ARTEFACT, not a finding. Coverage comes from
    /// <c>Line.Weight</c>, which is what the assertions below use.
    /// </para>
    /// <para>
    /// The line is centred on the boundary, so it covers
    /// <c>bottom - w/2 .. bottom + w/2</c>. The bands must END INSIDE that: stopping
    /// short leaves the overhang visible, overshooting paints past the very line meant
    /// to close it.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_closing_line_covers_the_band_overhang_by_arithmetic_not_measurement()
    {
        FrameBandsCreationOutcome outcome = Build(CreateRequest(0.5, 0.25));
        Assert.True(outcome.Succeeded, $"Refused: {outcome.Refusal}");

        SceneLine closing = Assert.Single(
            outcome.Result!.Primitives.OfType<SceneLine>(),
            line => string.Equals(line.PrimitiveId, "chart:plot-closing", StringComparison.Ordinal));

        // It terminates the stacks, so it sits at the FRAME layer, above the bands
        // (AlternateBand 10) and the grid (Grid 20).
        Assert.Equal(ZLayer.Frame, closing.ZLayer);

        SceneRect band = outcome.Result!.Primitives.OfType<SceneRect>()
            .Single(rect => rect.PrimitiveId.StartsWith("chart:band:", StringComparison.Ordinal));
        // The line's width IS its coverage, so it is read from the resolved style rather than
        // assumed. A style with no outline width would have no coverage at all, which is
        // why the assertion below would then fail loudly instead of silently passing.
        Assert.True(closing.Style.OutlineWidthPt is { } width && width > 0);
        double halfLine = closing.Style.OutlineWidthPt!.Value / 2;

        // The line's centre IS the bands' bottom: one expression of the boundary, so
        // they cannot drift apart.
        Assert.Equal(closing.From!.Y, closing.To!.Y, 6);
        Assert.Equal(band.Bounds.Bottom, closing.From.Y, 6);

        // The overhang is COVERED: the line reaches above AND below it.
        Assert.True(
            closing.From.Y - halfLine <= band.Bounds.Bottom,
            $"The closing line stops {band.Bounds.Bottom - (closing.From.Y - halfLine)}pt above the band bottom.");
        Assert.True(
            closing.From.Y + halfLine >= band.Bounds.Bottom,
            "The closing line must reach past the band bottom, or the overhang is visible.");

        // It spans the PLOT's full width, not one band's: it closes the whole stack of
        // bands and grid lines, so it is asserted against the plot bounds rather than
        // against whichever band the Single above happened to return.
        RectD plot = CreateRequest(0.5, 0.25).PlotBounds;
        Assert.Equal(plot.Left, closing.From.X, 6);
        Assert.Equal(plot.Right, closing.To!.X, 6);
    }

    /// <summary>
    /// The bottom FRAME stays at the chart's own bottom edge, distinct from the closing
    /// line (ADR-0038 D4/D5).
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>Two lines is the intent, not an artefact.</b> The closing line closes the
    /// stacks at the data boundary; the frame closes the chart at its padding-row
    /// margin. Collapsing them into one line at the data boundary was the rejected
    /// alternative -- it would have made the chart's bottom margin 0.75pt instead of the
    /// padding row's height and broken ADR-0031 D2 in the process.
    /// </para>
    /// <para>
    /// Asserted as an ORDERING (frame below closing) rather than as two magic numbers,
    /// so it survives a retune of either token while still failing a collapsed frame.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_bottom_frame_and_the_closing_line_are_two_distinct_lines()
    {
        FrameBandsCreationOutcome outcome = Build(CreateRequest(0.5, 0.25));
        Assert.True(outcome.Succeeded, $"Refused: {outcome.Refusal}");

        SceneLine frameBottom = outcome.Result!.Primitives.OfType<SceneLine>()
            .Single(line => string.Equals(line.PrimitiveId, "chart:frame:bottom", StringComparison.Ordinal));
        SceneLine closing = outcome.Result!.Primitives.OfType<SceneLine>()
            .Single(line => string.Equals(line.PrimitiveId, "chart:plot-closing", StringComparison.Ordinal));

        // Unchanged: the frame closes the chart at its own bounds, not at the data.
        Assert.Equal(outcome.Result.Geometry.ChartBounds.Bottom, frameBottom.From!.Y, 6);

        // The frame closes BELOW the data boundary, or the bottom margin collapses.
        Assert.True(
            frameBottom.From.Y > closing.From!.Y,
            "The chart frame must close below the plot's data boundary.");
    }

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
    public void A_narrow_month_band_steps_the_MMM_label_down_to_MM()
    {
        // ADR-0039's fit ladder. A full year across a 300pt plot gives
        // each month 25pt; the x10 measurer makes "Jan" 30pt, so the
        // selected form cannot fit while the interval is still above the
        // 18pt suppression minimum. The two-digit form is emitted instead.
        FrameBandsResult result = FrameBandsBuilder
            .TryBuild(
                CreateFullYearRequest(plotRightPt: 400, minimumHeaderLabelWidthPt: 18),
                new FixedMeasurer(text => text.Length * 10.0))
            .Result!;

        SceneText[] periodLabels = [.. result.Primitives.OfType<SceneText>()
            .Where(text => text.PrimitiveId.StartsWith("chart:period:", StringComparison.Ordinal))];
        Assert.Equal(
            Enumerable.Range(1, 12).Select(month => month.ToString("D2", CultureInfo.InvariantCulture)),
            periodLabels.Select(text => text.Text));
        // The label's identity is role-derived (R3.17), so the step down
        // does not move the reconciliation key a renderer matches on.
        Assert.Equal(
            "chart:period:2024-01-01:label",
            periodLabels.Single(text => text.Text == "01").PrimitiveId);
    }

    [Fact]
    public void A_wide_month_band_keeps_the_MMM_label()
    {
        // The control for the ladder above: the same x10 measurer, but
        // each month is 100pt wide, so "Jan" (30pt) fits and the
        // selected format is emitted verbatim.
        FrameBandsResult result = FrameBandsBuilder
            .TryBuild(CreateRequest(), new FixedMeasurer(text => text.Length * 10.0))
            .Result!;

        Assert.Contains(result.Primitives.OfType<SceneText>(), text =>
            text.PrimitiveId.StartsWith("chart:period:", StringComparison.Ordinal) && text.Text == "Jan");
        Assert.DoesNotContain(result.Primitives.OfType<SceneText>(), text =>
            text.PrimitiveId.StartsWith("chart:period:", StringComparison.Ordinal) && text.Text == "01");
    }

    [Fact]
    public void The_MM_fallback_is_used_even_when_MM_measures_wider_than_the_band()
    {
        // The owner's ruling for the ladder's end condition: the step down
        // is unconditional. A 210pt plot gives each month 17.5pt, so even
        // "01" (20pt at the x10 measurer) measures wider than its interval;
        // it is still emitted and left to the existing host clipping.
        FrameBandsResult result = FrameBandsBuilder
            .TryBuild(
                CreateFullYearRequest(plotRightPt: 310, minimumHeaderLabelWidthPt: 10),
                new FixedMeasurer(text => text.Length * 10.0))
            .Result!;

        Assert.Equal(
            Enumerable.Range(1, 12).Select(month => month.ToString("D2", CultureInfo.InvariantCulture)),
            result.Primitives.OfType<SceneText>()
                .Where(text => text.PrimitiveId.StartsWith("chart:period:", StringComparison.Ordinal))
                .Select(text => text.Text));
    }

    [Theory]
    [InlineData(GanttTimeScale.Quarter, GanttPeriodLabelFormat.Quarter, 1, 12, 150, "Q1")]
    [InlineData(GanttTimeScale.Week, GanttPeriodLabelFormat.Week, 1, 1, 200, "W01")]
    [InlineData(GanttTimeScale.Month, GanttPeriodLabelFormat.MM, 1, 12, 400, "01")]
    public void The_step_down_applies_to_no_other_scale_or_format(
        GanttTimeScale scale,
        GanttPeriodLabelFormat format,
        int startMonth,
        int finishMonth,
        double plotRightPt,
        string expectedLabel)
    {
        // Each row picks a geometry where the selected label measures wider
        // than its interval at the x10 measurer ("Q1" 20pt against 12.5pt,
        // "W01" 30pt against 25pt) or is already at the format floor (MM).
        // None of these has a defined shorter form, so the label is verbatim.
        DateOnly finish = new(2024, finishMonth, DateTime.DaysInMonth(2024, finishMonth));
        TimeScale timeScale = TimeScale
            .TryCreate(new DateOnly(2024, startMonth, 1), finish, 100, plotRightPt)
            .Scale!;
        FrameBandsResult result = FrameBandsBuilder
            .TryBuild(
                CreateRequest() with
                {
                    TimeScale = timeScale,
                    Scale = scale,
                    PeriodLabelFormat = format,
                    PlotBounds = new RectD(100, 100, plotRightPt - 100, 100),
                    MinimumHeaderLabelWidthPt = 10,
                },
                new FixedMeasurer(text => text.Length * 10.0))
            .Result!;

        Assert.Contains(result.Primitives.OfType<SceneText>(), text =>
            text.PrimitiveId.StartsWith("chart:period:", StringComparison.Ordinal) && text.Text == expectedLabel);
    }

    [Fact]
    public void Suppression_still_wins_over_the_step_down()
    {
        // The ladder runs only where ShowLabel is set: a 17.5pt interval
        // below a 20pt minimum suppresses its label rather than stepping
        // down to a form that would not have been shown either.
        FrameBandsCreationOutcome outcome = FrameBandsBuilder
            .TryBuild(
                CreateFullYearRequest(plotRightPt: 310, minimumHeaderLabelWidthPt: 20),
                new FixedMeasurer(text => text.Length * 10.0));

        Assert.True(outcome.Succeeded, $"Refused: {outcome.Refusal}");
        Assert.DoesNotContain(outcome.Result!.Primitives.OfType<SceneText>(), text =>
            text.PrimitiveId.StartsWith("chart:period:", StringComparison.Ordinal));
        Assert.Contains(outcome.Result.Warnings, warning => warning.Code == "AllPeriodLabelsSuppressed");
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
    [InlineData(GanttTimeScale.Week, GanttPeriodLabelFormat.Week, "W01")]
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

    private static FrameBandsRequest CreateRequest(
        double plotBandHeaderOverlapPt = 0.5,
        double chartAnchorRowHeightPt = 0.25)
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
            plotBandHeaderOverlapPt,
            chartAnchorRowHeightPt,
            _theme
        );
    }

    /// <summary>
    /// A Month-scale request spanning all of 2024, whose plot width and
    /// suppression minimum are the two knobs the ADR-0039 ladder tests
    /// turn: <paramref name="plotRightPt"/> sets each month's visible
    /// width (the plot is 100pt to <paramref name="plotRightPt"/>), and
    /// <paramref name="minimumHeaderLabelWidthPt"/> sets the floor below
    /// which a label suppresses instead of stepping down.
    /// </summary>
    private static FrameBandsRequest CreateFullYearRequest(
        double plotRightPt,
        double minimumHeaderLabelWidthPt)
    {
        TimeScale scale = TimeScale
            .TryCreate(new DateOnly(2024, 1, 1), new DateOnly(2024, 12, 31), 100, plotRightPt)
            .Scale!;
        return CreateRequest() with
        {
            TimeScale = scale,
            PlotBounds = new RectD(100, 100, plotRightPt - 100, 100),
            MinimumHeaderLabelWidthPt = minimumHeaderLabelWidthPt,
        };
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
