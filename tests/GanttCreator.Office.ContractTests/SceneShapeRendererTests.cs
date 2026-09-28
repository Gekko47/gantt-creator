using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Xunit.Abstractions;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R4.3 contract tests for the translation from resolved scene primitives to the
/// shape requests the live writer consumes. No Excel: the renderer holds no
/// worksheet reference, which is what makes D4's no-table-read guarantee
/// structural rather than a convention.
/// </summary>
public class SceneShapeRendererTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    private static SceneOwnerId ChartOwner => SceneOwnerId.Chart;

    private static SceneOwnerId RowOwner => SceneOwnerId.ForRow(GanttRowId.New());

    private static SceneStyle Style => new("Default");

    private static GanttScene SceneOf(params ScenePrimitive[] primitives) =>
        GanttScene
            .TryCreate(new RectD(0, 0, 100, 100), new RectD(10, 10, 80, 80), primitives, [])
            .Scene ?? throw new InvalidOperationException("The fixture scene failed to validate.");

    private static SceneRect Rect(string id, RectD bounds, ZLayer layer = ZLayer.ActivityBody) =>
        new(id, RowOwner, layer, bounds, Style);

    private static SceneLine Line(
        string id,
        PointD from,
        PointD to,
        ZLayer layer = ZLayer.Grid) =>
        new(id, RowOwner, layer, from, to, Style);

    [Fact]
    public void A_rectangle_becomes_a_rectangle_request_named_for_its_primitive_id()
    {
        var renderer = new SceneShapeRenderer(ChartOriginDelta.Identity);

        bool translated = renderer.TryTranslate(
            Rect("row-1:bar", new RectD(10, 20, 100, 30)),
            out OfficeShapeRequest? request);

        Assert.True(translated);
        Assert.NotNull(request);
        Assert.Equal("row-1:bar", request.PrimitiveId);
        Assert.Equal(OfficeShapeKind.Rectangle, request.Kind);
        Assert.Equal(ZLayer.ActivityBody, request.ZLayer);
        Assert.Equal(new RectD(10, 20, 100, 30), request.Geometry.Bounds);
    }

    [Fact]
    public void A_line_becomes_a_line_request_carrying_both_resolved_endpoints()
    {
        var renderer = new SceneShapeRenderer(ChartOriginDelta.Identity);

        bool translated = renderer.TryTranslate(
            Line("chart:grid:0", new PointD(5, 10), new PointD(95, 10)),
            out OfficeShapeRequest? request);

        Assert.True(translated);
        Assert.NotNull(request);
        Assert.Equal(OfficeShapeKind.Line, request.Kind);
        Assert.Equal(new PointD(5, 10), request.Geometry.From);
        Assert.Equal(new PointD(95, 10), request.Geometry.To);
    }

    /// <summary>
    /// A right-to-left or bottom-up line is resolved data, not a modelling
    /// error. The host's AddLine takes absolute endpoints, so the direction is
    /// expressible at creation and must survive translation; normalising it to a
    /// bounding box here would silently discard a resolved endpoint.
    /// </summary>
    [Fact]
    public void A_right_to_left_line_keeps_its_endpoint_order_rather_than_being_normalised()
    {
        var renderer = new SceneShapeRenderer(ChartOriginDelta.Identity);

        _ = renderer.TryTranslate(
            Line("chart:delineator:1", new PointD(90, 40), new PointD(10, 40)),
            out OfficeShapeRequest? request);

        Assert.NotNull(request);
        Assert.Equal(new PointD(90, 40), request.Geometry.From);
        Assert.Equal(new PointD(10, 40), request.Geometry.To);

        // The normalisation belongs to the in-place update path, where the host
        // exposes only Left/Top/Width/Height. Asserting the request keeps the
        // direction proves the two paths cannot be confused later.
        Assert.True(request.Geometry.From!.Value.X > request.Geometry.To!.Value.X);
    }

    /// <summary>
    /// The zero-origin translation is the row's load-bearing rule, and the R4.3
    /// probe is why: a live host SILENTLY clamps a negative offset to zero
    /// rather than rejecting it. This test pins the shape of the guarantee - one
    /// delta applied to every primitive kind - which is what makes a missing
    /// translation impossible to introduce later without a red test.
    /// </summary>
    [Fact]
    public void One_uniform_delta_is_applied_to_every_primitive_kind()
    {
        var delta = new ChartOriginDelta(25, 40);
        var renderer = new SceneShapeRenderer(delta);
        var scene = SceneOf(
            Rect("row-1:bar", new RectD(-5, -5, 100, 30), ZLayer.ActivityBody),
            Line("chart:grid:0", new PointD(-5, 50), new PointD(95, 50), ZLayer.Grid));

        SceneTranslationOutcome outcome = renderer.Translate(scene);

        Assert.Equal(2, outcome.Requests.Count);

        // Looked up by identifier, not by index: the scene orders by layer, so a
        // positional assertion would re-state the ordering test's job here and
        // fail for the wrong reason if the layer table ever changed.
        OfficeShapeRequest bar = outcome.Requests.Single(r => r.PrimitiveId == "row-1:bar");
        OfficeShapeRequest grid = outcome.Requests.Single(r => r.PrimitiveId == "chart:grid:0");

        RectD rectBounds = bar.Geometry.Bounds!.Value;
        Assert.Equal(new RectD(20, 35, 100, 30), rectBounds);

        Assert.Equal(new PointD(20, 90), grid.Geometry.From);
        Assert.Equal(new PointD(120, 90), grid.Geometry.To);

        // The relationship inside the scene is unchanged by the translation: a
        // gap measured in the scene is still the same gap on the worksheet.
        double sceneGap = 50d - (-5d);
        double hostGap = grid.Geometry.From!.Value.Y - rectBounds.Y;
        Assert.Equal(sceneGap, hostGap, 10);
    }

    /// <summary>
    /// The translation is computed from the chart bounds, so a chart whose
    /// content starts at the origin - whose bounds therefore start at minus one
    /// outer padding - produces a positive delta that lands the bounds' top-left
    /// corner exactly on the worksheet origin.
    /// </summary>
    [Fact]
    public void The_delta_is_computed_from_the_chart_bounds_so_a_negative_origin_becomes_the_worksheet_origin()
    {
        var chartBounds = new RectD(-12, -12, 400, 300);

        ChartOriginDelta delta = ChartOriginDelta.ForChartBounds(chartBounds);

        RectD placed = delta.Apply(chartBounds);
        Assert.Equal(0d, placed.X);
        Assert.Equal(0d, placed.Y);

        // The extents are untouched: the delta moves the origin, it does not
        // scale or pad the chart.
        Assert.Equal(chartBounds.Width, placed.Width);
        Assert.Equal(chartBounds.Height, placed.Height);
    }

    /// <summary>
    /// Positive test for the delta's own validator. A non-finite component would
    /// silently produce a non-finite shape coordinate, which the host then
    /// clamps - the exact silent corruption the delta exists to prevent.
    /// </summary>
    [Theory]
    [InlineData(double.NaN, 0d)]
    [InlineData(0d, double.NaN)]
    [InlineData(double.PositiveInfinity, 0d)]
    [InlineData(0d, double.NegativeInfinity)]
    public void A_non_finite_delta_is_refused(double dx, double dy)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new ChartOriginDelta(dx, dy));
    }

    /// <summary>
    /// A primitive whose family a later Phase-4 row owns is reported by name
    /// rather than dropped. Silently omitting a shape is indistinguishable from a
    /// correct render until a user looks at the chart, so the deferral is part of
    /// the contract.
    /// </summary>
    [Fact]
    public void A_primitive_family_this_renderer_cannot_draw_is_reported_rather_than_silently_dropped()
    {
        var renderer = new SceneShapeRenderer(ChartOriginDelta.Identity);
        var scene = SceneOf(
            Rect("row-1:bar", new RectD(10, 20, 100, 30)),
            new ScenePolygon(
                "row-2:marker",
                RowOwner,
                ZLayer.Milestone,
                [new PointD(50, 50), new PointD(60, 60), new PointD(40, 60), new PointD(50, 70)],
                Style));

        SceneTranslationOutcome outcome = renderer.Translate(scene);

        Assert.False(outcome.Complete);
        Assert.Single(outcome.Requests);
        Assert.Equal(
            [("row-2:marker", ScenePrimitiveKind.Polygon)],
            outcome.Deferred.Select(deferred => (deferred.PrimitiveId, deferred.Kind)));
    }

    [Fact]
    public void A_scene_of_rectangles_lines_and_text_is_complete()
    {
        var renderer = new SceneShapeRenderer(ChartOriginDelta.Identity);
        var scene = SceneOf(
            Rect("row-1:bar", new RectD(10, 20, 100, 30)),
            Line("chart:grid:0", new PointD(10, 60), new PointD(90, 60)),
            new SceneText(
                "row-1:label",
                RowOwner,
                ZLayer.Label,
                "Site survey",
                new RectD(120, 20, 60, 12),
                Style,
                GanttTextAlignment.Right));

        SceneTranslationOutcome outcome = renderer.Translate(scene);

        Assert.True(outcome.Complete);
        Assert.Empty(outcome.Deferred);
        Assert.Equal(3, outcome.Requests.Count);
        _output.WriteLine("complete scene rendered " + outcome.Requests.Count + " shapes");
    }

    /// <summary>
    /// The renderer consumes the scene's own back-to-front order and never
    /// re-sorts. R4.5 applies that order deterministically, so a second sort
    /// here would be a second, divergent contract for the same question.
    /// </summary>
    [Fact]
    public void The_request_order_is_the_scene_order_and_is_not_re_sorted()
    {
        var renderer = new SceneShapeRenderer(ChartOriginDelta.Identity);

        // Supplied deliberately out of order: the scene sorts by layer, so the
        // grid line (Grid, 20) precedes the bar (ActivityBody, 40).
        var scene = SceneOf(
            Rect("row-1:bar", new RectD(10, 20, 100, 30), ZLayer.ActivityBody),
            Line("chart:grid:0", new PointD(10, 60), new PointD(90, 60), ZLayer.Grid));

        SceneTranslationOutcome outcome = renderer.Translate(scene);

        Assert.Equal(
            ["chart:grid:0", "row-1:bar"],
            outcome.Requests.Select(request => request.PrimitiveId));
    }

    /// <summary>
    /// D1 positive test: a scene label the planner placed on the Left stays on
    /// the Left, even though the Right would look emptier. The renderer copies
    /// the resolved bounds verbatim and never re-selects a side.
    /// </summary>
    [Fact]
    public void A_label_the_scene_placed_on_the_left_is_not_moved_to_the_right()
    {
        var renderer = new SceneShapeRenderer(ChartOriginDelta.Identity);
        var bounds = new RectD(10, 20, 60, 12);
        var scene = SceneOf(
            new SceneText(
                "row-1:label",
                RowOwner,
                ZLayer.Label,
                "Site survey",
                bounds,
                Style,
                GanttTextAlignment.Left));

        SceneTranslationOutcome outcome = renderer.Translate(scene);

        OfficeShapeRequest label = Assert.Single(outcome.Requests);
        Assert.Equal(bounds, label.Geometry.Bounds);
        Assert.Equal(GanttTextAlignment.Left, label.Alignment);
    }

    /// <summary>
    /// D3: the alignment mapping is a closed three-member table. Each scene
    /// alignment reaches the request unchanged, and an alignment outside the
    /// enum is refused rather than defaulted.
    /// </summary>
    [Theory]
    [InlineData(GanttTextAlignment.Left)]
    [InlineData(GanttTextAlignment.Centre)]
    [InlineData(GanttTextAlignment.Right)]
    public void Every_scene_alignment_reaches_the_request_unchanged(GanttTextAlignment alignment)
    {
        var renderer = new SceneShapeRenderer(ChartOriginDelta.Identity);
        var scene = SceneOf(
            new SceneText(
                "row-1:label",
                RowOwner,
                ZLayer.Label,
                "Site survey",
                new RectD(10, 20, 60, 12),
                Style,
                alignment));

        OfficeShapeRequest label = Assert.Single(renderer.Translate(scene).Requests);

        Assert.Equal(alignment, label.Alignment);
    }

    /// <summary>
    /// D4 positive test: an over-long label that the scene already truncated
    /// reaches the request byte-for-byte, ellipsis included. The renderer
    /// applies no truncation of its own, so the string it was given is the
    /// string it emits.
    /// </summary>
    [Fact]
    public void An_over_long_label_is_passed_through_verbatim_with_its_ellipsis_intact()
    {
        var renderer = new SceneShapeRenderer(ChartOriginDelta.Identity);
        const string Clipped = "Site survey and ground investigation phase one…";
        var scene = SceneOf(
            new SceneText(
                "row-1:label",
                RowOwner,
                ZLayer.Label,
                Clipped,
                new RectD(10, 20, 40, 12),
                Style,
                GanttTextAlignment.Left));

        OfficeShapeRequest label = Assert.Single(renderer.Translate(scene).Requests);

        Assert.Equal(Clipped, label.Text);
        Assert.EndsWith("…", label.Text, StringComparison.Ordinal);
    }

    /// <summary>
    /// D2: typography is copied from the resolved style tokens and nothing is
    /// invented between token and property. A token the scene did not resolve
    /// stays absent rather than being defaulted, because a defaulted font would
    /// read as a resolved one.
    /// </summary>
    [Fact]
    public void Typography_is_copied_from_the_style_tokens_and_an_absent_token_stays_absent()
    {
        var renderer = new SceneShapeRenderer(ChartOriginDelta.Identity);
        var styled = new SceneStyle("Label", fontFamily: "Aptos", fontSizePt: 9d, bold: true);
        var unstyled = new SceneStyle("Default");

        var scene = SceneOf(
            new SceneText(
                "row-1:styled",
                RowOwner,
                ZLayer.Label,
                "Styled",
                new RectD(10, 20, 60, 12),
                styled,
                GanttTextAlignment.Left),
            new SceneText(
                "row-2:unstyled",
                RowOwner,
                ZLayer.Label,
                "Unstyled",
                new RectD(10, 40, 60, 12),
                unstyled,
                GanttTextAlignment.Left));

        SceneTranslationOutcome outcome = renderer.Translate(scene);

        OfficeShapeRequest withTokens = outcome.Requests.Single(r => r.PrimitiveId == "row-1:styled");
        Assert.Equal("Aptos", withTokens.FontFamily);
        Assert.Equal(9d, withTokens.FontSizePt);
        Assert.True(withTokens.Bold);

        OfficeShapeRequest withoutTokens = outcome.Requests.Single(r => r.PrimitiveId == "row-2:unstyled");
        Assert.Null(withoutTokens.FontFamily);
        Assert.Null(withoutTokens.FontSizePt);
        Assert.Null(withoutTokens.Bold);
    }

    /// <summary>
    /// A text primitive's bounds take the same origin delta as every other
    /// family. A label translated differently from the bar it belongs to would
    /// detach the two.
    /// </summary>
    [Fact]
    public void A_text_primitive_takes_the_same_origin_delta_as_every_other_family()
    {
        var renderer = new SceneShapeRenderer(new ChartOriginDelta(12, 12));
        var scene = SceneOf(
            Rect("row-1:bar", new RectD(40, 30, 180, 24)),
            new SceneText(
                "row-1:label",
                RowOwner,
                ZLayer.Label,
                "Site survey",
                new RectD(230, 30, 60, 12),
                Style,
                GanttTextAlignment.Left));

        SceneTranslationOutcome outcome = renderer.Translate(scene);
        Assert.True(outcome.Complete);

        RectD bar = outcome.Requests.Single(r => r.PrimitiveId == "row-1:bar").Geometry.Bounds!.Value;
        RectD label = outcome.Requests.Single(r => r.PrimitiveId == "row-1:label").Geometry.Bounds!.Value;

        Assert.Equal(new RectD(52, 42, 180, 24), bar);
        Assert.Equal(new RectD(242, 42, 60, 12), label);

        // The label is still 10pt right of the bar's right edge, as in the scene.
        Assert.Equal(10d, label.Left - bar.Right, 10);
    }

    /// <summary>
    /// D1/D2: the renderer declares no measurement seam and holds no field, so a
    /// measure dependency cannot be added without failing this assertion. The
    /// scene already measured through R3.6's <c>ITextMetrics</c>; a second
    /// measurement on the rendering path would be a second, disagreeing answer.
    /// </summary>
    [Fact]
    public void The_renderer_declares_no_text_measurement_dependency()
    {
        Assert.DoesNotContain(
            typeof(SceneShapeRenderer).GetProperties(),
            property => property.PropertyType.Name.Contains("Metrics", StringComparison.Ordinal)
                || property.PropertyType.Name.Contains("Measure", StringComparison.Ordinal));

        Assert.Empty(typeof(SceneShapeRenderer).GetFields());
        Assert.Equal(
            [typeof(ChartOriginDelta)],
            typeof(SceneShapeRenderer).GetConstructors().Single().GetParameters()
                .Select(parameter => parameter.ParameterType));
    }

    /// <summary>
    /// D4: the renderer never recalculates. It is constructed with a delta and
    /// handed a scene - it holds no worksheet, table, or reader reference - so a
    /// table read during rendering is not a convention this class keeps but a
    /// dependency it cannot express. Asserted structurally rather than by
    /// observing an absence, because an absence cannot fail a test when the
    /// offending dependency is added later.
    /// </summary>
    [Fact]
    public void The_renderer_exposes_no_worksheet_or_table_dependency_to_read_from()
    {
        var constructor = typeof(SceneShapeRenderer).GetConstructors().Single();

        Assert.Equal(
            [typeof(ChartOriginDelta)],
            constructor.GetParameters().Select(parameter => parameter.ParameterType));

        // No injected port on the type's fields: a reader dependency added later
        // is a compile-visible change to this assertion.
        Assert.Empty(typeof(SceneShapeRenderer).GetFields());
        Assert.DoesNotContain(
            typeof(SceneShapeRenderer).GetProperties(),
            property => property.PropertyType.Name.Contains("Reader", StringComparison.Ordinal)
                || property.PropertyType.Name.Contains("Table", StringComparison.Ordinal)
                || property.PropertyType.Name.Contains("Worksheet", StringComparison.Ordinal));
    }
}
