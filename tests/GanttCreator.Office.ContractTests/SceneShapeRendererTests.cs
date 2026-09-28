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
    /// A milestone diamond is a four-point polygon that becomes a diamond
    /// auto-shape request at the points' bounding box, with the points still
    /// carried so the translation is auditable. The box is what the host object
    /// is placed from, and it takes the same uniform delta as every other family.
    /// </summary>
    [Fact]
    public void A_four_point_diamond_becomes_a_diamond_request_at_its_translated_box()
    {
        var renderer = new SceneShapeRenderer(new ChartOriginDelta(12, 12));
        PointD[] diamond = [new(50, 40), new(60, 50), new(50, 60), new(40, 50)];
        var scene = SceneOf(
            new ScenePolygon("row-1:marker", RowOwner, ZLayer.Milestone, diamond, Style));

        SceneTranslationOutcome outcome = renderer.Translate(scene);
        Assert.True(outcome.Complete);

        OfficeShapeRequest marker = Assert.Single(outcome.Requests);
        Assert.Equal(OfficeShapeKind.Diamond, marker.Kind);
        Assert.Equal(ZLayer.Milestone, marker.ZLayer);

        // The scene box (40,40,20,20) translated by the delta.
        Assert.Equal(new RectD(52, 52, 20, 20), marker.Geometry.Bounds);
        Assert.Equal(
            [new PointD(62, 52), new PointD(72, 62), new PointD(62, 72), new PointD(52, 62)],
            marker.Geometry.Points);
    }

    /// <summary>
    /// A polygon the host's diamond auto-shape cannot represent is REFUSED with a
    /// typed reason, never approximated. A rotated square, a rectangle, and an
    /// off-centre vertex are all not-diamonds, and silently drawing any of them
    /// as a diamond would change the entity the scene resolved.
    /// </summary>
    /// <param name="shape">Which non-diamond case to build.</param>
    [Theory]
    [InlineData("rectangle")]
    [InlineData("rotated-square")]
    [InlineData("oblique-square")]
    [InlineData("off-centre-vertex")]
    [InlineData("triangle")]
    [InlineData("degenerate-point")]
    public void A_polygon_that_is_not_a_symmetric_diamond_is_refused(string shape)
    {
        var renderer = new SceneShapeRenderer(ChartOriginDelta.Identity);
        var scene = SceneOf(
            new ScenePolygon(
                "row-1:marker",
                RowOwner,
                ZLayer.Milestone,
                NotADiamond(shape),
                Style));

        SceneTranslationOutcome outcome = renderer.Translate(scene);

        Assert.False(outcome.Complete);
        Assert.Empty(outcome.Requests);
        Assert.Empty(outcome.Deferred);
        Assert.Equal(
            [("row-1:marker", SceneTranslationRefusalReason.NotADiamond)],
            outcome.Refusals.Select(r => (r.PrimitiveId, r.Reason)));
    }

    /// <summary>
    /// Builds a valid quadrilateral that is nonetheless not a diamond. The cases
    /// are named rather than passed as data so each one appears as its own row
    /// in the test output, which is what makes a newly-refused shape visible.
    /// </summary>
    /// <param name="shape">The case name.</param>
    /// <returns>The polygon points.</returns>
    private static PointD[] NotADiamond(string shape) =>
        shape switch
        {
            // Four points, but not one per edge midpoint of the box.
            "rectangle" => [new(40, 40), new(80, 40), new(80, 60), new(40, 60)],

            // A square rotated 45 degrees: its vertices are on the box's
            // corners, not its edge midpoints, so it is not a diamond.
            "rotated-square" => [new(40, 40), new(60, 40), new(60, 60), new(40, 60)],

            // A square rotated 30 degrees, which lands no vertex on any edge
            // midpoint.
            "oblique-square" => [new(50, 30), new(68, 45), new(50, 60), new(32, 45)],

            // Three vertices are correct; the fourth is off the opposite
            // edge's midpoint.
            "off-centre-vertex" => [new(50, 40), new(60, 50), new(50, 60), new(45, 50)],

            // A triangle cannot be a diamond.
            "triangle" => [new(40, 40), new(60, 40), new(50, 60)],

            // Every vertex coincides, so the box has no extent.
            _ => [new(50, 50), new(50, 50), new(50, 50), new(50, 50)],
        };

    /// <summary>
    /// A diamond in a different draw order is still a diamond: the host object
    /// is defined by the box, so vertex order must not decide whether the
    /// primitive renders.
    /// </summary>
    [Fact]
    public void A_diamond_is_recognised_whatever_its_vertex_draw_order()
    {
        var renderer = new SceneShapeRenderer(ChartOriginDelta.Identity);
        PointD[] rotated = [new(60, 50), new(50, 60), new(40, 50), new(50, 40)];
        var scene = SceneOf(
            new ScenePolygon("row-1:marker", RowOwner, ZLayer.Milestone, rotated, Style));

        SceneTranslationOutcome outcome = renderer.Translate(scene);

        Assert.True(outcome.Complete);
        Assert.Empty(outcome.Refusals);
        Assert.Equal(new RectD(40, 40, 20, 20), Assert.Single(outcome.Requests).Geometry.Bounds);
    }

    /// <summary>
    /// The tip-to-tip extent the scene resolved is preserved exactly, so a
    /// milestone diamond is never drawn at a size the scene did not ask for.
    /// </summary>
    [Fact]
    public void A_diamonds_tip_to_tip_extent_survives_translation()
    {
        const double Size = 20d;
        const double Centre = 100d;
        var renderer = new SceneShapeRenderer(new ChartOriginDelta(30, 30));

        // A diamond whose horizontal AND vertical tip-to-tip extent is Size*2,
        // matching the entity guide's MilestoneSizePt contract.
        PointD[] diamond =
        [
            new(Centre, Centre - Size),
            new(Centre + Size, Centre),
            new(Centre, Centre + Size),
            new(Centre - Size, Centre),
        ];
        var scene = SceneOf(
            new ScenePolygon("row-1:marker", RowOwner, ZLayer.Milestone, diamond, Style));

        RectD bounds = Assert.Single(renderer.Translate(scene).Requests).Geometry.Bounds!.Value;

        Assert.Equal(Size * 2, bounds.Width, 10);
        Assert.Equal(Size * 2, bounds.Height, 10);
    }

    /// <summary>
    /// A group primitive owns no geometry of its own, so it is reported as
    /// deferred rather than translated into an empty shape. Grouping itself is
    /// Phase 6's export concern.
    /// </summary>
    [Fact]
    public void A_group_primitive_is_reported_as_deferred_rather_than_drawn()
    {
        var renderer = new SceneShapeRenderer(ChartOriginDelta.Identity);
        var scene = SceneOf(
            Rect("row-1:bar", new RectD(10, 20, 100, 30)),
            new SceneGroup("row-1:group", RowOwner, ZLayer.ActivityBody, ["row-1:bar"]));

        SceneTranslationOutcome outcome = renderer.Translate(scene);

        Assert.False(outcome.Complete);
        Assert.Single(outcome.Requests);
        Assert.Equal(
            [("row-1:group", ScenePrimitiveKind.Group)],
            outcome.Deferred.Select(deferred => (deferred.PrimitiveId, deferred.Kind)));
    }

    /// <summary>
    /// Every primitive family is now renderable, so a mixed scene translates
    /// completely. This is the row's completion pin: a scene holding a bar, a
    /// grid line, a milestone diamond, and a label yields four requests and no
    /// deferrals.
    /// </summary>
    [Fact]
    public void A_mixed_scene_of_every_family_translates_completely()
    {
        var renderer = new SceneShapeRenderer(ChartOriginDelta.Identity);
        var scene = SceneOf(
            Rect("row-1:bar", new RectD(10, 20, 100, 30)),
            Line("chart:grid:0", new PointD(10, 60), new PointD(90, 60)),
            new ScenePolygon(
                "row-2:marker",
                RowOwner,
                ZLayer.Milestone,
                [new PointD(50, 50), new PointD(60, 60), new PointD(40, 60), new PointD(50, 70)],
                Style),
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

        // One shape per primitive, each of the right family.
        Assert.Equal(4, outcome.Requests.Count);
        Assert.Equal(
            [OfficeShapeKind.Line, OfficeShapeKind.Rectangle, OfficeShapeKind.Diamond, OfficeShapeKind.TextBox],
            outcome.Requests.Select(request => request.Kind));
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
