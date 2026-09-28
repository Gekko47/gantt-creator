using GanttCreator.Core.Scene;

namespace GanttCreator.Office;

/// <summary>The scene primitive families the live renderer must be able to draw.</summary>
public enum ScenePrimitiveKind
{
    /// <summary>A rectangle primitive.</summary>
    Rect = 0,

    /// <summary>A line primitive.</summary>
    Line = 1,

    /// <summary>A polygon primitive.</summary>
    Polygon = 2,

    /// <summary>A text primitive.</summary>
    Text = 3,

    /// <summary>A group primitive, which owns no geometry of its own.</summary>
    Group = 4,
}

/// <summary>One scene primitive the renderer did not translate, with its family.</summary>
/// <param name="PrimitiveId">The stable role-derived primitive identifier.</param>
/// <param name="Kind">The primitive family.</param>
public sealed record DeferredPrimitive(string PrimitiveId, ScenePrimitiveKind Kind);

/// <summary>The typed result of translating a scene into shape requests.</summary>
/// <param name="Requests">The translated requests, in the scene's back-to-front order.</param>
/// <param name="Deferred">The primitives this renderer did not translate.</param>
/// <remarks>
/// A deferred primitive is <em>not</em> an error. Phase 4 introduces the
/// families one row at a time - R4.3 rectangles and lines, R4.4 text, R4.5
/// polygons - so a scene legitimately contains families a not-yet-written
/// renderer cannot draw. They are reported by name rather than dropped silently,
/// because a silently missing shape is indistinguishable from a correct render
/// until a user looks at the chart.
/// </remarks>
public sealed record SceneTranslationOutcome(
    IReadOnlyList<OfficeShapeRequest> Requests,
    IReadOnlyList<DeferredPrimitive> Deferred)
{
    /// <summary>Gets whether every primitive in the scene was translated.</summary>
    public bool Complete => Deferred.Count == 0;
}

/// <summary>
/// Translates resolved scene primitives into the primitive shape requests the
/// live shape writer consumes.
/// </summary>
/// <remarks>
/// <para>
/// This renderer <strong>consumes</strong> and never recalculates. It reads the
/// geometry, style, identity, and layer the scene already resolved and moves
/// nothing else: no date is re-derived, no lane position recomputed, no colour
/// chosen, and no label side re-selected. It holds no reference to the worksheet
/// at all, so a table read is not merely avoided by convention - it is
/// unrepresentable.
/// </para>
/// <para>
/// The output order is the <see cref="GanttScene"/>'s own order, which sorts by
/// layer, then activity subtype, then lane, stack, sort, and identifier. The
/// renderer does not re-sort: R4.5 applies the scene's order rather than
/// deriving one, and a second sort here would be a second, divergent contract.
/// </para>
/// <para>
/// Every coordinate passes through one uniform <see cref="ChartOriginDelta"/>.
/// That is structural rather than conventional: the host silently clamps a
/// negative offset to zero, so a missing or partial translation corrupts the
/// chart invisibly.
/// </para>
/// </remarks>
public sealed class SceneShapeRenderer(ChartOriginDelta originDelta)
{
    /// <summary>Gets the translation applied to every primitive.</summary>
    public ChartOriginDelta OriginDelta => originDelta;

    /// <summary>Translates every primitive in a scene, in the scene's own order.</summary>
    /// <param name="scene">The resolved scene.</param>
    /// <returns>The translated requests plus every primitive not translated.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="scene"/> is <see langword="null"/>.</exception>
    public SceneTranslationOutcome Translate(GanttScene scene)
    {
        ArgumentNullException.ThrowIfNull(scene);

        List<OfficeShapeRequest> requests = [];
        List<DeferredPrimitive> deferred = [];
        foreach (ScenePrimitive primitive in scene.Primitives)
        {
            OfficeShapeRequest? request = TranslatePrimitive(primitive);
            if (request is not null)
            {
                requests.Add(request);
            }
            else
            {
                deferred.Add(new DeferredPrimitive(primitive.PrimitiveId, KindOf(primitive)));
            }
        }

        return new SceneTranslationOutcome(requests, deferred);
    }

    /// <summary>Translates one primitive.</summary>
    /// <param name="primitive">The scene primitive.</param>
    /// <param name="request">The translated request when successful.</param>
    /// <returns>
    /// <see langword="true"/> when the primitive's family is renderable; a group
    /// or a family a later Phase-4 row owns returns <see langword="false"/>.
    /// </returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="primitive"/> is <see langword="null"/>.</exception>
    public bool TryTranslate(ScenePrimitive primitive, out OfficeShapeRequest? request)
    {
        request = TranslatePrimitive(primitive);
        return request is not null;
    }

    private OfficeShapeRequest? TranslatePrimitive(ScenePrimitive primitive)
    {
        ArgumentNullException.ThrowIfNull(primitive);

        return primitive switch
        {
            SceneRect rect => BuildRect(rect),
            SceneLine line => BuildLine(line),
            ScenePolygon polygon => BuildPolygon(polygon),
            SceneText text => BuildText(text),
            _ => null,
        };
    }

    /// <summary>Determines the family a primitive belongs to.</summary>
    /// <param name="primitive">The scene primitive.</param>
    /// <returns>The primitive family.</returns>
    /// <remarks>
    /// The default arm names <see cref="ScenePrimitiveKind.Group"/> deliberately:
    /// a new concrete primitive added later must not silently classify itself as
    /// something already handled.
    /// </remarks>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="primitive"/> is <see langword="null"/>.</exception>
    public static ScenePrimitiveKind KindOf(ScenePrimitive primitive)
    {
        ArgumentNullException.ThrowIfNull(primitive);

        return primitive switch
        {
            SceneRect => ScenePrimitiveKind.Rect,
            SceneLine => ScenePrimitiveKind.Line,
            ScenePolygon => ScenePrimitiveKind.Polygon,
            SceneText => ScenePrimitiveKind.Text,
            _ => ScenePrimitiveKind.Group,
        };
    }

    private OfficeShapeRequest BuildRect(SceneRect rect) =>
        new(
            rect.PrimitiveId,
            OfficeShapeKind.Rectangle,
            new OfficeShapeGeometry(Bounds: originDelta.Apply(rect.Bounds)),
            rect.ZLayer);

    /// <summary>Builds a plain-line request, preserving the scene's endpoint order.</summary>
    /// <remarks>
    /// A right-to-left or bottom-up line keeps its direction here: the host's
    /// <c>AddLine</c> takes absolute endpoints, so the scene's own
    /// <c>From</c>/<c>To</c> is expressible exactly. Normalising to a
    /// min/absolute-extent bounding box belongs to an in-place <em>update</em>,
    /// where <c>Shape</c> exposes only <c>Left</c>/<c>Top</c>/<c>Width</c>/
    /// <c>Height</c> and direction is unrecoverable. Doing it here would silently
    /// discard a resolved endpoint.
    /// </remarks>
    /// <param name="line">The scene line.</param>
    /// <returns>The translated shape request.</returns>
    private OfficeShapeRequest BuildLine(SceneLine line) =>
        new(
            line.PrimitiveId,
            OfficeShapeKind.Line,
            new OfficeShapeGeometry(
                From: originDelta.Apply(line.From),
                To: originDelta.Apply(line.To)),
            line.ZLayer);

    /// <summary>Builds a freeform request from the polygon's ordered points.</summary>
    /// <remarks>
    /// R4.5 D1. The points are translated by the same uniform delta as every other
    /// family and their draw order is preserved exactly: a milestone is a
    /// four-point polygon whose tip-to-tip extent the entity guide fixes, and
    /// reordering or normalising the vertices here would change the rendered
    /// shape rather than translate it. Vertices are absolute sheet points, which is
    /// what the R4.5 Step-0 probe established for the host's freeform builder.
    /// </remarks>
    /// <param name="polygon">The scene polygon.</param>
    /// <returns>The translated shape request.</returns>
    private OfficeShapeRequest BuildPolygon(ScenePolygon polygon) =>
        new(
            polygon.PrimitiveId,
            OfficeShapeKind.Polygon,
            new OfficeShapeGeometry(
                Points: [.. polygon.Points.Select(originDelta.Apply)]),
            polygon.ZLayer);

    /// <summary>Builds a text request, consuming the scene's resolved text verbatim.</summary>
    /// <remarks>
    /// <para>
    /// R4.4 D1: this consumes and never re-measures. <see cref="SceneText.Text"/>
    /// is written exactly as the scene produced it - including the single-
    /// character ellipsis R3.6's overflow policy already applied at
    /// <c>MaximumExternalLabelWidthPt</c> - and the shape is positioned at the
    /// resolved <see cref="SceneText.TextBounds"/>. No measuring API is called, no
    /// truncation is applied here, and the label side the scene chose is not
    /// revisited, even when the opposite side would look emptier.
    /// </para>
    /// <para>
    /// The alignment is <see cref="SceneText.Alignment"/>, which ADR-0018 D4's
    /// rule already resolved in the scene through
    /// <c>LabelPlanner.TextAlignmentFor</c>. This method copies it; it does not
    /// reimplement the rule.
    /// </para>
    /// <para>
    /// Typography is copied from the resolved <see cref="SceneStyle"/> tokens.
    /// Nothing is invented between token and property, and an absent token stays
    /// absent rather than being defaulted - R4.6 owns the full style matrix.
    /// </para>
    /// </remarks>
    /// <param name="text">The scene text.</param>
    /// <returns>The translated shape request.</returns>
    private OfficeShapeRequest BuildText(SceneText text) =>
        new(
            text.PrimitiveId,
            OfficeShapeKind.TextBox,
            new OfficeShapeGeometry(Bounds: originDelta.Apply(text.TextBounds)),
            text.ZLayer,
            FillColour: text.Style.FillColour,
            StrokeColour: text.Style.StrokeColour,
            LineWidthPt: text.Style.OutlineWidthPt,
            FontFamily: text.Style.FontFamily,
            FontSizePt: text.Style.FontSizePt,
            Bold: text.Style.Bold,
            Text: text.Text,
            Alignment: text.Alignment);
}
