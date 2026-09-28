using GanttCreator.Core;
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

/// <summary>Why the renderer refused a primitive it was asked to translate.</summary>
public enum SceneTranslationRefusalReason
{
    /// <summary>
    /// The polygon is not a symmetric axis-aligned diamond, so the host's
    /// diamond auto-shape cannot represent the points the scene resolved.
    /// </summary>
    /// <remarks>
    /// Refused rather than approximated. A rotated, skewed, or rectangular
    /// quadrilateral is not a diamond, and drawing it as one would silently
    /// change the entity the scene resolved - the exact substitution the entity
    /// guide's "renderers consume, never recalculate" rule forbids.
    /// </remarks>
    NotADiamond = 0,
}

/// <summary>One primitive the renderer refused, with the reason.</summary>
/// <param name="PrimitiveId">The stable role-derived primitive identifier.</param>
/// <param name="Reason">Why the primitive was refused.</param>
public sealed record SceneTranslationRefusal(string PrimitiveId, SceneTranslationRefusalReason Reason);

/// <summary>The typed result of translating a scene into shape requests.</summary>
/// <param name="Requests">The translated requests, in the scene's back-to-front order.</param>
/// <param name="Deferred">The primitives this renderer did not translate.</param>
/// <param name="Refusals">The primitives this renderer refused, with the reason.</param>
/// <remarks>
/// <para>
/// A deferred primitive is <em>not</em> an error. Phase 4 introduces the
/// families one row at a time, so a scene legitimately contains families a
/// not-yet-written renderer cannot draw. They are reported by name rather than
/// dropped silently, because a silently missing shape is indistinguishable from
/// a correct render until a user looks at the chart.
/// </para>
/// <para>
/// A <em>refusal</em> is different: the primitive is a family this renderer does
/// handle, but its content is not something the host object can represent. That
/// is a data fault, not a capability gap, and it is reported separately so a
/// caller can tell "not built yet" from "cannot be drawn".
/// </para>
/// </remarks>
public sealed record SceneTranslationOutcome(
    IReadOnlyList<OfficeShapeRequest> Requests,
    IReadOnlyList<DeferredPrimitive> Deferred,
    IReadOnlyList<SceneTranslationRefusal> Refusals)
{
    /// <summary>Gets whether every primitive in the scene was translated.</summary>
    public bool Complete => Deferred.Count == 0 && Refusals.Count == 0;
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
        List<SceneTranslationRefusal> refusals = [];
        foreach (ScenePrimitive primitive in scene.Primitives)
        {
            switch (primitive)
            {
                case ScenePolygon polygon
                    when !TryGetDiamondBounds(polygon.Points, out RectD _):
                    refusals.Add(new SceneTranslationRefusal(
                        primitive.PrimitiveId,
                        SceneTranslationRefusalReason.NotADiamond));
                    break;

                default:
                    OfficeShapeRequest? request = TranslatePrimitive(primitive);
                    if (request is not null)
                    {
                        requests.Add(request);
                    }
                    else
                    {
                        deferred.Add(new DeferredPrimitive(primitive.PrimitiveId, KindOf(primitive)));
                    }

                    break;
            }
        }

        return new SceneTranslationOutcome(requests, deferred, refusals);
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

            // A polygon reaches here only when the caller's guard already
            // confirmed it is a diamond; the bounding box is recomputed rather
            // than threaded through, so the two cannot disagree.
            ScenePolygon polygon
                when TryGetDiamondBounds(polygon.Points, out RectD diamondBounds) =>
                BuildDiamond(polygon, diamondBounds),

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

    /// <summary>
    /// Builds a rectangle request carrying the scene's resolved style.
    /// </summary>
    /// <remarks>
    /// R4.6: the style is mapped, not re-derived. Every fill, stroke, width, and
    /// hatch value here came from the style resolver by way of the scene; the
    /// renderer copies it and would be wrong to choose one.
    /// </remarks>
    /// <param name="rect">The scene rectangle.</param>
    /// <returns>The translated shape request.</returns>
    private OfficeShapeRequest BuildRect(SceneRect rect) =>
        new(
            rect.PrimitiveId,
            OfficeShapeKind.Rectangle,
            new OfficeShapeGeometry(Bounds: originDelta.Apply(rect.Bounds)),
            rect.ZLayer,
            FillColour: rect.Style.FillColour,
            StrokeColour: rect.Style.StrokeColour,
            LineWidthPt: rect.Style.OutlineWidthPt,
            HatchPattern: rect.Style.HatchPattern);

    /// <summary>
    /// Builds a plain-line request, preserving the scene's endpoint order and
    /// carrying the scene's resolved stroke.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A right-to-left or bottom-up line keeps its direction here: the host's
    /// <c>AddLine</c> takes absolute endpoints, so the scene's own
    /// <c>From</c>/<c>To</c> is expressible exactly. Normalising to a
    /// min/absolute-extent bounding box belongs to an in-place <em>update</em>,
    /// where <c>Shape</c> exposes only <c>Left</c>/<c>Top</c>/<c>Width</c>/
    /// <c>Height</c> and direction is unrecoverable. Doing it here would silently
    /// discard a resolved endpoint.
    /// </para>
    /// <para>
    /// A line's <see cref="SceneStyle.FillColour"/> is deliberately not carried. A
    /// line has no interior, and section 16's critical overlay and section 24's
    /// delineator are stroke-only entities; forwarding a fill for them would
    /// assert an appearance the entity does not have. The stroke is the whole
    /// style of a line and is carried in full.
    /// </para>
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
            line.ZLayer,
            StrokeColour: line.Style.StrokeColour,
            LineWidthPt: line.Style.OutlineWidthPt,
            HatchPattern: line.Style.HatchPattern);

    /// <summary>
    /// Determines whether four points form a symmetric axis-aligned diamond and,
    /// if so, returns the bounding box the host diamond auto-shape is placed at.
    /// </summary>
    /// <param name="points">The scene polygon's points.</param>
    /// <param name="bounds">The bounding box when the points are a diamond.</param>
    /// <returns><see langword="true"/> when the points form a diamond.</returns>
    /// <remarks>
    /// <para>
    /// The host draws a milestone as a diamond auto-shape, which is defined by
    /// its bounding box rather than by four stored vertices. This method is the
    /// one place that decides whether the scene's four points can be represented
    /// that way, so a polygon that cannot be is <strong>refused</strong> rather
    /// than approximated into a diamond that was never asked for.
    /// </para>
    /// <para>
    /// The test is deliberately strict: exactly four points, exactly one vertex
    /// on each edge midpoint of the box, the two diagonals equal and spanning
    /// the full width and height, and all comparisons within
    /// <see cref="GeometryMath.Epsilon"/>. A rotated square, a rectangle, or a
    /// four-point polygon with an off-centre vertex all fail.
    /// </para>
    /// </remarks>
    private static bool TryGetDiamondBounds(IReadOnlyList<PointD> points, out RectD bounds)
    {
        bounds = default;

        if (points.Count != 4)
        {
            return false;
        }

        var left = points.Min(point => point.X);
        var right = points.Max(point => point.X);
        var top = points.Min(point => point.Y);
        var bottom = points.Max(point => point.Y);
        var centreX = (left + right) / 2d;
        var centreY = (top + bottom) / 2d;

        // One vertex per side midpoint, and nothing else. This single rule
        // rejects a rotated square, a rectangle, and any off-centre vertex,
        // because none of those puts all four points on those four spots.
        var onLeft = 0;
        var onRight = 0;
        var onTop = 0;
        var onBottom = 0;
        foreach (PointD point in points)
        {
            var atLeft = Near(point.X, left) && Near(point.Y, centreY);
            var atRight = Near(point.X, right) && Near(point.Y, centreY);
            var atTop = Near(point.Y, top) && Near(point.X, centreX);
            var atBottom = Near(point.Y, bottom) && Near(point.X, centreX);

            if (atLeft)
            {
                onLeft++;
            }

            if (atRight)
            {
                onRight++;
            }

            if (atTop)
            {
                onTop++;
            }

            if (atBottom)
            {
                onBottom++;
            }
        }

        if (onLeft != 1 || onRight != 1 || onTop != 1 || onBottom != 1)
        {
            return false;
        }

        // A zero-extent box is a point, not a diamond.
        if (!GeometryMath.ApproximatelyEqual(right - left, bottom - top))
        {
            return false;
        }

        if (right - left <= GeometryMath.Epsilon)
        {
            return false;
        }

        bounds = new RectD(left, top, right - left, bottom - top);
        return true;

        static bool Near(double left, double right)
        {
            return GeometryMath.ApproximatelyEqual(left, right);
        }
    }

    /// <summary>
    /// Builds a milestone request as a diamond auto-shape at the polygon's
    /// bounding box.
    /// </summary>
    /// <remarks>
    /// R4.5 D1. The four points are the scene's <em>source</em>; the host object
    /// is a diamond auto-shape placed at their bounding box, which measurement
    /// showed is exact (0.000 EMU) where a freeform is quantised to +1 EMU. The
    /// points are still carried on the request so the translation is auditable,
    /// but the adapter places the shape from <c>Bounds</c>.
    /// </remarks>
    /// <param name="polygon">The scene polygon.</param>
    /// <param name="bounds">The diamond's bounding box.</param>
    /// <returns>The translated shape request.</returns>
    /// <remarks>
    /// R4.6: the milestone's fill and outline are mapped, not chosen. Section 21's
    /// subtype table gives planned, actual, baseline, and critical milestones
    /// distinct fill/outline pairs, and a diamond drawn in the host's default
    /// accent colour would erase that distinction entirely.
    /// </remarks>
    private OfficeShapeRequest BuildDiamond(ScenePolygon polygon, RectD bounds) =>
        new(
            polygon.PrimitiveId,
            OfficeShapeKind.Diamond,
            new OfficeShapeGeometry(
                Bounds: originDelta.Apply(bounds),
                Points: [.. polygon.Points.Select(originDelta.Apply)]),
            polygon.ZLayer,
            FillColour: polygon.Style.FillColour,
            StrokeColour: polygon.Style.StrokeColour,
            LineWidthPt: polygon.Style.OutlineWidthPt,
            HatchPattern: polygon.Style.HatchPattern);

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
    /// absent rather than being defaulted; R4.6 owns the full style matrix.
    /// </para>
    /// <para>
    /// A label's text <em>colour</em> is a separate fact from its shape fill and
    /// is deliberately not set here. <see cref="SceneStyle"/> carries no text-colour
    /// member, so a label inside a delay event cannot express the
    /// <c>DelayText</c> token through this path; that is a recorded gap for the
    /// scene model, not something the renderer may approximate by leaving the
    /// host's automatic colour in place.
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
            HatchPattern: text.Style.HatchPattern,
            FontFamily: text.Style.FontFamily,
            FontSizePt: text.Style.FontSizePt,
            Bold: text.Style.Bold,
            Text: text.Text,
            Alignment: text.Alignment);
}
