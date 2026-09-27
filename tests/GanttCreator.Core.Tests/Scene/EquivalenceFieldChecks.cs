using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// The field factories behind <see cref="EquivalenceFields"/>. Each one transcribes one
/// field of the entity guide's field-contract table and states the probe that decides
/// it, so the row definitions read as the guide reads and the geometry lives here.
/// </summary>
/// <remarks>
/// A probe is written to fail when the field's *substance* is gone, not merely when a
/// member is missing. That distinction is the R3.14 milestone lesson: a probe that only
/// counted four points passed a polygon whose vertices were rotated one step, because
/// the rotation preserves both the count and the tip-to-tip extent.
/// </remarks>
internal static class EquivalenceFieldChecks
{
    /// <summary>The shared contract citation every point-geometry field cites.</summary>
    internal const string Shared = "Entity guide 'Shared entity contract' (lines 41-51)";

    /// <summary>The z-order contract citation every layer and ordering field cites.</summary>
    internal const string ZOrder = "Entity guide 'Z-order contract' (lines 255-274)";

    /// <summary>The identity citation every ownership and style field cites.</summary>
    internal const string Identity =
        "Entity guide 'Shared entity contract' - stable EntityId, resolved StyleKey, z-layer with deterministic order (lines 43-48)";

    /// <summary>
    /// Gets the milestone tip-to-tip size the reference scene is built with, in points.
    /// </summary>
    /// <remarks>
    /// The <c>MilestoneSizePt</c> token, repeated here rather than read from a
    /// production default because §20 makes it a caller-supplied token, not a fixed one.
    /// </remarks>
    internal const double MilestoneSizePt = 8.0;

    /// <summary>
    /// Gets the critical-interval overlay height the reference scene is built with.
    /// </summary>
    /// <remarks>
    /// The <c>CriticalLinePt</c> token, repeated for the same reason as
    /// <see cref="MilestoneSizePt"/>: §11 makes it a caller-supplied token.
    /// </remarks>
    internal const double CriticalLinePt = 1.0;

    internal static EquivalenceField Kind(string member, string kind, string citation) =>
        new(
            $"the {member} primitive kind is {kind}",
            citation,
            member,
            (_, primitive) => KindOf(primitive) == kind);

    internal static EquivalenceField Extents(string member, string citation) =>
        new(
            $"the {member} geometry with non-zero extents",
            citation,
            member,
            (_, primitive) => BoundsOf(primitive) is { } bounds && HasPositiveExtents(bounds));

    internal static EquivalenceField TextExtents(string member, string citation) =>
        new(
            $"the {member} resolved text bounds",
            citation,
            member,
            (_, primitive) => primitive is SceneText text && HasPositiveExtents(text.TextBounds));

    internal static EquivalenceField TextContent(string member, string citation) =>
        new(
            $"the {member} text content",
            citation,
            member,
            (_, primitive) => primitive is SceneText text && !string.IsNullOrWhiteSpace(text.Text));

    internal static EquivalenceField AlignmentDefined(string member, string citation) =>
        new(
            $"the {member} resolved text alignment",
            citation,
            member,
            (_, primitive) => primitive is SceneText text && Enum.IsDefined(text.Alignment));

    internal static EquivalenceField CentredAlignment(string member, string citation) =>
        new(
            $"the {member} is centred in its band",
            citation,
            member,
            (_, primitive) => primitive is SceneText { Alignment: GanttTextAlignment.Centre });

    internal static EquivalenceField AtLayer(string member, ZLayer layer, string citation) =>
        new(
            $"the {member} z-layer",
            citation,
            member,
            (_, primitive) => primitive.ZLayer == layer);

    internal static EquivalenceField LineEndpoints(string member, string citation) =>
        new(
            $"the {member} start and end points",
            citation,
            member,
            (_, primitive) => primitive is SceneLine line && IsFinitePoint(line.From) && IsFinitePoint(line.To));

    /// <summary>
    /// Asserts the R3.15 contract from the catalogue side: the scene's declared chart
    /// bounds and the chart:background primitive are the same rectangle.
    /// </summary>
    internal static EquivalenceField EqualsSceneChartBounds(string member, string citation) =>
        new(
            "the scene ChartBounds equal the chart background",
            citation,
            member,
            (scene, primitive) =>
                primitive is SceneRect rect
                && GeometryMath.ApproximatelyEqual(scene.ChartBounds.X, rect.Bounds.X)
                && GeometryMath.ApproximatelyEqual(scene.ChartBounds.Y, rect.Bounds.Y)
                && GeometryMath.ApproximatelyEqual(scene.ChartBounds.Width, rect.Bounds.Width)
                && GeometryMath.ApproximatelyEqual(scene.ChartBounds.Height, rect.Bounds.Height));

    /// <summary>
    /// Asserts a delineator line spans the plot's full height, which is §24's defining
    /// property and what distinguishes it from a lane-bound entity.
    /// </summary>
    internal static EquivalenceField FullPlotHeight(string member, string citation) =>
        new(
            "the line spans the plot's full height",
            citation,
            member,
            (scene, primitive) =>
                primitive is SceneLine line
                && GeometryMath.ApproximatelyEqual(line.From.Y, scene.PlotBounds.Top)
                && GeometryMath.ApproximatelyEqual(line.To.Y, scene.PlotBounds.Bottom));

    /// <summary>
    /// Asserts a critical overlay is exactly <c>CriticalLinePt</c> tall, so a renderer
    /// reading its top edge draws a line of the resolved thickness rather than a band.
    /// </summary>
    internal static EquivalenceField CriticalLineHeight(string member, string citation) =>
        new(
            "the CriticalLinePt overlay height",
            citation,
            member,
            (_, primitive) =>
                primitive is SceneRect rect && GeometryMath.ApproximatelyEqual(rect.Bounds.Height, CriticalLinePt));

    /// <summary>
    /// Asserts a milestone diamond's vertices are in §20's draw order, not merely
    /// present. A polygon rotated one step keeps the same four points and the same
    /// tip-to-tip extent, so only the order assertion catches it.
    /// </summary>
    internal static EquivalenceField DiamondDrawOrder(string member, string citation) =>
        new(
            "the four points in draw order (top, right, bottom, left)",
            citation,
            member,
            (_, primitive) => HasDiamondDrawOrder(primitive));

    internal static EquivalenceField TipToTipExtent(string member, string citation) =>
        new(
            "the tip-to-tip extent, MilestoneSizePt on both axes",
            citation,
            member,
            (_, primitive) => HasTipToTipExtent(primitive, MilestoneSizePt));

    /// <summary>
    /// Asserts a header label names its parent rectangle by appending ":label" to the
    /// parent's identifier - R3.17's convention, and the exact text a renderer
    /// reconciles on. The parent rectangle must actually exist, so a label that merely
    /// ends in ":label" without a rectangle of that name does not pass.
    /// </summary>
    internal static EquivalenceField LabelNamesParentRectangle(string member, string citation) =>
        new(
            "the label names an existing parent rectangle by appending ':label'",
            citation,
            member,
            (scene, primitive) =>
            {
                const string suffix = ":label";
                if (!primitive.PrimitiveId.EndsWith(suffix, StringComparison.Ordinal))
                {
                    return false;
                }

                // A plain string, not a span: a span over a primitive's identifier
                // would be a ref local, which cannot be captured by the predicate.
                string parentId = primitive.PrimitiveId[..^suffix.Length];
                return scene.Primitives.Any(candidate =>
                    string.Equals(candidate.PrimitiveId, parentId, StringComparison.Ordinal));
            });

    internal static EquivalenceField FillResolved(string member, string citation) =>
        new(
            "the resolved fill colour",
            $"{citation} ({Identity})",
            member,
            (_, primitive) => StyleOf(primitive)?.FillColour is not null);

    internal static EquivalenceField StrokeResolved(string member, string citation) =>
        new(
            "the resolved stroke colour",
            $"{citation} ({Identity})",
            member,
            (_, primitive) =>
                StyleOf(primitive) is { } style
                && style.StrokeColour is not null
                && !string.IsNullOrWhiteSpace(style.StyleKey));

    internal static EquivalenceField OutlineWidthResolved(string member, string citation) =>
        new(
            "the resolved outline width",
            $"{citation} ({Identity})",
            member,
            (_, primitive) => StyleOf(primitive)?.OutlineWidthPt is { } width && width > 0);

    internal static EquivalenceField NoHatch(string member, string citation) =>
        new(
            $"no hatch pattern on the {member}",
            citation,
            member,
            (_, primitive) => StyleOf(primitive)?.HatchPattern is GanttHatchPattern.None);

    /// <summary>
    /// Asserts the hatch field is <em>readable</em> without asserting its value.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The reference style set resolves HatchPattern to None for every style, so no
    /// committed primitive carries a hatch and a value assertion would test a value
    /// nothing produces. What remains is that the field exists on the style.
    /// </para>
    /// <para>
    /// <b>This field is unfalsifiable by construction, and deliberately so.</b> Every
    /// styled primitive requires a non-null <see cref="SceneStyle"/> in its
    /// constructor, so no valid scene can make this probe return false. Rather than
    /// manufacture a field-stripped positive that would prove nothing, the field is
    /// kept as a documented existence claim and the real contract is proved where it
    /// is observable: <c>SceneSnapshotTests</c> round-trips the hatch value and
    /// <c>Deserialize_rejects_undefined_hatch_pattern_as_invalid_scene_data</c> refuses
    /// an undefined pattern. Recording the exemption is the honest outcome; a
    /// fabricated positive here would be the exact defect R3.14's review pass found in
    /// the milestone draw-order probe.
    /// </para>
    /// </remarks>
    internal static EquivalenceField HatchFieldReadable(string member, string citation) =>
        new(
            "the hatch field is readable on the style",
            citation,
            member,
            (_, primitive) => StyleOf(primitive) is not null);

    internal static EquivalenceField OrderKeys(string member, string citation) =>
        new(
            "the deterministic lane and stack order keys",
            citation,
            member,
            (_, primitive) => primitive.LaneOrder is not null && primitive.StackIndex is not null);

    internal static EquivalenceField NoOrderKeys(string member, string citation) =>
        new(
            $"the {member} carries no lane or stack order keys",
            citation,
            member,
            (_, primitive) => primitive.LaneOrder is null && primitive.StackIndex is null);

    internal static EquivalenceField RoleDerivedId(string member, string role, string citation) =>
        new(
            $"the role-derived ':{role}' identifier",
            citation,
            member,
            (_, primitive) =>
                string.Equals(primitive.PrimitiveId, ScenePrimitive.CreateId(primitive.OwnerId, role), StringComparison.Ordinal));

    internal static EquivalenceField RowOwned(string member, string citation) =>
        new(
            $"the {member} owning row identity",
            citation,
            member,
            (_, primitive) => primitive.OwnerId.Kind == SceneOwnerKind.Row && primitive.OwnerId.OwnedRows.Count == 1);

    internal static EquivalenceField ChartOwned(string member, string citation) =>
        new(
            $"the {member} chart owner, not a row",
            citation,
            member,
            (_, primitive) => primitive.OwnerId.Kind == SceneOwnerKind.Chart);

    /// <summary>
    /// Accepts either a single-row owner or a shared <c>Rows</c> owner with two or more
    /// members. Deliberately distinct from <see cref="RowOwned"/>: a shared delineator
    /// line is legitimately not row-owned, and asserting that it was would contradict
    /// ADR-0017.
    /// </summary>
    internal static EquivalenceField SharedOrRowOwner(string member, string citation) =>
        new(
            "the owning row or shared-rows identity",
            citation,
            member,
            (_, primitive) =>
                primitive.OwnerId.Kind switch
                {
                    SceneOwnerKind.Row => primitive.OwnerId.OwnedRows.Count == 1,
                    SceneOwnerKind.Rows => primitive.OwnerId.OwnedRows.Count >= 2,
                    _ => false,
                });

    /// <summary>Gets the primitive's kind as the guide's table spells it.</summary>
    private static string KindOf(ScenePrimitive primitive) =>
        primitive switch
        {
            SceneRect => "rect",
            SceneLine => "line",
            ScenePolygon => "polygon",
            SceneText => "text",
            _ => primitive.GetType().Name,
        };

    /// <summary>
    /// Gets the primitive's own rectangle - a rect's bounds or a text's resolved text
    /// bounds - or <see langword="null"/> for a kind that has none.
    /// </summary>
    private static RectD? BoundsOf(ScenePrimitive primitive) =>
        primitive switch
        {
            SceneRect rect => rect.Bounds,
            SceneText text => text.TextBounds,
            _ => null,
        };

    /// <summary>
    /// Gets the resolved style of a styled primitive, or <see langword="null"/> for a
    /// primitive that carries none.
    /// </summary>
    /// <param name="primitive">The primitive under test.</param>
    /// <returns>The primitive's resolved style, or <see langword="null"/>.</returns>
    /// <remarks>
    /// <see cref="ScenePrimitive"/> deliberately does not expose a style: a
    /// <see cref="SceneGroup"/> has none, so the property lives on each styled
    /// primitive instead of being an uninitialised base member.
    /// </remarks>
    private static SceneStyle? StyleOf(ScenePrimitive primitive) =>
        primitive switch
        {
            SceneRect rect => rect.Style,
            SceneLine line => line.Style,
            ScenePolygon polygon => polygon.Style,
            SceneText text => text.Style,
            _ => null,
        };

    private static bool HasPositiveExtents(RectD bounds) =>
        double.IsFinite(bounds.X)
        && double.IsFinite(bounds.Y)
        && bounds.Width > 0
        && bounds.Height > 0;

    private static bool IsFinitePoint(PointD point) => double.IsFinite(point.X) && double.IsFinite(point.Y);

    private static bool HasDiamondDrawOrder(ScenePrimitive primitive)
    {
        if (primitive is not ScenePolygon { Points.Count: 4 } polygon)
        {
            return false;
        }

        // §20's draw order: the top vertex, then the right, the bottom, and the left.
        // The scene's Y axis grows downwards, so top is the minimum Y and the right
        // vertex is the maximum X - matching the order MilestoneMarkerBuilder emits.
        double minX = polygon.Points.Min(point => point.X);
        double maxX = polygon.Points.Max(point => point.X);
        double minY = polygon.Points.Min(point => point.Y);
        double maxY = polygon.Points.Max(point => point.Y);

        return GeometryMath.ApproximatelyEqual(polygon.Points[0].Y, minY)
            && GeometryMath.ApproximatelyEqual(polygon.Points[1].X, maxX)
            && GeometryMath.ApproximatelyEqual(polygon.Points[2].Y, maxY)
            && GeometryMath.ApproximatelyEqual(polygon.Points[3].X, minX);
    }

    private static bool HasTipToTipExtent(ScenePrimitive primitive, double size) =>
        primitive is ScenePolygon polygon
        && polygon.Points.Count > 0
        && GeometryMath.ApproximatelyEqual(polygon.Points.Max(point => point.X) - polygon.Points.Min(point => point.X), size)
        && GeometryMath.ApproximatelyEqual(polygon.Points.Max(point => point.Y) - polygon.Points.Min(point => point.Y), size);
}
