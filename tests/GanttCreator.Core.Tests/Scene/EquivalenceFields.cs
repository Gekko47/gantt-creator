using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// One required field of one entity-guide equivalence-table row, as transcribed at
/// R3.14 Step 0 against <c>docs/07-GANTT-ENTITY-GUIDE.md</c> revision 5.
/// </summary>
/// <remarks>
/// <para>
/// The field list is the contract under test, so it is transcribed exactly and each
/// entry carries the guide section it comes from. A field that the scene does not
/// carry with a non-default, correctly typed value is a defect in the model, never a
/// renderer-side workaround (R3.14 D1).
/// </para>
/// <para>
/// <b>Why the list is assembled rather than lifted from one table.</b> The guide's
/// "Entity-to-renderer equivalence" section states which representation each entity
/// takes in each of the four renderers; the per-entity <em>fields</em> those
/// representations need come from the same guide's "Shared entity contract" bullet
/// list plus the per-entity section for the row. Amending the guide to add a field
/// table would be an entity-contract change, so it is not done here: the transcription
/// cites both and <c>EquivalenceThinSliceTests</c> guards that every field keeps a
/// citation.
/// </para>
/// <para>
/// <b>Ownership.</b> "Shared entity contract" requires ownership metadata beginning
/// <c>GanttCreator.</c> <em>in Excel shapes</em>. The scene's ownership metadata is
/// <see cref="ScenePrimitive.OwnerId"/> plus the role-derived
/// <see cref="ScenePrimitive.PrimitiveId"/>, which is what a renderer composes that
/// tag from; the literal prefixed tag is a Phase-4 port member owned by R4.1 D2 /
/// R4.3. Asserting a <c>GanttCreator.</c> string in Core would be a model extension,
/// which this row forbids.
/// </para>
/// </remarks>
/// <param name="Name">The stable field name reported when the scene fails the field.</param>
/// <param name="GuideCitation">The entity-guide section this field is transcribed from.</param>
/// <param name="IsSatisfied">
/// Whether the scene carries this field with a non-default, correctly typed value.
/// The whole scene is passed because one field - an external label's placement - is
/// only decidable against the entity it belongs to.
/// </param>
internal sealed record EquivalenceField(
    string Name,
    string GuideCitation,
    Func<GanttScene, ScenePrimitive, bool> IsSatisfied);

/// <summary>
/// One entity-guide equivalence-table row, with every field it requires of all four
/// renderers.
/// </summary>
/// <param name="Name">The row's name, used in assertion messages.</param>
/// <param name="GuideCitation">The equivalence-table row this transcription covers.</param>
/// <param name="Fields">The required fields, in transcription order.</param>
internal sealed record EquivalenceRow(
    string Name,
    string GuideCitation,
    IReadOnlyList<EquivalenceField> Fields)
{
    /// <summary>
    /// Collects the names of the fields this row does not find satisfied in the scene.
    /// </summary>
    /// <param name="scene">The scene under test.</param>
    /// <param name="primitiveId">The stable identifier of the row's primitive.</param>
    /// <returns>
    /// The unsatisfied field names, or a single entry naming the absent primitive.
    /// An empty list means the scene carries every field of the row.
    /// </returns>
    /// <remarks>
    /// This collects rather than asserts so a test can prove which field a
    /// deliberately field-stripped variant loses, which is the row's positive test.
    /// </remarks>
    public IReadOnlyList<string> UnsatisfiedFields(GanttScene scene, string primitiveId)
    {
        ArgumentNullException.ThrowIfNull(scene);

        ScenePrimitive? primitive = scene.Primitives.FirstOrDefault(candidate =>
            string.Equals(candidate.PrimitiveId, primitiveId, StringComparison.Ordinal));
        if (primitive is null)
        {
            return [$"the primitive '{primitiveId}' is absent from the scene"];
        }

        return
        [
            .. Fields.Where(field => !field.IsSatisfied(scene, primitive)).Select(field => field.Name),
        ];
    }
}

/// <summary>
/// The R3.14 thin slice: the three equivalence-table rows the row names, each
/// transcribed field by field from the entity guide.
/// </summary>
internal static class EquivalenceFields
{
    private const string _shared = "Entity guide 'Shared entity contract' (lines 41-51)";
    private const string _zOrder = "Entity guide 'Z-order contract' (lines 255-274)";
    private const string _identity = "Entity guide 'Shared entity contract' - stable EntityId, resolved StyleKey, z-layer with deterministic order (lines 43-48)";

    /// <summary>
    /// The span-bar row. The equivalence table's "Activity/delay" row requires a
    /// rectangle shape in the live worksheet, a rectangle shape in the editable
    /// composition and PowerPoint, and a raster rectangle in the PNG; all three are
    /// the same <see cref="SceneRect"/>, so the fields below are what each renderer
    /// reads to draw it.
    /// </summary>
    public static EquivalenceRow SpanBar { get; } =
        new(
            "span bar",
            "Entity guide 'Entity-to-renderer equivalence', row 'Activity/delay' (line 697); geometry and style from §12 'General span activity' (lines 448-462) and §13 'As-planned activity' (lines 464-474)",
            [
                new("the primitive kind is a rectangle", $"{_shared} - point-based geometry", static (_, primitive) => primitive is SceneRect),
                new(
                    "resolved bar geometry with non-zero extents",
                    "Entity guide §12 'Geometry' (line 454) and 'Shared coordinate rules' (lines 243-253)",
                    static (_, primitive) => primitive is SceneRect rect && HasPositiveExtents(rect.Bounds)),
                FillResolved("§12 'Style' (line 456): fill and outline are explicit; no renderer defaults"),
                StrokeResolved("§12 'Style' (line 456): fill and outline are explicit; no renderer defaults"),
                OutlineWidthResolved("§12 'Style' (line 456): the standard outline is resolved before rendering"),
                NoHatch("Entity guide 'Type catalogue' - an activity is Fill + outline, never Hatch + outline (line 88)"),
                new("the activity-body z-layer", $"{_zOrder} - layer 40", static (_, primitive) => primitive.ZLayer == ZLayer.ActivityBody),
                OrderKeys("Entity guide 'Z-order contract' (line 272): after subtype priority, order by lane, stack, SortOrder, and stable ID"),
                RoleDerivedId("bar", "Entity guide §12 with the R3.6 role-derived ':bar' identifier"),
                RowOwnership("Entity guide 'Shared entity contract' (line 43): a stable EntityId unrelated to worksheet row number"),
            ]);

    /// <summary>
    /// The external description-label row. The equivalence table's "Labels" row
    /// requires a text box shape in the live worksheet, a text box shape in the
    /// editable composition and PowerPoint, and raster text at the scene bounds in
    /// the PNG, so the same resolved bounds serve all three.
    /// </summary>
    public static EquivalenceRow ExternalDescriptionLabel { get; } =
        new(
            "external description label",
            "Entity guide 'Entity-to-renderer equivalence', row 'Labels' (line 701); fields from §22 'Activity or milestone description label' (lines 592-618)",
            [
                new("the primitive kind is text", $"{_shared} - point-based geometry", static (_, primitive) => primitive is SceneText),
                new(
                    "the label text content",
                    "Entity guide §22 - the description is user-facing text (line 596)",
                    static (_, primitive) => primitive is SceneText text && !string.IsNullOrWhiteSpace(text.Text)),
                new(
                    "the resolved text bounds",
                    "Entity guide §22 'Text measurement' (line 614): the label bounds are resolved during scene construction and renderers consume them without reselecting the side",
                    static (_, primitive) => primitive is SceneText text && HasPositiveExtents(text.TextBounds)),
                new(
                    "the external placement, resolved into the text bounds",
                    "Entity guide §22 'Candidate geometry' (lines 600-608): an external candidate stands off the shape by LabelGapPt, and the cascade is resolved in the scene",
                    static (scene, primitive) => IsOutsideParentBar(scene, primitive)),
                TextColourResolved("Entity guide §13 (line 470): a label uses DefaultText and Auto unless explicitly set"),
                TextAlignmentResolved("Entity guide 'Shared colour and typography tokens' - text alignment is a resolved style value, not a label position (ADR-0018)"),
                new("the label z-layer", $"{_zOrder} - layer 70", static (_, primitive) => primitive.ZLayer == ZLayer.Label),
                OrderKeys("Entity guide 'Z-order contract' (line 272): after subtype priority, order by lane, stack, SortOrder, and stable ID"),
                RoleDerivedId("label", "Entity guide §22 (line 626): label IDs are derived from the parent event ID and label role"),
                SharedParentRowOwnership("Entity guide §22 - a description label is owned by the event it describes"),
            ]);

    /// <summary>
    /// The milestone row. The equivalence table's "Milestone" row requires a
    /// four-point freeform polygon in the live worksheet, a freeform polygon in the
    /// editable composition and PowerPoint, and a raster polygon in the PNG.
    /// </summary>
    public static EquivalenceRow MilestoneDiamond { get; } =
        new(
            "milestone diamond",
            "Entity guide 'Entity-to-renderer equivalence', row 'Milestone' (line 700); geometry from §20 'General milestone diamond' (lines 554-575) and style from §21 (lines 577-590)",
            [
                new("the primitive kind is a polygon", $"{_shared} - point-based geometry", static (_, primitive) => primitive is ScenePolygon),
                new(
                    "the four points in draw order (top, right, bottom, left)",
                    "Entity guide §20 'Geometry' (line 560): build a four-point polygon - not a rotated square",
                    static (_, primitive) => HasDiamondDrawOrder(primitive)),
                new(
                    "the tip-to-tip extent, MilestoneSizePt on both axes",
                    "Entity guide §20 'Geometry' (line 569): size = MilestoneSizePt makes the tip-to-tip bounds exact and consistent across Excel, PowerPoint, and PNG",
                    static (_, primitive) => HasTipToTipExtent(primitive, MilestoneSizePt)),
                FillResolved("Entity guide §21 (line 583): a planned milestone fills with PlannedFill"),
                StrokeResolved("Entity guide §21 (line 583): a planned milestone outlines with PlannedOutline"),
                OutlineWidthResolved("Entity guide §21 (line 583): the resolved outline is drawn as resolved"),
                NoHatch("Entity guide 'Type catalogue' - a milestone is Fill + outline, never Hatch + outline (lines 97-100)"),
                new("the milestone z-layer", $"{_zOrder} - layer 60", static (_, primitive) => primitive.ZLayer == ZLayer.Milestone),
                OrderKeys("Entity guide 'Z-order contract' (line 272): after subtype priority, order by lane, stack, SortOrder, and stable ID"),
                RoleDerivedId("marker", "Entity guide §20 with the R3.7 role-derived ':marker' identifier"),
                RowOwnership("Entity guide 'Shared entity contract' (line 43): a stable EntityId unrelated to worksheet row number"),
            ]);

    /// <summary>Every row of the thin slice, in the order the work item names them.</summary>
    public static IReadOnlyList<EquivalenceRow> All { get; } = [SpanBar, ExternalDescriptionLabel, MilestoneDiamond];

    /// <summary>
    /// Gets the milestone tip-to-tip size the thin slice is built with, in points.
    /// </summary>
    /// <remarks>
    /// This is the <c>MilestoneSizePt</c> the reference scene is built with; the
    /// value is repeated here rather than read from a production default because
    /// §20 makes it a caller-supplied token, not a fixed one.
    /// </remarks>
    public const double MilestoneSizePt = 8.0;

    private static EquivalenceField FillResolved(string citation) =>
        new(
            "the resolved fill colour",
            $"{citation} ({_identity})",
            static (_, primitive) => StyleOf(primitive)?.FillColour is not null);

    private static EquivalenceField StrokeResolved(string citation) =>
        new(
            "the resolved stroke colour",
            $"{citation} ({_identity})",
            static (_, primitive) =>
                StyleOf(primitive) is { } style &&
                style.StrokeColour is not null &&
                !string.IsNullOrWhiteSpace(style.StyleKey));

    private static EquivalenceField OutlineWidthResolved(string citation) =>
        new(
            "the resolved outline width",
            $"{citation} ({_identity})",
            static (_, primitive) => StyleOf(primitive)?.OutlineWidthPt is { } width && width > 0);

    private static EquivalenceField NoHatch(string citation) =>
        new(
            "no hatch pattern on this row",
            citation,
            static (_, primitive) => StyleOf(primitive)?.HatchPattern is GanttHatchPattern.None);

    private static EquivalenceField TextColourResolved(string citation) =>
        new(
            "the resolved text colour and style key",
            citation,
            static (_, primitive) =>
                StyleOf(primitive) is { } style &&
                style.FillColour is not null &&
                !string.IsNullOrWhiteSpace(style.StyleKey));

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

    private static EquivalenceField TextAlignmentResolved(string citation) =>
        new(
            "the resolved text alignment",
            citation,
            static (_, primitive) => primitive is SceneText text && Enum.IsDefined(text.Alignment));

    private static EquivalenceField OrderKeys(string citation) =>
        new(
            "the deterministic lane and stack order keys",
            citation,
            static (_, primitive) => primitive.LaneOrder is not null && primitive.StackIndex is not null);

    private static EquivalenceField RoleDerivedId(string role, string citation) =>
        new(
            $"the role-derived ':{role}' identifier",
            citation,
            (_, primitive) => string.Equals(
                primitive.PrimitiveId,
                ScenePrimitive.CreateId(primitive.OwnerId, role),
                StringComparison.Ordinal));

    private static EquivalenceField RowOwnership(string citation) =>
        new(
            "the owning row identity",
            citation,
            static (_, primitive) =>
                primitive.OwnerId.Kind == SceneOwnerKind.Row && primitive.OwnerId.OwnedRows.Count == 1);

    private static EquivalenceField SharedParentRowOwnership(string citation) =>
        new(
            "the owning row identity, shared with the parent event",
            citation,
            static (scene, primitive) =>
                primitive.OwnerId.Kind == SceneOwnerKind.Row &&
                primitive.OwnerId.OwnedRows.Count == 1 &&
                FindSibling<SceneRect>(scene, primitive.OwnerId, "bar") is not null);

    private static bool HasPositiveExtents(RectD bounds) =>
        double.IsFinite(bounds.X) &&
        double.IsFinite(bounds.Y) &&
        bounds.Width > 0 &&
        bounds.Height > 0;

    /// <summary>
    /// Determines whether a primitive is a four-point polygon whose vertices are in
    /// the §20 draw order: top, right, bottom, left.
    /// </summary>
    /// <param name="primitive">The primitive under test.</param>
    /// <returns>
    /// <see langword="true"/> when the polygon has four vertices and the first is the
    /// topmost, the second the rightmost, the third the bottommost, and the fourth the
    /// leftmost.
    /// </returns>
    /// <remarks>
    /// <para>
    /// The scene's coordinate convention is screen-like: Y grows downwards, so the top
    /// vertex is the one with the minimum Y and the right vertex the maximum X, exactly
    /// as <c>MilestoneMarkerBuilder</c> emits them. A vertex count alone is not the
    /// field §20 requires: a polygon with four vertices in any other order would draw a
    /// different shape in a freeform renderer, which is why the order is asserted here
    /// and not inferred from the count.
    /// </para>
    /// <para>
    /// A polygon rotated one step (right, bottom, left, top) keeps the same four points
    /// and the same tip-to-tip extent, so <c>HasTipToTipExtent</c> cannot catch it; only
    /// this can.
    /// </para>
    /// </remarks>
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
        primitive is ScenePolygon polygon &&
        polygon.Points.Count > 0 &&
        GeometryMath.ApproximatelyEqual(polygon.Points.Max(point => point.X) - polygon.Points.Min(point => point.X), size) &&
        GeometryMath.ApproximatelyEqual(polygon.Points.Max(point => point.Y) - polygon.Points.Min(point => point.Y), size);

    /// <summary>
    /// Finds the sibling primitive of the requested kind that the same row owns under
    /// the given role, or <see langword="null"/> when the row owns no such primitive.
    /// </summary>
    /// <typeparam name="TPrimitive">The primitive kind to find.</typeparam>
    /// <param name="scene">The scene under test.</param>
    /// <param name="ownerId">The owning row.</param>
    /// <param name="role">The role suffix of the sibling's identifier.</param>
    /// <returns>The sibling primitive, or <see langword="null"/>.</returns>
    /// <remarks>
    /// A field that is only decidable against a sibling - an external label's
    /// placement against its bar - resolves it here from the scene the probe was
    /// handed. Nothing is cached in static state, so the catalogue stays safe for
    /// xUnit's parallel test collections.
    /// </remarks>
    private static TPrimitive? FindSibling<TPrimitive>(GanttScene scene, SceneOwnerId ownerId, string role)
        where TPrimitive : ScenePrimitive
    {
        string siblingId = ScenePrimitive.CreateId(ownerId, role);
        return scene.Primitives
            .OfType<TPrimitive>()
            .FirstOrDefault(candidate => string.Equals(candidate.PrimitiveId, siblingId, StringComparison.Ordinal));
    }

    private static bool IsOutsideParentBar(GanttScene scene, ScenePrimitive primitive)
    {
        SceneRect? bar = FindSibling<SceneRect>(scene, primitive.OwnerId, "bar");
        return bar is not null
            && primitive is SceneText text
            && !text.TextBounds.IntersectsWith(bar.Bounds);
    }
}
