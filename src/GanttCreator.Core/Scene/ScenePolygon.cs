namespace GanttCreator.Core.Scene;

/// <summary>An immutable polygon primitive whose points are kept in draw order.</summary>
public sealed record ScenePolygon : ScenePrimitive
{
    /// <summary>Initialises a polygon with an immutable point list.</summary>
    /// <param name="primitiveId">The stable role-derived primitive identifier.</param>
    /// <param name="ownerId">The stable source row identifier.</param>
    /// <param name="zLayer">The scene layer.</param>
    /// <param name="points">The ordered polygon points.</param>
    /// <param name="style">The resolved style.</param>
    /// <param name="entityType">The source entity type, when known.</param>
    /// <param name="laneOrder">The lane ordering value, when known.</param>
    /// <param name="stackIndex">The stack ordering value, when known.</param>
    /// <param name="sortOrder">The explicit user ordering value, when known.</param>
    public ScenePolygon(
        string primitiveId,
        GanttRowId ownerId,
        ZLayer zLayer,
        IReadOnlyList<PointD> points,
        SceneStyle style,
        GanttEntityType? entityType = null,
        int? laneOrder = null,
        int? stackIndex = null,
        int? sortOrder = null)
        : base(primitiveId, ownerId, zLayer, entityType, laneOrder, stackIndex, sortOrder)
    {
        ArgumentNullException.ThrowIfNull(points);
        ArgumentNullException.ThrowIfNull(style);
        Points = [.. points];
        Style = style;
        if (Points.Count < 3)
        {
            throw new ArgumentOutOfRangeException(nameof(points), "A polygon requires at least three points.");
        }

        if (Points.Any(point => !double.IsFinite(point.X) || !double.IsFinite(point.Y)))
        {
            throw new ArgumentOutOfRangeException(nameof(points), "Polygon points must be finite.");
        }
    }

    /// <summary>Gets the ordered polygon points.</summary>
    public IReadOnlyList<PointD> Points { get; }

    /// <summary>Gets the resolved style.</summary>
    public SceneStyle Style { get; }
}
