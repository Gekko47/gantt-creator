namespace GanttCreator.Core.Scene;

/// <summary>An immutable rectangle primitive.</summary>
public sealed record SceneRect : ScenePrimitive
{
    /// <summary>Initialises a rectangle primitive.</summary>
    /// <param name="primitiveId">The stable role-derived primitive identifier.</param>
    /// <param name="ownerId">The stable source row identifier.</param>
    /// <param name="zLayer">The scene layer.</param>
    /// <param name="bounds">The rectangle bounds in points.</param>
    /// <param name="style">The resolved style.</param>
    /// <param name="entityType">The source entity type, when known.</param>
    /// <param name="laneOrder">The lane ordering value, when known.</param>
    /// <param name="stackIndex">The stack ordering value, when known.</param>
    /// <param name="sortOrder">The explicit user ordering value, when known.</param>
    public SceneRect(
        string primitiveId,
        SceneOwnerId ownerId,
        ZLayer zLayer,
        RectD bounds,
        SceneStyle style,
        GanttEntityType? entityType = null,
        int? laneOrder = null,
        int? stackIndex = null,
        int? sortOrder = null
    )
        : base(primitiveId, ownerId, zLayer, entityType, laneOrder, stackIndex, sortOrder)
    {
        ArgumentNullException.ThrowIfNull(style);
        Bounds = bounds;
        Style = style;
    }

    /// <summary>Gets the rectangle bounds in points.</summary>
    public RectD Bounds { get; }

    /// <summary>Gets the resolved style.</summary>
    public SceneStyle Style { get; }
}
