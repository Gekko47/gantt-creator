namespace GanttCreator.Core.Scene;

/// <summary>An immutable line primitive.</summary>
public sealed record SceneLine : ScenePrimitive
{
    /// <summary>Initialises a line primitive.</summary>
    /// <param name="primitiveId">The stable role-derived primitive identifier.</param>
    /// <param name="ownerId">The stable source row identifier.</param>
    /// <param name="zLayer">The scene layer.</param>
    /// <param name="from">The line start point.</param>
    /// <param name="to">The line end point.</param>
    /// <param name="style">The resolved style.</param>
    /// <param name="entityType">The source entity type, when known.</param>
    /// <param name="laneOrder">The lane ordering value, when known.</param>
    /// <param name="stackIndex">The stack ordering value, when known.</param>
    /// <param name="sortOrder">The explicit user ordering value, when known.</param>
    public SceneLine(
        string primitiveId,
        GanttRowId ownerId,
        ZLayer zLayer,
        PointD from,
        PointD to,
        SceneStyle style,
        GanttEntityType? entityType = null,
        int? laneOrder = null,
        int? stackIndex = null,
        int? sortOrder = null)
        : base(primitiveId, ownerId, zLayer, entityType, laneOrder, stackIndex, sortOrder)
    {
        ArgumentNullException.ThrowIfNull(style);
        From = from;
        To = to;
        Style = style;
    }

    /// <summary>Gets the line start point.</summary>
    public PointD From { get; }

    /// <summary>Gets the line end point.</summary>
    public PointD To { get; }

    /// <summary>Gets the resolved style.</summary>
    public SceneStyle Style { get; }
}
