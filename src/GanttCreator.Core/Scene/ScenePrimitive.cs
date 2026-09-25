namespace GanttCreator.Core.Scene;

/// <summary>Base type for an immutable, ordered scene primitive.</summary>
public abstract record ScenePrimitive
{
    /// <summary>Initialises primitive identity and ordering metadata.</summary>
    /// <param name="primitiveId">The stable role-derived primitive identifier.</param>
    /// <param name="ownerId">The stable source row identifier.</param>
    /// <param name="zLayer">The scene layer.</param>
    /// <param name="entityType">The source entity type, when known.</param>
    /// <param name="laneOrder">The lane ordering value, when known.</param>
    /// <param name="stackIndex">The stack ordering value, when known.</param>
    /// <param name="sortOrder">The explicit user ordering value, when known.</param>
    protected ScenePrimitive(
        string primitiveId,
        SceneOwnerId ownerId,
        ZLayer zLayer,
        GanttEntityType? entityType = null,
        int? laneOrder = null,
        int? stackIndex = null,
        int? sortOrder = null
    )
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(primitiveId);
        ArgumentNullException.ThrowIfNull(ownerId);
        if (!Enum.IsDefined(zLayer))
        {
            throw new ArgumentOutOfRangeException(nameof(zLayer));
        }

        if (laneOrder is { } lane && lane < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(laneOrder));
        }

        if (stackIndex is { } stack && stack < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(stackIndex));
        }

        if (sortOrder is { } sort && sort < 0)
        {
            // SortOrder is validated as a non-negative integer in the row validator
            // (GanttValidationCodes.BadSortOrder). A negative value reaching a
            // primitive would silently reorder the scene, so it is refused here too.
            throw new ArgumentOutOfRangeException(nameof(sortOrder));
        }

        PrimitiveId = primitiveId;
        OwnerId = ownerId;
        ZLayer = zLayer;
        EntityType = entityType;
        LaneOrder = laneOrder;
        StackIndex = stackIndex;
        SortOrder = sortOrder;
    }

    /// <summary>Gets the stable role-derived primitive identifier.</summary>
    public string PrimitiveId { get; }

    /// <summary>Gets the stable scene owner identifier.</summary>
    public SceneOwnerId OwnerId { get; }

    /// <summary>Gets the scene layer.</summary>
    public ZLayer ZLayer { get; }

    /// <summary>Gets the source entity type, when known.</summary>
    public GanttEntityType? EntityType { get; }

    /// <summary>Gets the lane ordering value, when known.</summary>
    public int? LaneOrder { get; }

    /// <summary>Gets the stack ordering value, when known.</summary>
    public int? StackIndex { get; }

    /// <summary>Gets the explicit user ordering value, when known.</summary>
    public int? SortOrder { get; }

    /// <summary>Creates a stable role-derived primitive identifier.</summary>
    /// <param name="ownerId">The stable source row identifier.</param>
    /// <param name="role">The role suffix for this primitive.</param>
    /// <returns>The stable primitive identifier.</returns>
    public static string CreateId(SceneOwnerId ownerId, string role)
    {
        ArgumentNullException.ThrowIfNull(ownerId);
        ArgumentException.ThrowIfNullOrWhiteSpace(role);
        return $"{ownerId.Value}:{role}";
    }
}
