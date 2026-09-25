namespace GanttCreator.Core.Scene;

/// <summary>An immutable group primitive describing an ordered set of child primitive IDs.</summary>
public sealed record SceneGroup : ScenePrimitive
{
    /// <summary>Initialises a group with an immutable child list.</summary>
    /// <param name="primitiveId">The stable role-derived group identifier.</param>
    /// <param name="ownerId">The stable source row identifier.</param>
    /// <param name="zLayer">The scene layer.</param>
    /// <param name="childPrimitiveIds">The ordered child primitive identifiers.</param>
    /// <param name="entityType">The source entity type, when known.</param>
    /// <param name="laneOrder">The lane ordering value, when known.</param>
    /// <param name="stackIndex">The stack ordering value, when known.</param>
    /// <param name="sortOrder">The explicit user ordering value, when known.</param>
    public SceneGroup(
        string primitiveId,
        SceneOwnerId ownerId,
        ZLayer zLayer,
        IReadOnlyList<string> childPrimitiveIds,
        GanttEntityType? entityType = null,
        int? laneOrder = null,
        int? stackIndex = null,
        int? sortOrder = null
    )
        : base(primitiveId, ownerId, zLayer, entityType, laneOrder, stackIndex, sortOrder)
    {
        ArgumentNullException.ThrowIfNull(childPrimitiveIds);
        ChildPrimitiveIds = [.. childPrimitiveIds];
        if (ChildPrimitiveIds.Any(string.IsNullOrWhiteSpace))
        {
            throw new ArgumentException("Group child primitive IDs must not be blank.", nameof(childPrimitiveIds));
        }

        // A repeated child would render the same primitive twice and make the
        // snapshot non-canonical, so duplicates are refused rather than collapsed.
        if (ChildPrimitiveIds.Count != ChildPrimitiveIds.Distinct(StringComparer.Ordinal).Count())
        {
            throw new ArgumentException("Group child primitive IDs must be unique.", nameof(childPrimitiveIds));
        }

        // A group that lists itself, directly or through a chain, is a cycle and
        // would make the scene unserialisable. Nested groups are not resolved here
        // (a child ID is a string, not a reference), so only the direct self-reference
        // is detectable at this boundary; the snapshot layer rejects wider cycles.
        if (ChildPrimitiveIds.Contains(primitiveId, StringComparer.Ordinal))
        {
            throw new ArgumentException("A group cannot contain itself as a child.", nameof(childPrimitiveIds));
        }
    }

    /// <summary>Gets the ordered child primitive identifiers.</summary>
    public IReadOnlyList<string> ChildPrimitiveIds { get; }
}
