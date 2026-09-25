namespace GanttCreator.Core.Scene;

/// <summary>The typed result of attempting to create a validated scene.</summary>
/// <param name="Scene">The created scene when successful.</param>
/// <param name="Refusal">The typed refusal when unsuccessful.</param>
public sealed record SceneCreationOutcome(
    GanttScene? Scene,
    SceneCreationRefusal? Refusal)
{
    /// <summary>Gets a value indicating whether scene creation succeeded.</summary>
    public bool Succeeded => Scene is not null;
}

/// <summary>An immutable, deterministically ordered collection of scene primitives and warnings.</summary>
public sealed record GanttScene
{
    private GanttScene(
        RectD chartBounds,
        RectD plotBounds,
        IReadOnlyList<ScenePrimitive> primitives,
        IReadOnlyList<SceneWarning> warnings)
    {
        ChartBounds = chartBounds;
        PlotBounds = plotBounds;
        Primitives = primitives;
        Warnings = warnings;
    }

    /// <summary>Gets the complete chart bounds in points.</summary>
    public RectD ChartBounds { get; }

    /// <summary>Gets the plot bounds in points.</summary>
    public RectD PlotBounds { get; }

    /// <summary>Gets the immutable, ordered scene primitives.</summary>
    public IReadOnlyList<ScenePrimitive> Primitives { get; }

    /// <summary>Gets the immutable scene warnings.</summary>
    public IReadOnlyList<SceneWarning> Warnings { get; }

    /// <summary>Attempts to create a validated, deterministically ordered scene.</summary>
    /// <param name="chartBounds">The complete chart bounds.</param>
    /// <param name="plotBounds">The plot bounds.</param>
    /// <param name="primitives">The source primitives to validate and order.</param>
    /// <param name="warnings">The scene warnings.</param>
    /// <returns>A typed scene creation outcome.</returns>
    public static SceneCreationOutcome TryCreate(
        RectD chartBounds,
        RectD plotBounds,
        IEnumerable<ScenePrimitive> primitives,
        IEnumerable<SceneWarning> warnings)
    {
        if (!IsFinite(chartBounds) || !IsFinite(plotBounds))
        {
            return new SceneCreationOutcome(null, SceneCreationRefusal.InvalidBounds);
        }

        ScenePrimitive[] primitiveList = [.. primitives ?? []];
        if (primitiveList.Any(primitive => !IsValid(primitive)))
        {
            return new SceneCreationOutcome(null, SceneCreationRefusal.InvalidPrimitive);
        }

        IGrouping<string, ScenePrimitive>? duplicate = primitiveList
            .GroupBy(primitive => primitive.PrimitiveId, StringComparer.Ordinal)
            .FirstOrDefault(group => group.Count() > 1);
        if (duplicate is not null)
        {
            return new SceneCreationOutcome(null, SceneCreationRefusal.DuplicatePrimitiveId);
        }

        var ids = primitiveList.Select(primitive => primitive.PrimitiveId).ToHashSet(StringComparer.Ordinal);
        if (primitiveList.OfType<SceneGroup>().Any(group => group.ChildPrimitiveIds.Any(id => !ids.Contains(id))))
        {
            return new SceneCreationOutcome(null, SceneCreationRefusal.UnresolvedGroupChild);
        }

        // Unresolved children are checked first, so every child ID names a real
        // primitive here and a cycle cannot be masked as a missing reference.
        if (HasGroupCycle(primitiveList))
        {
            return new SceneCreationOutcome(null, SceneCreationRefusal.CyclicGroupChild);
        }

        ScenePrimitive[] ordered = [.. Order(primitiveList)];
        SceneWarning[] warningList = [.. warnings ?? []];
        return new SceneCreationOutcome(
            new GanttScene(chartBounds, plotBounds, ordered, warningList),
            null);
    }

    /// <summary>
    /// Detects a cycle in the scene's group-to-group containment graph.
    /// </summary>
    /// <remarks>
    /// Only edges that lead from one group to another can close a cycle; an edge
    /// to a non-group primitive is a leaf and cannot continue a path. The walk is
    /// iterative rather than recursive so a deeply nested scene cannot overflow
    /// the stack, and it colours nodes (unvisited, on the current path, done) so
    /// each node is visited once. A node reached again while still on the current
    /// path closes a cycle; reaching one already coloured done is a shared
    /// descendant, which is legal and not a cycle.
    /// </remarks>
    /// <param name="primitives">The primitives whose group links are inspected.</param>
    /// <returns><see langword="true"/> when a cycle exists.</returns>
    private static bool HasGroupCycle(IEnumerable<ScenePrimitive> primitives)
    {
        var groups = primitives
            .OfType<SceneGroup>()
            .ToDictionary(group => group.PrimitiveId, StringComparer.Ordinal);
        if (groups.Count == 0)
        {
            return false;
        }

        var onPath = new HashSet<string>(StringComparer.Ordinal);
        var done = new HashSet<string>(StringComparer.Ordinal);
        foreach (var start in groups.Keys)
        {
            if (done.Contains(start))
            {
                continue;
            }

            // Each stack entry is a group and the index of the child to visit
            // next, so a group is only marked done once all its children are.
            var path = new Stack<(SceneGroup Group, int NextChild)>();
            path.Push((groups[start], 0));
            _ = onPath.Add(start);

            while (path.Count > 0)
            {
                (SceneGroup group, var nextChild) = path.Pop();
                if (nextChild >= group.ChildPrimitiveIds.Count)
                {
                    _ = onPath.Remove(group.PrimitiveId);
                    _ = done.Add(group.PrimitiveId);
                    continue;
                }

                path.Push((group, nextChild + 1));
                var childId = group.ChildPrimitiveIds[nextChild];
                if (!groups.TryGetValue(childId, out SceneGroup? child))
                {
                    continue;
                }

                if (onPath.Contains(childId))
                {
                    return true;
                }

                if (done.Contains(childId))
                {
                    continue;
                }

                _ = onPath.Add(childId);
                path.Push((child, 0));
            }
        }

        return false;
    }

    private static IEnumerable<ScenePrimitive> Order(IEnumerable<ScenePrimitive> primitives) =>
        primitives
            .OrderBy(primitive => (int)primitive.ZLayer)
            .ThenBy(primitive => ActivitySubtypePriority.For(primitive.EntityType))
            .ThenBy(primitive => primitive.LaneOrder ?? int.MaxValue)
            .ThenBy(primitive => primitive.StackIndex ?? int.MaxValue)
            .ThenBy(primitive => primitive.SortOrder ?? int.MaxValue)
            .ThenBy(primitive => primitive.PrimitiveId, StringComparer.Ordinal);

    private static bool IsFinite(RectD rect) =>
        double.IsFinite(rect.X) &&
        double.IsFinite(rect.Y) &&
        double.IsFinite(rect.Width) &&
        double.IsFinite(rect.Height);

    private static bool IsValid(ScenePrimitive primitive) => primitive switch
    {
        SceneRect rect => IsFinite(rect.Bounds),
        SceneLine line => IsFinite(line.From) && IsFinite(line.To),
        ScenePolygon polygon => polygon.Points.Count >= 3 && polygon.Points.All(IsFinite),
        SceneText text => IsFinite(text.TextBounds),
        SceneGroup => true,
        _ => false,
    };

    private static bool IsFinite(PointD point) =>
        double.IsFinite(point.X) && double.IsFinite(point.Y);
}
