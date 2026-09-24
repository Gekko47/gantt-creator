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

        ScenePrimitive[] ordered = [.. Order(primitiveList)];
        SceneWarning[] warningList = [.. warnings ?? []];
        return new SceneCreationOutcome(
            new GanttScene(chartBounds, plotBounds, ordered, warningList),
            null);
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
