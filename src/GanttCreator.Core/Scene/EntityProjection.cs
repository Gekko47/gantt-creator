namespace GanttCreator.Core.Scene;

/// <summary>
/// How an entity is placed relative to the lane its own row would imply.
/// </summary>
public enum ProjectionMode
{
    /// <summary>The entity owns a lane of its own. A top-level activity or milestone.</summary>
    OwnLane = 0,

    /// <summary>The entity draws on its parent's lane as a line-like overlay. A Critical Interval.</summary>
    OverlayParent = 1,

    /// <summary>The entity draws on its parent's lane as a bar that may overlap siblings. A child activity.</summary>
    StackOnParent = 2,

    /// <summary>The entity draws on its parent's lane as a milestone. A child milestone.</summary>
    MilestoneOnParent = 3,

    /// <summary>The entity spans the whole plot and owns no lane. A delineator.</summary>
    PlotGlobal = 4,

    /// <summary>The entity occupies a structural lane of its own. A splitter or spacer.</summary>
    Structural = 5,
}

/// <summary>
/// One entity's resolved projection: which lane it draws on, and how.
/// </summary>
/// <param name="SourceEntityId">The entity this projection is for.</param>
/// <param name="RenderLaneOwnerId">
/// The entity whose lane this one draws on. Equal to <paramref name="SourceEntityId"/>
/// for an own-lane, structural or plot-global entity.
/// </param>
/// <param name="Mode">How the entity is placed on that lane.</param>
public sealed record EntityProjection(
    GanttRowId SourceEntityId,
    GanttRowId RenderLaneOwnerId,
    ProjectionMode Mode)
{
    /// <summary>
    /// Whether the entity renders on a lane other than its own row's. A projected
    /// child creates no lane and consumes no lane height.
    /// </summary>
    public bool IsProjected => RenderLaneOwnerId != SourceEntityId;
}

/// <summary>Why a projection could not be resolved.</summary>
public enum ProjectionRefusal
{
    /// <summary>The input collection was null.</summary>
    NullInput = 0,

    /// <summary>An input item or its event was null.</summary>
    NullEvent = 1,

    /// <summary>Two input events carried the same stable row ID.</summary>
    DuplicateEventId = 2,

    /// <summary>
    /// A child named a parent that is not in the batch, or named a duplicate, so the
    /// lane it draws on cannot be known.
    /// </summary>
    UnresolvableParent = 3,

    /// <summary>
    /// A child of a type that may not own children, or a child of a type the
    /// hierarchy matrix does not permit.
    /// </summary>
    IncompatibleHierarchy = 4,

    /// <summary>
    /// The parent chain is deeper than
    /// <see cref="EntityHierarchyCatalog.MaxDepth"/>, so the render lane is ambiguous.
    /// </summary>
    HierarchyTooDeep = 5,
}

/// <summary>The typed result of attempting to resolve projections.</summary>
/// <param name="Projections">The resolved projections in worksheet row order.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record ProjectionResolution(
    IReadOnlyList<EntityProjection> Projections,
    ProjectionRefusal? Refusal)
{
    /// <summary>Gets whether resolution succeeded.</summary>
    public bool Succeeded => Refusal is null;
}

/// <summary>
/// Resolves every entity's render lane and projection mode, in Core, before any
/// scene is built.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this lives in Core and not in the Office layer.</b> Placement must not be
/// decided by reading the worksheet. If Excel decided which lane a row draws on, the
/// same table could produce two different charts depending on how it was read, and a
/// projected child could be placed differently by a different command.
/// </para>
/// <para>
/// <b>The source row is not the render lane.</b> A projected child has its own source
/// row — that is where the user types its Type, dates and description — but its shape
/// renders on the parent's lane. Conflating the two is what makes collapsing a parent
/// appear to delete its children.
/// </para>
/// <para>
/// <b>Capability comes from <see cref="EntityHierarchyCatalog"/>, never from a
/// second list.</b> R4.7A introduced that matrix as the single authority for which
/// types may own or be a child; this resolver consumes it, so a type cannot be a valid
/// child for Add Child and an invalid child for the scene.
/// </para>
/// <para>
/// <b>Refusal is whole.</b> A table with one unresolvable parent produces no partial
/// assignment, because half-projected children are indistinguishable from correctly
/// projected ones once rendered.
/// </para>
/// </remarks>
public static class ProjectionResolver
{
    /// <summary>
    /// Resolves each event's projection. A top-level event owns its lane; a child
    /// resolves to its parent's lane; a delineator is plot-global; a splitter or
    /// spacer is structural.
    /// </summary>
    /// <param name="events">The validated events, in any order. Not mutated.</param>
    /// <returns>
    /// The resolutions in worksheet row order, or a typed refusal. An empty input is
    /// a successful empty result, matching <see cref="LaneLayoutBuilder"/>: a scene of
    /// only delineators legitimately owns no lanes.
    /// </returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="events"/> is null.</exception>
    public static ProjectionResolution Resolve(IReadOnlyList<GanttEvent>? events)
    {
        if (events is null)
        {
            return new ProjectionResolution([], ProjectionRefusal.NullInput);
        }

        Dictionary<GanttRowId, GanttEvent> canonicalById = [];
        HashSet<GanttRowId> duplicated = [];

        foreach (GanttEvent @event in events)
        {
            if (@event is null)
            {
                return new ProjectionResolution([], ProjectionRefusal.NullEvent);
            }

            if (!canonicalById.TryAdd(@event.Id, @event))
            {
                _ = duplicated.Add(@event.Id);
            }
        }

        if (duplicated.Count > 0)
        {
            return new ProjectionResolution([], ProjectionRefusal.DuplicateEventId);
        }

        Dictionary<GanttRowId, int> rowById = [];
        foreach (GanttEvent @event in events)
        {
            rowById[@event.Id] = @event.RowNumber;
        }

        List<EntityProjection> resolved = [];

        foreach (GanttEvent @event in events.OrderBy(static e => e.RowNumber).ThenBy(static e => e.Id.Value, StringComparer.Ordinal))
        {
            ProjectionMode mode = ModeFor(@event);

            if (mode is ProjectionMode.PlotGlobal or ProjectionMode.Structural or ProjectionMode.OwnLane)
            {
                resolved.Add(new EntityProjection(@event.Id, @event.Id, mode));
                continue;
            }

            // Every remaining mode is "a child drawing on a parent's lane", so the
            // parent must resolve first.
            if (@event.ParentId is not { } parentId || !canonicalById.TryGetValue(parentId, out GanttEvent? parent))
            {
                return new ProjectionResolution([], ProjectionRefusal.UnresolvableParent);
            }

            if (!EntityHierarchyCatalog.MayBeChild(@event.Type)
                || !EntityHierarchyCatalog.MayOwnChildren(parent.Type)
                || !EntityHierarchyCatalog.MayBeChild(parent.Type))
            {
                return new ProjectionResolution([], ProjectionRefusal.IncompatibleHierarchy);
            }

            // Depth: the lane owner must itself be top-level. A Critical Interval may
            // parent another Critical Interval while remaining a child of an activity,
            // so the test is on the parent of the CHILD, not on the child's own depth
            // -- the same distinction the validator's depth check uses.
            if (parent.ParentId is not null)
            {
                return new ProjectionResolution([], ProjectionRefusal.HierarchyTooDeep);
            }

            resolved.Add(new EntityProjection(@event.Id, parentId, mode));
        }

        return new ProjectionResolution(
            [.. resolved.OrderBy(p => rowById[p.SourceEntityId]).ThenBy(p => p.SourceEntityId.Value, StringComparer.Ordinal)],
            null);
    }

    /// <summary>
    /// The projection mode for one event, derived from its type and whether it is a
    /// child. Every type is covered explicitly, so a type added later cannot silently
    /// take the default.
    /// </summary>
    /// <param name="atEvent">The validated event.</param>
    /// <returns>The mode implied by the type's role.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="atEvent"/> is null.</exception>
    public static ProjectionMode ModeFor(GanttEvent atEvent)
    {
        ArgumentNullException.ThrowIfNull(atEvent);

        return atEvent.Type switch
        {
            // Plot-global: §24 makes a delineator a full-height line consuming no lane.
            GanttEntityType.Delineator => ProjectionMode.PlotGlobal,

            // Structural: ADR-0021 made these lane participants of their own.
            GanttEntityType.Splitter or GanttEntityType.Spacer => ProjectionMode.Structural,

            GanttEntityType.CriticalInterval => atEvent.ParentId is null
                ? ProjectionMode.OwnLane
                : ProjectionMode.OverlayParent,

            GanttEntityType.AsBuiltMilestone
                or GanttEntityType.AsPlannedMilestone
                or GanttEntityType.BaselineMilestone
                or GanttEntityType.CriticalMilestone => atEvent.ParentId is null
                ? ProjectionMode.OwnLane
                : ProjectionMode.MilestoneOnParent,

            GanttEntityType.AsBuiltActivity
                or GanttEntityType.AsPlannedActivity
                or GanttEntityType.BaselineActivity
                or GanttEntityType.DelayEvent
                or GanttEntityType.AsBuiltProcurement
                or GanttEntityType.AsPlannedProcurement
                or GanttEntityType.BaselineProcurement
                or GanttEntityType.CustomActivity => atEvent.ParentId is null
                ? ProjectionMode.OwnLane
                : ProjectionMode.StackOnParent,

            _ => ProjectionMode.OwnLane,
        };
    }
}
