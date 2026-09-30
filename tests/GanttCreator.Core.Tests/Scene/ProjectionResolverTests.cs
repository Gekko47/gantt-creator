using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// Tests for <see cref="ProjectionResolver"/> (R4.7B): every projection mode,
/// the source-row/render-lane split, and every typed refusal.
/// </summary>
public sealed class ProjectionResolverTests
{
    private static GanttRowId NewId() => GanttRowId.New();

    private static GanttEvent Event(
        int rowNumber,
        GanttRowId id,
        GanttEntityType type = GanttEntityType.AsPlannedActivity,
        GanttRowId? parentId = null) =>
        new(
            rowNumber,
            id,
            NewId(),
            0,
            type,
            "Build frame",
            new DateOnly(2026, 1, 1),
            new DateOnly(2026, 12, 31),
            parentId,
            "AsPlannedActivity",
            null,
            null,
            null,
            true,
            null);

    private static ProjectionResolution Resolve(params GanttEvent[] events) =>
        ProjectionResolver.Resolve(events);

    /// <summary>
    /// A top-level row owns its lane: the render lane owner is itself. This is the
    /// whole table in the pre-hierarchy model, so it must keep resolving unchanged.
    /// </summary>
    [Fact]
    public void A_top_level_activity_owns_its_own_lane()
    {
        GanttRowId id = NewId();

        ProjectionResolution resolution = Resolve(Event(2, id));

        Assert.True(resolution.Succeeded);
        EntityProjection projection = Assert.Single(resolution.Projections);
        Assert.Equal(id, projection.RenderLaneOwnerId);
        Assert.Equal(ProjectionMode.OwnLane, projection.Mode);
        Assert.False(projection.IsProjected);
    }

    /// <summary>
    /// The source-row/render-lane split: a child keeps its own source row and ID but
    /// draws on its parent's lane. Conflating the two is what makes a collapsed parent
    /// appear to delete its children.
    /// </summary>
    [Fact]
    public void A_child_activity_draws_on_its_parents_lane()
    {
        GanttRowId parent = NewId();
        GanttRowId child = NewId();

        ProjectionResolution resolution = Resolve(Event(2, parent), Event(3, child, parentId: parent));

        Assert.True(resolution.Succeeded);
        EntityProjection projected = Assert.Single(resolution.Projections, p => p.SourceEntityId == child);
        Assert.Equal(parent, projected.RenderLaneOwnerId);
        Assert.Equal(ProjectionMode.StackOnParent, projected.Mode);
        Assert.True(projected.IsProjected);
    }

    /// <summary>
    /// A Critical Interval child overlays its parent — the mode R4.7E's filled
    /// rectangle draws in.
    /// </summary>
    [Fact]
    public void A_critical_interval_child_overlays_its_parent()
    {
        GanttRowId parent = NewId();
        GanttRowId child = NewId();

        ProjectionResolution resolution = Resolve(
            Event(2, parent),
            Event(3, child, GanttEntityType.CriticalInterval, parent));

        EntityProjection projected = Assert.Single(resolution.Projections, p => p.SourceEntityId == child);
        Assert.Equal(parent, projected.RenderLaneOwnerId);
        Assert.Equal(ProjectionMode.OverlayParent, projected.Mode);
    }

    [Fact]
    public void A_child_milestone_projects_onto_its_parent()
    {
        GanttRowId parent = NewId();
        GanttRowId child = NewId();

        ProjectionResolution resolution = Resolve(
            Event(2, parent),
            Event(3, child, GanttEntityType.AsPlannedMilestone, parent));

        EntityProjection projected = Assert.Single(resolution.Projections, p => p.SourceEntityId == child);
        Assert.Equal(ProjectionMode.MilestoneOnParent, projected.Mode);
        Assert.Equal(parent, projected.RenderLaneOwnerId);
    }

    /// <summary>
    /// A delineator is plot-global and owns no lane; a splitter or spacer is
    /// structural. Neither is ever projected onto a parent, even if it somehow
    /// carried a ParentId the matrix would not permit.
    /// </summary>
    [Theory]
    [InlineData(GanttEntityType.Delineator, ProjectionMode.PlotGlobal)]
    [InlineData(GanttEntityType.Splitter, ProjectionMode.Structural)]
    [InlineData(GanttEntityType.Spacer, ProjectionMode.Structural)]
    public void Structural_and_plot_global_types_own_their_own_lane(
        GanttEntityType type,
        ProjectionMode expected)
    {
        GanttRowId id = NewId();

        ProjectionResolution resolution = Resolve(Event(2, id, type));

        EntityProjection projection = Assert.Single(resolution.Projections);
        Assert.Equal(expected, projection.Mode);
        Assert.Equal(id, projection.RenderLaneOwnerId);
        Assert.False(projection.IsProjected);
    }

    /// <summary>
    /// A Critical Interval may parent another Critical Interval while remaining a
    /// child of an activity. That is landed behaviour, not a new case, so the depth
    /// rule must not refuse it. The depth test is on the lane owner -- here the
    /// outer interval, which is itself a child of the activity -- and the inner
    /// interval still resolves onto the outer one's lane.
    /// </summary>
    [Fact]
    public void A_critical_interval_child_of_a_critical_interval_refuses_as_too_deep()
    {
        GanttRowId activity = NewId();
        GanttRowId outerInterval = NewId();
        GanttRowId innerInterval = NewId();

        ProjectionResolution resolution = Resolve(
            Event(2, activity),
            Event(3, outerInterval, GanttEntityType.CriticalInterval, activity),
            Event(4, innerInterval, GanttEntityType.CriticalInterval, outerInterval));

        // The inner interval's lane owner (the outer interval) is itself a child, so
        // the render lane is ambiguous. This is the same rule the validator applies
        // and the two must not disagree.
        Assert.False(resolution.Succeeded);
        Assert.Equal(ProjectionRefusal.HierarchyTooDeep, resolution.Refusal);
    }

    /// <summary>
    /// A child whose parent is absent cannot be placed, and the WHOLE resolution is
    /// refused rather than half-assigned.
    /// </summary>
    [Fact]
    public void A_child_of_an_absent_parent_refuses_the_whole_resolution()
    {
        ProjectionResolution resolution = Resolve(Event(2, NewId(), parentId: NewId()));

        Assert.False(resolution.Succeeded);
        Assert.Equal(ProjectionRefusal.UnresolvableParent, resolution.Refusal);
        Assert.Empty(resolution.Projections);
    }

    /// <summary>
    /// A duplicated Id has no single parent, so a child naming it is refused rather
    /// than pointed at an arbitrary row.
    /// </summary>
    [Fact]
    public void A_duplicate_id_refuses()
    {
        GanttRowId duplicated = NewId();
        GanttRowId child = NewId();

        ProjectionResolution resolution = Resolve(
            Event(2, duplicated),
            Event(3, duplicated),
            Event(4, child, parentId: duplicated));

        Assert.False(resolution.Succeeded);
        Assert.Equal(ProjectionRefusal.DuplicateEventId, resolution.Refusal);
    }

    /// <summary>
    /// A grandchild has no single render lane: its parent is itself a child, so the
    /// lane it would draw on is ambiguous. Refused, per ADR-0026 D7.
    /// </summary>
    [Fact]
    public void A_grandchild_refuses_as_too_deep()
    {
        GanttRowId parent = NewId();
        GanttRowId child = NewId();
        GanttRowId grandchild = NewId();

        ProjectionResolution resolution = Resolve(
            Event(2, parent),
            Event(3, child, parentId: parent),
            Event(4, grandchild, parentId: child));

        Assert.False(resolution.Succeeded);
        Assert.Equal(ProjectionRefusal.HierarchyTooDeep, resolution.Refusal);
    }

    /// <summary>
    /// A child of a type the hierarchy matrix does not permit is refused, so the
    /// scene cannot disagree with what Add Child would allow.
    /// </summary>
    [Fact]
    public void A_child_of_a_non_child_capable_parent_refuses()
    {
        GanttRowId parent = NewId();
        GanttRowId child = NewId();

        ProjectionResolution resolution = Resolve(
            Event(2, parent, GanttEntityType.Delineator),
            Event(3, child, parentId: parent));

        Assert.False(resolution.Succeeded);
        Assert.Equal(ProjectionRefusal.IncompatibleHierarchy, resolution.Refusal);
    }

    /// <summary>
    /// A structural type is never projected onto a parent, even when a hand-edited
    /// ParentId names one. <c>ModeFor</c> returns Structural for a Spacer regardless
    /// of its parent, so the reference is ignored rather than honoured — the row
    /// draws on its own structural lane.
    /// </summary>
    [Fact]
    public void A_structural_type_with_a_parent_id_is_not_projected()
    {
        GanttRowId parent = NewId();
        GanttRowId spacer = NewId();

        ProjectionResolution resolution = Resolve(
            Event(2, parent),
            Event(3, spacer, GanttEntityType.Spacer, parent));

        Assert.True(resolution.Succeeded);
        EntityProjection projection = Assert.Single(resolution.Projections, p => p.SourceEntityId == spacer);
        Assert.Equal(ProjectionMode.Structural, projection.Mode);
        Assert.Equal(spacer, projection.RenderLaneOwnerId);
    }

    /// <summary>
    /// An empty batch resolves successfully with nothing, matching
    /// <see cref="LaneLayoutBuilder"/>: a scene of only delineators owns no lanes.
    /// </summary>
    [Fact]
    public void An_empty_batch_resolves_successfully()
    {
        ProjectionResolution resolution = Resolve();

        Assert.True(resolution.Succeeded);
        Assert.Empty(resolution.Projections);
    }

    [Fact]
    public void A_null_batch_is_refused()
    {
        ProjectionResolution resolution = ProjectionResolver.Resolve(null);

        Assert.False(resolution.Succeeded);
        Assert.Equal(ProjectionRefusal.NullInput, resolution.Refusal);
    }

    /// <summary>
    /// Determinism: the result is in worksheet row order regardless of the order the
    /// events are supplied in, so the same table always yields the same sequence.
    /// </summary>
    [Fact]
    public void Projections_come_back_in_worksheet_row_order()
    {
        GanttRowId first = NewId();
        GanttRowId second = NewId();
        GanttRowId third = NewId();

        GanttEvent[] events =
        [
            Event(2, first),
            Event(3, second),
            Event(4, third, GanttEntityType.Delineator),
        ];

        ProjectionResolution ascending = Resolve(events);
        ProjectionResolution descending = Resolve([.. events.Reverse()]);

        Assert.Equal(ascending.Projections, descending.Projections);
        Assert.Equal(
            [first, second, third],
            ascending.Projections.Select(p => p.SourceEntityId));
    }

    /// <summary>
    /// Resolution does not mutate its input, so the caller can still apply a plan or
    /// abandon it.
    /// </summary>
    [Fact]
    public void Resolving_does_not_mutate_the_input_events()
    {
        GanttRowId parent = NewId();
        GanttEvent[] events = [Event(2, parent), Event(3, NewId(), parentId: parent)];
        GanttEvent[] snapshot = [.. events];

        _ = Resolve(events);

        Assert.Equal(snapshot, events);
    }

    /// <summary>
    /// Every type resolves to a mode explicitly, so a type added later cannot take
    /// the default silently. This is the projection equivalent of the repository's
    /// exhaustive-switch convention.
    /// </summary>
    [Fact]
    public void Every_catalogue_type_has_an_explicit_mode()
    {
        GanttRowId id = NewId();

        foreach (GanttEntityType type in Enum.GetValues<GanttEntityType>())
        {
            ProjectionMode ownLane = ProjectionResolver.ModeFor(Event(2, id, type));
            ProjectionMode asChild = ProjectionResolver.ModeFor(Event(2, id, type, id));

            Assert.True(Enum.IsDefined(ownLane));
            Assert.True(Enum.IsDefined(asChild));

            // A type that can be a child must change mode when it is given a parent,
            // otherwise the parent reference is ignored rather than honoured.
            if (EntityHierarchyCatalog.MayBeChild(type))
            {
                Assert.NotEqual(ProjectionMode.OwnLane, asChild);
            }
        }
    }
}
