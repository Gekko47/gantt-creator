namespace GanttCreator.Core.Tests;

/// <summary>
/// Tests for <see cref="EntityHierarchyCatalog"/>, the single authority for
/// which entity types may own or be a child.
/// </summary>
public sealed class EntityHierarchyCatalogTests
{
    /// <summary>
    /// The exhaustiveness guard. Every <see cref="GanttEntityType"/> member must
    /// be decided by the matrix: classified as a parent, classified as a child,
    /// or reported as excluded. A type in neither list and not excluded would
    /// silently default, which is the failure this test prevents.
    /// </summary>
    [Fact]
    public void Every_catalogue_type_is_classified_by_the_hierarchy_matrix()
    {
        var classified = EntityHierarchyCatalog.TypesThatMayOwnChildren()
            .Concat(EntityHierarchyCatalog.TypesThatMayBeChild())
            .ToHashSet();

        var all = Enum.GetValues<GanttEntityType>();

        foreach (GanttEntityType type in all)
        {
            bool decided =
                classified.Contains(type)
                || EntityHierarchyCatalog.IsExcludedFromHierarchy(type);

            Assert.True(
                decided,
                $"'{type}' is neither classified nor reported as excluded from the hierarchy.");
        }
    }

    /// <summary>
    /// The excluded set is exactly the structural remainder, so
    /// <see cref="EntityHierarchyCatalog.IsExcludedFromHierarchy"/> cannot
    /// disagree with the two tables it is derived from.
    /// </summary>
    [Fact]
    public void Excluded_set_is_exactly_the_structural_types()
    {
        var classified = EntityHierarchyCatalog.TypesThatMayOwnChildren()
            .Concat(EntityHierarchyCatalog.TypesThatMayBeChild())
            .ToHashSet();

        var excluded = Enum
            .GetValues<GanttEntityType>()
            .Where(EntityHierarchyCatalog.IsExcludedFromHierarchy)
            .ToHashSet();

        Assert.Equal(3, excluded.Count);
        Assert.Contains(GanttEntityType.Splitter, excluded);
        Assert.Contains(GanttEntityType.Spacer, excluded);
        Assert.Contains(GanttEntityType.Delineator, excluded);

        // Nothing is both classified and excluded.
        Assert.Empty(classified.Intersect(excluded));
    }

    /// <summary>
    /// The matrix must name only members of <see cref="GanttEntityType"/>. A
    /// parallel type list is the specific failure this test prevents.
    /// </summary>
    [Fact]
    public void Matrix_names_only_real_entity_types()
    {
        var all = Enum.GetValues<GanttEntityType>().ToHashSet();

        foreach (GanttEntityType type in EntityHierarchyCatalog.TypesThatMayOwnChildren())
        {
            Assert.Contains(type, all);
        }

        foreach (GanttEntityType type in EntityHierarchyCatalog.TypesThatMayBeChild())
        {
            Assert.Contains(type, all);
        }
    }

    /// <summary>
    /// A Critical Interval is both a child of an activity and a parent, but the
    /// two are not independent: a Critical Interval may own a child only while it
    /// is TOP-LEVEL. Once it is a child itself, any child of its own is depth 3
    /// and the uniform depth check refuses it. This asserts the capability pair
    /// separately from that depth rule, which
    /// <c>GanttHierarchyLimitsTests</c> and <c>ProjectionResolverTests</c> own.
    /// </summary>
    [Fact]
    public void Critical_interval_may_be_both_a_child_and_a_parent()
    {
        Assert.True(EntityHierarchyCatalog.MayBeChild(GanttEntityType.CriticalInterval));
        Assert.True(EntityHierarchyCatalog.MayOwnChildren(GanttEntityType.CriticalInterval));
    }

    /// <summary>
    /// Span activities and procurements may both own and be children. A child
    /// activity sharing its parent's lane is REV5's <c>StackOnParent</c> case.
    /// </summary>
    [Theory]
    [InlineData(GanttEntityType.AsBuiltActivity)]
    [InlineData(GanttEntityType.AsPlannedActivity)]
    [InlineData(GanttEntityType.BaselineActivity)]
    [InlineData(GanttEntityType.DelayEvent)]
    [InlineData(GanttEntityType.AsBuiltProcurement)]
    [InlineData(GanttEntityType.AsPlannedProcurement)]
    [InlineData(GanttEntityType.BaselineProcurement)]
    [InlineData(GanttEntityType.CustomActivity)]
    public void Span_types_may_own_children_and_be_children(GanttEntityType type)
    {
        Assert.True(EntityHierarchyCatalog.MayOwnChildren(type));
        Assert.True(EntityHierarchyCatalog.MayBeChild(type));
    }

    /// <summary>
    /// Milestones may be children (REV5's <c>MilestoneOnParent</c> case) but a
    /// milestone is a point event, so it does not own children.
    /// </summary>
    [Theory]
    [InlineData(GanttEntityType.AsBuiltMilestone)]
    [InlineData(GanttEntityType.AsPlannedMilestone)]
    [InlineData(GanttEntityType.BaselineMilestone)]
    [InlineData(GanttEntityType.CriticalMilestone)]
    public void Milestones_may_be_children_but_never_own_them(GanttEntityType type)
    {
        Assert.True(EntityHierarchyCatalog.MayBeChild(type));
        Assert.False(EntityHierarchyCatalog.MayOwnChildren(type));
    }

    /// <summary>
    /// Structural rows take no part in a hierarchy in either direction. This is
    /// the positive assertion for the widened <c>ParentId</c> rule: a
    /// <c>Splitter</c> carrying a <c>ParentId</c> is not permitted, so it still
    /// reports <c>NotUsedByType</c> rather than becoming silently accepted.
    /// </summary>
    [Theory]
    [InlineData(GanttEntityType.Splitter)]
    [InlineData(GanttEntityType.Spacer)]
    [InlineData(GanttEntityType.Delineator)]
    public void Structural_types_are_excluded_in_both_directions(GanttEntityType type)
    {
        Assert.False(EntityHierarchyCatalog.MayOwnChildren(type));
        Assert.False(EntityHierarchyCatalog.MayBeChild(type));
        Assert.True(EntityHierarchyCatalog.IsExcludedFromHierarchy(type));
    }

    /// <summary>
    /// The widened rule: a <c>ParentId</c> is meaningful exactly when the type is
    /// child-capable. Anything else reports "not used by type".
    /// </summary>
    [Fact]
    public void ParentId_relevance_tracks_child_capability_exactly()
    {
        foreach (GanttEntityType type in EntityHierarchyCatalog.TypesThatMayBeChild())
        {
            Assert.True(EntityHierarchyCatalog.ParentIdIsRelevant(type));
        }

        foreach (GanttEntityType type in Enum.GetValues<GanttEntityType>())
        {
            if (EntityHierarchyCatalog.MayBeChild(type))
            {
                continue;
            }

            Assert.False(
                EntityHierarchyCatalog.ParentIdIsRelevant(type),
                $"'{type}' is not child-capable, so a ParentId must not be relevant to it.");
        }
    }

    /// <summary>
    /// The seven-child maximum is the resolved product value, pinned so a change
    /// to it is a deliberate act rather than drift.
    /// </summary>
    [Fact]
    public void Maximum_children_is_seven()
    {
        Assert.Equal(7, EntityHierarchyCatalog.MaxChildrenPerParent);
    }

    /// <summary>
    /// Nesting is capped at two levels, so a grandchild is unrepresentable. This
    /// is the "explicitly rejected" branch of REV5 §4's requirement that
    /// multi-level nesting be supported or explicitly rejected.
    /// </summary>
    [Fact]
    public void Maximum_depth_is_two_levels()
    {
        Assert.Equal(2, EntityHierarchyCatalog.MaxDepth);
    }

    /// <summary>
    /// A type that may own children must be child-capable too, or the hierarchy
    /// could describe a parent that no child may reference.
    /// </summary>
    [Fact]
    public void Every_child_owning_type_is_also_child_capable()
    {
        foreach (GanttEntityType type in EntityHierarchyCatalog.TypesThatMayOwnChildren())
        {
            Assert.True(
                EntityHierarchyCatalog.MayBeChild(type),
                $"'{type}' may own children but may not itself be a child.");
        }
    }

    /// <summary>
    /// The returned lists are fresh and ordered, so a caller cannot mutate shared
    /// state or rely on insertion order.
    /// </summary>
    [Fact]
    public void Returned_lists_are_ordered_and_not_shared()
    {
        IReadOnlyList<GanttEntityType> first = EntityHierarchyCatalog.TypesThatMayBeChild();
        IReadOnlyList<GanttEntityType> second = EntityHierarchyCatalog.TypesThatMayBeChild();

        Assert.Equal(first, second);
        Assert.Equal([.. first.Order()], first);
        Assert.NotSame(first, second);
    }
}
