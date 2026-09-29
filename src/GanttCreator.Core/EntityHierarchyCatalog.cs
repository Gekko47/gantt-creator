namespace GanttCreator.Core;

/// <summary>
/// The single Core authority for which <see cref="GanttEntityType"/> values may
/// participate in a parent/child hierarchy. Immutable; the static tables are
/// built once.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this exists separately from <see cref="EntityTypeCatalog"/>.</b> That
/// catalogue owns the user-selectable Type values and their display names,
/// date mode, colour capability and label positions. It does not answer "may
/// this row own a child?" — that is a hierarchy question, and the guide treats
/// hierarchy as a distinct concern. This type answers exactly two questions per
/// entity type and nothing else.
/// </para>
/// <para>
/// <b>Why it is a single authority.</b> Two consumers need the same answer: the
/// seven-child maximum (R4.7A D4) and the widening of <c>ParentId</c> from
/// Critical-Interval-only to every child-capable type (R4.7A D9). A second,
/// ad-hoc type list answering either question is how a seventh child becomes
/// valid for one command and invalid for another.
/// </para>
/// <para>
/// <b>No second type list.</b> Every entry is a member of
/// <see cref="GanttEntityType"/>. This type never declares a Type of its own, and
/// <c>EntityHierarchyCatalogTests</c> carries an exhaustiveness test that fails
/// if any <see cref="GanttEntityType"/> member is unclassified.
/// </para>
/// <para>
/// <b>Nesting depth.</b> Per the R4.7A guide, multi-level nesting beyond two
/// levels (a top-level parent, then children) is <b>rejected</b> rather than
/// supported. <see cref="MaxDepth"/> is the authority for that limit, and
/// <c>CriticalInterval</c> is a child only — it may not itself own children, so
/// a grandchild is unrepresentable rather than merely discouraged.
/// </para>
/// </remarks>
public static class EntityHierarchyCatalog
{
    /// <summary>
    /// The maximum number of children one parent may own. An eighth child is
    /// refused with a typed error carrying the parent id and the current count.
    /// </summary>
    public const int MaxChildrenPerParent = 7;

    /// <summary>
    /// The maximum hierarchy depth. A top-level parent is depth 1 and its
    /// children are depth 2; a depth-3 row is refused rather than ignored.
    /// </summary>
    public const int MaxDepth = 2;

    private static readonly HashSet<GanttEntityType> _mayOwnChildren =
    [
        GanttEntityType.AsBuiltActivity,
        GanttEntityType.AsPlannedActivity,
        GanttEntityType.BaselineActivity,
        GanttEntityType.DelayEvent,
        GanttEntityType.AsBuiltProcurement,
        GanttEntityType.AsPlannedProcurement,
        GanttEntityType.BaselineProcurement,
        GanttEntityType.CustomActivity,
    ];

    private static readonly HashSet<GanttEntityType> _mayBeChild =
    [
        GanttEntityType.CriticalInterval,
        GanttEntityType.AsBuiltActivity,
        GanttEntityType.AsPlannedActivity,
        GanttEntityType.BaselineActivity,
        GanttEntityType.DelayEvent,
        GanttEntityType.AsBuiltProcurement,
        GanttEntityType.AsPlannedProcurement,
        GanttEntityType.BaselineProcurement,
        GanttEntityType.CustomActivity,
        GanttEntityType.AsBuiltMilestone,
        GanttEntityType.AsPlannedMilestone,
        GanttEntityType.BaselineMilestone,
        GanttEntityType.CriticalMilestone,
    ];

    /// <summary>
    /// The entity types that may own children, in ascending enum order. Returned
    /// for display and tests; the set itself is not exposed for mutation.
    /// </summary>
    /// <returns>The child-owning types in ascending <see cref="GanttEntityType"/> order.</returns>
    public static IReadOnlyList<GanttEntityType> TypesThatMayOwnChildren() =>
        [.. _mayOwnChildren.Order()];

    /// <summary>
    /// The entity types that may appear as a child, in ascending enum order.
    /// </summary>
    /// <returns>The child-capable types in ascending <see cref="GanttEntityType"/> order.</returns>
    public static IReadOnlyList<GanttEntityType> TypesThatMayBeChild() =>
        [.. _mayBeChild.Order()];

    /// <summary>
    /// Whether the type may own children.
    /// </summary>
    /// <param name="type">The entity type to test.</param>
    /// <returns><see langword="true"/> when the type may own children.</returns>
    public static bool MayOwnChildren(GanttEntityType type) => _mayOwnChildren.Contains(type);

    /// <summary>
    /// Whether the type may appear as a child row.
    /// </summary>
    /// <param name="type">The entity type to test.</param>
    /// <returns><see langword="true"/> when the type may be a child.</returns>
    public static bool MayBeChild(GanttEntityType type) => _mayBeChild.Contains(type);

    /// <summary>
    /// Whether the type may carry a <c>ParentId</c> at all. This is the widened
    /// rule R4.7A D9 introduces: a row may name a parent exactly when its type is
    /// child-capable. Every other type continues to report
    /// <see cref="GanttValidationCodes.NotUsedByType"/> for a <c>ParentId</c>,
    /// so nothing becomes silently permitted.
    /// </summary>
    /// <param name="type">The entity type to test.</param>
    /// <returns><see langword="true"/> when a <c>ParentId</c> is meaningful.</returns>
    public static bool ParentIdIsRelevant(GanttEntityType type) => MayBeChild(type);

    /// <summary>
    /// Whether the type is excluded from hierarchy entirely. Structural rows
    /// (<c>Splitter</c>, <c>Spacer</c>) and <c>Delineator</c> are chart
    /// structure rather than schedule entities, so they neither own nor are
    /// children.
    /// </summary>
    /// <param name="type">The entity type to test.</param>
    /// <returns><see langword="true"/> when the type takes no part in a hierarchy.</returns>
    public static bool IsExcludedFromHierarchy(GanttEntityType type) =>
        !MayOwnChildren(type) && !MayBeChild(type);
}
