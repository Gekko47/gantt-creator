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
/// supported. <see cref="MaxDepth"/> is the authority for that limit, and it is
/// uniform: a parent that is itself a child puts any child of it at depth 3,
/// whatever that child's type is. A <c>CriticalInterval</c> is therefore a
/// level-2 child and cannot itself own a child. This matches the depth check in
/// <c>GanttRowValidator</c> and the check in <c>EntityProjection</c>, which must
/// agree or one would accept a hierarchy the other refuses to render. The cycle
/// walk in <c>GanttRowValidator</c> terminates at a span parent, and this
/// classification is what makes that termination correct rather incidental.
/// </para>
/// <para>
/// <b>R5.3 roles (owner ruling 2026-10-10).</b> <c>Splitter</c> and
/// <c>Spacer</c> are standalone top-level rows: they neither own nor are
/// children. Milestones (all four) are leaves: top-level or child, never
/// owning. Activities, procurement, delay events and <c>CustomActivity</c>
/// do both. <c>CriticalInterval</c> and <c>Delineator</c> are child-only and
/// never own; a child <c>Delineator</c> keeps <c>PlotGlobal</c> rendering and
/// its <c>ParentId</c> exists only so the outline can collapse it.
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

        // R5.3: Critical Interval is child-only and never owns (owner ruling
        // 2026-10-10). Milestones are leaves and never own. Splitter, Spacer
        // and Delineator never own either — none of them appear here.
    ];

    private static readonly HashSet<GanttEntityType> _mayBeChild =
    [
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
        GanttEntityType.CustomMilestone,

        // R5.3: Critical Interval is always a child; Delineator gains a
        // collapse-only ParentId (owner ruling 2026-10-10). Splitter and
        // Spacer remain standalone and may never be children.
        GanttEntityType.CriticalInterval,
        GanttEntityType.Delineator,
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
    /// Whether the type is excluded from hierarchy entirely. Standalone rows
    /// (<c>Splitter</c>, <c>Spacer</c>) are chart structure rather than schedule
    /// entities, so they neither own nor are children. Every other type takes
    /// part in exactly the direction its matrix entries allow.
    /// </summary>
    /// <param name="type">The entity type to test.</param>
    /// <returns><see langword="true"/> when the type takes no part in a hierarchy.</returns>
    public static bool IsExcludedFromHierarchy(GanttEntityType type) =>
        !MayOwnChildren(type) && !MayBeChild(type);
}
