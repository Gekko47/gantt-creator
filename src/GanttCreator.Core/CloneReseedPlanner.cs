namespace GanttCreator.Core;

/// <summary>The typed outcome of planning a clone reseed.</summary>
/// <param name="Plan">The plan when one was produced, otherwise <see langword="null"/>.</param>
/// <param name="Refusal">The reason when no plan was produced, otherwise <see langword="null"/>.</param>
public sealed record CloneReseedOutcome(CloneReseedPlan? Plan, CloneReseedRefusal? Refusal)
{
    /// <summary>Gets whether a plan was produced.</summary>
    public bool Succeeded => Plan is not null;
}

/// <summary>Why a clone reseed could not be planned.</summary>
public enum CloneReseedRefusal
{
    /// <summary>No duplicated identifier was found, so there is no clone to reseed.</summary>
    NoDuplicatedIds = 0,

    /// <summary>
    /// The duplication cannot be resolved to a single structure, so any plan
    /// would be a guess.
    /// </summary>
    UnresolvableStructure = 1,

    /// <summary>Too many rows carry the same identifier to attribute them to one structure.</summary>
    TooManyDuplicates = 2,
}

/// <summary>
/// One row's identity change: which rows get a new identifier, and which
/// identifiers the rest of the structure must now point at.
/// </summary>
/// <param name="NewIdsByOldId">
/// The old-to-new identifier map. A parent and its child cloned together both
/// appear here, and a child's rewritten <c>ParentId</c> is looked up through this
/// same map so the cloned relationship survives under the new identities.
/// </param>
/// <param name="DuplicatedRowCount">How many rows carried a duplicated identifier.</param>
public sealed record CloneReseedPlan(
    IReadOnlyDictionary<GanttRowId, GanttRowId> NewIdsByOldId,
    int DuplicatedRowCount)
{
    /// <summary>
    /// The <c>ParentId</c> a row should carry after the reseed, or
    /// <see langword="null"/> when it becomes top-level. A parent reference is
    /// translated only when the parent was itself reseeded; an ambiguous reference
    /// is left alone rather than pointed at an arbitrary clone.
    /// </summary>
    /// <param name="currentParentId">The row's <c>ParentId</c> before the reseed.</param>
    /// <returns>The translated parent, or <see langword="null"/>.</returns>
    public GanttRowId? TranslateParentId(GanttRowId? currentParentId) =>
        currentParentId is not null && NewIdsByOldId.TryGetValue(currentParentId, out GanttRowId? translated)
            ? translated
            : null;
}

/// <summary>
/// Plans the atomic identity reseed of a duplicated managed structure.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this is a refusal-capable plan rather than a repair.</b> When a user
/// duplicates a worksheet or table, the clone arrives carrying the original's
/// identifiers. Left alone, two structures share one identity space and every
/// parent reference is ambiguous. R4.7A D7 requires an <b>atomic</b> reseed:
/// the whole old-to-new map is built first, and either every row is rewritten
/// under it or nothing is. A half-reseeded hierarchy is strictly worse than a
/// duplicated one, because the damage is no longer recognisable as duplication.
/// </para>
/// <para>
/// <b>What already exists.</b> <c>ExcelGanttRowIdentityRepairer</c> already
/// performs the per-row write and the <c>ParentId</c> rewrite. This type supplies
/// the part that was missing: deciding <em>which</em> rows form one clone, and
/// refusing when that cannot be decided. It is pure, so the decision is testable
/// without a workbook.
/// </para>
/// <para>
/// <b>Refusal is the common case for real damage.</b> If an identifier appears on
/// more than two rows, the duplication cannot be attributed to one clone, and
/// guessing which rows were cloned together would rewrite a relationship the user
/// never created.
/// </para>
/// </remarks>
public static class CloneReseedPlanner
{
    /// <summary>The largest number of rows that may carry one identifier and still be attributable to a clone.</summary>
    public const int MaxOccurrencesPerId = 2;

    /// <summary>
    /// Plans a reseed of every duplicated identifier in the supplied rows.
    /// </summary>
    /// <param name="rows">The rows to inspect, in worksheet order. Not mutated.</param>
    /// <param name="newId">Generates a fresh, unused identifier. Supplied so the caller controls identity generation and the test controls determinism.</param>
    /// <returns>
    /// A typed outcome. <see cref="CloneReseedRefusal.NoDuplicatedIds"/> when
    /// there is nothing to reseed, and
    /// <see cref="CloneReseedRefusal.TooManyDuplicates"/> when an identifier
    /// cannot be attributed to a single clone.
    /// </returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="rows"/> or <paramref name="newId"/> is null.</exception>
    public static CloneReseedOutcome Plan(
        IReadOnlyList<GanttRowId> rows,
        Func<GanttRowId> newId)
    {
        ArgumentNullException.ThrowIfNull(rows);
        ArgumentNullException.ThrowIfNull(newId);

        Dictionary<GanttRowId, int> occurrences = [];
        foreach (GanttRowId id in rows)
        {
            _ = occurrences.TryGetValue(id, out var count);
            occurrences[id] = count + 1;
        }

        // Ordered by identifier text, not by the enum's natural order: GanttRowId
        // implements IEquatable but deliberately not IComparable, so LINQ's default
        // key comparer throws. Ordinal on the exact workbook text is also the order
        // the rest of the identity code already uses, so the plan is deterministic
        // and consistent with the repairer it feeds.
        List<GanttRowId> duplicated =
        [
            .. occurrences
                .Where(static p => p.Value > 1)
                .Select(static p => p.Key)
                .OrderBy(static id => id.Value, StringComparer.Ordinal)
        ];

        if (duplicated.Count == 0)
        {
            return new CloneReseedOutcome(null, CloneReseedRefusal.NoDuplicatedIds);
        }

        if (occurrences.Values.Any(static count => count > MaxOccurrencesPerId))
        {
            return new CloneReseedOutcome(null, CloneReseedRefusal.TooManyDuplicates);
        }

        // The first occurrence of each duplicated identifier is the canonical row
        // and keeps its identity; the later one is the clone and is reseeded. This
        // is the same first-canonical rule GanttRowValidator already applies, so
        // the plan and the validator cannot disagree about which row is which.
        HashSet<GanttRowId> seen = [];
        Dictionary<GanttRowId, GanttRowId> map = [];
        HashSet<GanttRowId> issued = [.. rows];

        foreach (GanttRowId id in rows)
        {
            if (!occurrences[id].Equals(2) || seen.Add(id))
            {
                continue;
            }

            GanttRowId replacement = NextUniqueId(newId, issued);
            _ = issued.Add(replacement);
            map[id] = replacement;
        }

        return new CloneReseedOutcome(new CloneReseedPlan(map, map.Count), null);
    }

    private static GanttRowId NextUniqueId(Func<GanttRowId> factory, HashSet<GanttRowId> used)
    {
        while (true)
        {
            GanttRowId candidate = factory();
            if (used.Add(candidate))
            {
                return candidate;
            }
        }
    }
}
