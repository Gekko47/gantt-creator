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

    /// <summary>
    /// The identifier factory kept returning an identifier that was already in use
    /// for longer than <see cref="CloneReseedPlanner.MaxIdGenerationAttempts"/>, so
    /// no fresh identifier could be found. Refused rather than looped, because an
    /// unbounded retry here hangs a Refresh with no output and no refusal.
    /// </summary>
    IdentifierFactoryExhausted = 3,
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
    /// How many candidates the identifier factory may be asked for before the
    /// reseed is refused.
    /// </summary>
    /// <remarks>
    /// A factory that only ever returns identifiers already present in the table
    /// would otherwise loop forever. The cap is a multiple of
    /// <see cref="MaxOccurrencesPerId"/> so a well-behaved factory returning a
    /// small pool of candidates still has room, while a factory that never
    /// produces a fresh identifier is refused rather than hanging the Refresh.
    /// </remarks>
    public const int MaxIdGenerationAttempts = MaxOccurrencesPerId * 1000;

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

            if (!NextUniqueId(newId, issued, out GanttRowId? replacement) || replacement is null)
            {
                return new CloneReseedOutcome(null, CloneReseedRefusal.IdentifierFactoryExhausted);
            }

            _ = issued.Add(replacement);
            map[id] = replacement;
        }

        return new CloneReseedOutcome(new CloneReseedPlan(map, map.Count), null);
    }

    /// <summary>
    /// Asks the factory for an identifier that is not already in use, up to
    /// <see cref="MaxIdGenerationAttempts"/> candidates.
    /// </summary>
    /// <param name="factory">The caller-supplied identifier factory.</param>
    /// <param name="used">The identifiers already in use; a fresh one is added to it.</param>
    /// <param name="id">The unused identifier, or <see langword="null"/> when the cap was reached.</param>
    /// <returns>
    /// <see langword="true"/> when an unused identifier was produced;
    /// <see langword="false"/> when the attempt cap was reached first.
    /// </returns>
    /// <remarks>
    /// The loop is bounded rather than <c>while (true)</c>. A factory that
    /// repeatedly returns an identifier already in the table — a broken generator,
    /// or a fixed seed colliding with the rows being reseeded — would otherwise
    /// spin forever inside a Refresh, with no scene, no warning and no refusal. The
    /// cap converts that into the same typed refusal every other failure here uses.
    /// </remarks>
    private static bool NextUniqueId(Func<GanttRowId> factory, HashSet<GanttRowId> used, out GanttRowId? id)
    {
        for (var attempt = 0; attempt < MaxIdGenerationAttempts; attempt++)
        {
            GanttRowId candidate = factory();
            if (used.Add(candidate))
            {
                id = candidate;
                return true;
            }
        }

        id = null;
        return false;
    }
}
