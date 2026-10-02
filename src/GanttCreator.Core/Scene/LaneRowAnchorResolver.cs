namespace GanttCreator.Core.Scene;

/// <summary>
/// One lane's vertical anchor: the measured worksheet row the lane renders on
/// (ADR-0034 D1).
/// </summary>
/// <param name="LaneKey">
/// The lane's stable key, exactly as <see cref="LaneOrdering.LaneKey"/> produced it
/// for the lane's owning row. The anchor is looked up by this key, so an anchor can
/// never be applied to a different lane than the one it was resolved for.
/// </param>
/// <param name="TopOffsetPt">
/// The lane's top edge, measured from the plot's own top — which ADR-0030 D1 makes
/// the first body row's measured top. The lane's absolute worksheet top is therefore
/// <c>plotTop + TopOffsetPt</c>, which is the owning row's own top edge, so the
/// offset is a subtraction of the same origin the renderer already adds.
/// </param>
/// <param name="HeightPt">The owning row's measured height, in points.</param>
public sealed record LaneRowAnchor(string LaneKey, double TopOffsetPt, double HeightPt);

/// <summary>Why a lane could not be anchored to its worksheet row.</summary>
public enum LaneRowAnchorRefusal
{
    /// <summary>The lane-input collection was null.</summary>
    NullInput = 0,

    /// <summary>The event collection was null.</summary>
    NullEvents = 1,

    /// <summary>The measured grid was null.</summary>
    NullGrid = 2,

    /// <summary>
    /// A lane's owning row could not be identified among the supplied events, so the
    /// row it must coincide with is unknown. Refused rather than defaulted: an
    /// invented anchor would place the lane at a plausible but wrong height, which is
    /// the exact silent misalignment this resolver exists to remove.
    /// </summary>
    UnknownLaneOwner = 3,

    /// <summary>
    /// A lane's owning row is beyond the measured body rows, so the row's bounds are
    /// unknown. This is the guard against a measurement that returned fewer rows than
    /// the events reference: without it, the row lookup would fall off the end of the
    /// measured list and the lanes below it would be anchored to whatever came last.
    /// </summary>
    RowOutsideMeasuredRows = 4,

    /// <summary>
    /// A lane's owning row measures zero height, which cannot hold a lane. A hidden
    /// row measures zero and is legitimately consumed by nothing (ADR-0034), but a
    /// row that owns a lane is a visible row, so a zero here is a real disagreement
    /// between the hierarchy and the sheet rather than a row to be skipped.
    /// </summary>
    DegenerateRowHeight = 5,
}

/// <summary>The typed result of resolving lane anchors.</summary>
/// <param name="Anchors">The resolved anchors, in lane order.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record LaneRowAnchorResolution(
    IReadOnlyList<LaneRowAnchor> Anchors,
    LaneRowAnchorRefusal? Refusal)
{
    /// <summary>Gets whether every lane was anchored.</summary>
    public bool Succeeded => Refusal is null;

    /// <summary>
    /// Finds the anchor for a lane key.
    /// </summary>
    /// <param name="laneKey">The lane's stable key.</param>
    /// <param name="anchor">The anchor, when one exists for that key.</param>
    /// <returns><see langword="true"/> when the lane was anchored.</returns>
    /// <remarks>
    /// A lane with no anchor is not an error at this point: a caller may lay out a
    /// subset of the table, and the layout builder then falls back to its own
    /// stacking. The resolver refuses only when a lane it was asked about cannot be
    /// anchored at all.
    /// </remarks>
    public bool TryGet(string laneKey, out LaneRowAnchor? anchor)
    {
        foreach (LaneRowAnchor candidate in Anchors)
        {
            if (string.Equals(candidate.LaneKey, laneKey, StringComparison.Ordinal))
            {
                anchor = candidate;
                return true;
            }
        }

        anchor = null;
        return false;
    }
}

/// <summary>
/// Resolves each lane to the measured worksheet row it must coincide with, which is
/// the half of ADR-0026 D3 that <see cref="LaneMetricsResolver"/> cannot express
/// (ADR-0034 D1).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why a lane needs an anchor at all.</b> <see cref="LaneLayoutBuilder"/> stacked
/// lanes as <c>laneTop += lane.Height</c>, so lane <em>k</em> coincided with body row
/// <em>k</em> only while every row above it had opened a lane of its own. A row that
/// owns no lane — a <c>Delineator</c> (§24 makes it plot-global), a projected child
/// (ADR-0026 D7) — consumed no height, so every lane below it rendered one row too
/// high and the leftover space was stranded at the bottom of the plot. That is the
/// reported defect: shapes that no longer line up with their cells.
/// </para>
/// <para>
/// <b>Why the row is found by ORDINAL, not by <c>RowNumber</c> arithmetic.</b> The
/// measured heights are positional, so the only thing an anchor needs is the row's
/// position among the body rows. Taking the ordinal of the events in
/// <see cref="GanttEvent.RowNumber"/> order gives that position without assuming
/// whether a particular caller numbers its rows from one or from the header row —
/// the canonical fixture numbers them like worksheet rows while the live reader
/// numbers them as body indices, and an anchor that depended on the difference would
/// be right for one and silently wrong for the other.
/// </para>
/// <para>
/// <b>The anchor is relative to the plot's own top.</b> A lane's absolute position is
/// <c>plotTop + TopOffsetPt</c>, and the renderer already adds <c>plotTop</c> once
/// (ADR-0030 D1 makes that the first body row's measured top). Storing an absolute
/// row top here would therefore double the offset.
/// </para>
/// <para>
/// Pure and Office-free: it reads the measured grid and the validated events, so a
/// test can prove the anchoring without a host.
/// </para>
/// </remarks>
public static class LaneRowAnchorResolver
{
    /// <summary>Resolves an anchor for every lane in the supplied lane inputs.</summary>
    /// <param name="laneInputs">
    /// The lane participants, exactly as <see cref="LaneLayoutBuilder"/> receives
    /// them. A <c>Delineator</c> is absent by design: it owns no lane.
    /// </param>
    /// <param name="events">
    /// Every validated event, including the rows that own no lane. The full set is
    /// required because a lane's position is its row's position among <em>all</em>
    /// body rows, not among the lane-owning ones.
    /// </param>
    /// <param name="grid">The measured grid, whose row tops are the anchor positions.</param>
    /// <returns>The anchors in lane order, or a typed refusal.</returns>
    public static LaneRowAnchorResolution TryResolve(
        IReadOnlyList<LaneEventInput>? laneInputs,
        IReadOnlyList<GanttEvent>? events,
        PanelCellGrid? grid)
    {
        if (laneInputs is null)
        {
            return Refused(LaneRowAnchorRefusal.NullInput);
        }

        if (events is null)
        {
            return Refused(LaneRowAnchorRefusal.NullEvents);
        }

        if (grid is null)
        {
            return Refused(LaneRowAnchorRefusal.NullGrid);
        }

        // The body-row ordinal of every event: its position in row order, which is
        // the position its measured height occupies in the grid.
        Dictionary<GanttRowId, int> ordinalById = [];
        var ordinal = 0;
        foreach (GanttEvent @event in events
            .OrderBy(static candidate => candidate.RowNumber)
            .ThenBy(static candidate => candidate.Id.Value, StringComparer.Ordinal))
        {
            // First wins: a duplicate id is the validator's error to report, and a
            // last-wins ordinal would make the anchor depend on input order.
            _ = ordinalById.TryAdd(@event.Id, ordinal);
            ordinal++;
        }

        // Grouped exactly as LaneLayoutBuilder groups, so the keys cannot drift.
        List<LaneRowAnchor> anchors = [];
        foreach (IGrouping<string, LaneEventInput> group in LaneOrdering
            .OrderForLanes(laneInputs)
            .GroupBy(LaneOrdering.LaneKey))
        {
            LaneEventInput? owner = null;
            foreach (LaneEventInput input in group
                .OrderBy(static candidate => candidate.Event.RowNumber)
                .ThenBy(static candidate => candidate.Event.Id.Value, StringComparer.Ordinal))
            {
                // A projected child draws on its owner's lane, so the owning row is
                // the first member that owns its lane. A lane of only projected
                // members cannot happen: LaneKey resolves such a member to its
                // owner's key, and an unresolvable owner leaves it row-scoped.
                if (!input.IsProjected)
                {
                    owner = input;
                    break;
                }
            }

            if (owner is null || !ordinalById.TryGetValue(owner.Event.Id, out var ownerOrdinal))
            {
                return Refused(LaneRowAnchorRefusal.UnknownLaneOwner);
            }

            if (ownerOrdinal >= grid.RowHeightsPt.Count)
            {
                return Refused(LaneRowAnchorRefusal.RowOutsideMeasuredRows);
            }

            var rowHeight = grid.RowHeightsPt[ownerOrdinal];
            if (rowHeight <= 0)
            {
                return Refused(LaneRowAnchorRefusal.DegenerateRowHeight);
            }

            anchors.Add(
                new LaneRowAnchor(
                    group.Key,
                    grid.RowTopsPt[ownerOrdinal] - grid.OriginTopPt,
                    rowHeight));
        }

        return new LaneRowAnchorResolution(anchors, null);
    }

    private static LaneRowAnchorResolution Refused(LaneRowAnchorRefusal refusal) => new([], refusal);
}
