using System.Globalization;

namespace GanttCreator.Core.Scene;

/// <summary>The reason a lane layout could not be created.</summary>
public enum LaneLayoutRefusal
{
    /// <summary>The event collection was null.</summary>
    NullInput = 0,

    /// <summary>The metrics input was null.</summary>
    NullMetrics = 1,

    /// <summary>One or more metrics were non-finite or negative.</summary>
    InvalidMetrics = 2,

    /// <summary>An input item or its event was null.</summary>
    NullEvent = 3,

    /// <summary>A resolved event height was non-finite or invalid for its type.</summary>
    InvalidEventHeight = 4,

    /// <summary>An effective compatibility stack value was negative.</summary>
    InvalidEffectiveStack = 5,

    /// <summary>Two input events carried the same stable row ID.</summary>
    DuplicateEventId = 6,

    /// <summary>
    /// A projected event named a render-lane owner that is not in this layout, and
    /// carried no compatibility stack value of its own, so the slot it belongs to
    /// cannot be determined. Refused rather than defaulted to a slot: the child would
    /// otherwise be placed in an arbitrary lane position and render somewhere its
    /// hierarchy does not put it.
    /// </summary>
    UnresolvedRenderLaneOwner = 7,
}

/// <summary>The typed result of attempting to build lane geometry.</summary>
/// <param name="Layout">The successful layout, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record LaneLayoutCreationOutcome(LaneLayoutResult? Layout, LaneLayoutRefusal? Refusal)
{
    /// <summary>Gets whether layout creation succeeded.</summary>
    public bool Succeeded => Layout is not null;
}

/// <summary>Builds deterministic vertical lane geometry without rendering or measuring.</summary>
public static class LaneLayoutBuilder
{
    private const string _ambiguousStackOverlapCode = "AmbiguousStackOverlap";

    /// <summary>
    /// The warning code for lane content that does not fit the fixed row height
    /// (R4.7D, ADR-0026 D7). The content is neither compressed nor accommodated by
    /// growing the lane; the user is told.
    /// </summary>
    public const string LaneContentExceedsLaneHeightCode = "LaneContentExceedsRowHeight";

    /// <summary>
    /// Derives a stable 32-hex-digit suffix from a lane key, for the synthetic
    /// warning owner described on the overflow warning.
    /// </summary>
    /// <remarks>
    /// A row id is <c>G-</c> plus 32 lowercase hex digits, so a lane warning can
    /// only borrow that shape if the suffix really is 32 hex digits. SHA-256 is
    /// used purely to obtain a stable 32-hex-digit string, not for any security
    /// purpose: this is not authentication, and a collision would merge two
    /// identical overflow messages rather than misreport anything.
    /// </remarks>
    private static string StableLaneWarningSuffix(string laneKey)
    {
        var digest = System.Security.Cryptography.SHA256.HashData(
            System.Text.Encoding.UTF8.GetBytes(laneKey));

        // Formatted as lowercase hex directly rather than via
        // Convert.ToHexString(...).ToLowerInvariant(): a GanttRowId is
        // `G-` plus 32 LOWERCASE hex digits, so the case is a contract, and
        // CA1308 (prefer upper-case) would otherwise push a diagnostic-only
        // preference onto a value the row-id parser rejects. This is formatting,
        // not comparison, so the rule does not apply to it.
        return string.Create(32, digest, static (span, bytes) =>
        {
            const string Digits = "0123456789abcdef";
            for (var i = 0; i < 16; i++)
            {
                span[i * 2] = Digits[bytes[i] >> 4];
                span[(i * 2) + 1] = Digits[bytes[i] & 0x0F];
            }
        });
    }

    /// <summary>Attempts to build deterministic lane geometry.</summary>
    /// <param name="events">
    /// The validated event layout inputs. An empty collection is valid and produces an
    /// empty layout: §24 makes a Delineator a plot-global entity that consumes no lane.
    /// </param>
    /// <param name="metrics">The resolved lane metrics.</param>
    /// <returns>A typed layout outcome.</returns>
    public static LaneLayoutCreationOutcome TryBuild(IReadOnlyList<LaneEventInput>? events, LaneLayoutMetrics? metrics)
    {
        if (events is null)
        {
            return Refused(LaneLayoutRefusal.NullInput);
        }

        if (metrics is null)
        {
            return Refused(LaneLayoutRefusal.NullMetrics);
        }

        if (!ValidMetrics(metrics))
        {
            return Refused(LaneLayoutRefusal.InvalidMetrics);
        }

        // An empty input is a successful empty layout, not a refusal. Entity guide
        // §24 makes a Delineator a full-height plot line that consumes no lane, so a
        // scene of only delineators legitimately has no lanes at all. Refusing here
        // forced `SceneBuilder` to call that scene "empty" and refuse it as
        // `EmptyEvents`, which wrongly reported a renderable scene as having nothing
        // to render. Whether the scene has any content at all is the orchestrator's
        // question, not this builder's.
        if (events.Count == 0)
        {
            return new LaneLayoutCreationOutcome(new LaneLayoutResult([], []), null);
        }

        HashSet<GanttRowId> ids = [];
        foreach (LaneEventInput? input in events)
        {
            if (input?.Event is null)
            {
                return Refused(LaneLayoutRefusal.NullEvent);
            }

            if (!double.IsFinite(input.ResolvedHeightPt) || input.ResolvedHeightPt < 0)
            {
                return Refused(LaneLayoutRefusal.InvalidEventHeight);
            }

            if (input.EffectiveStackIndex is { } stack && stack < 0)
            {
                return Refused(LaneLayoutRefusal.InvalidEffectiveStack);
            }

            if (!ids.Add(input.Event.Id))
            {
                return Refused(LaneLayoutRefusal.DuplicateEventId);
            }
        }

        IReadOnlyList<LaneEventInput> ordered = LaneOrdering.OrderForLanes(events);
        List<LaneGeometry> lanes = [];
        List<SceneWarning> warnings = [];
        double laneTop = 0;
        var laneOrder = 0;
        foreach (IGrouping<string, LaneEventInput> group in ordered.GroupBy(LaneOrdering.LaneKey))
        {
            LaneEventInput[] laneInputs = [.. group];
            var isSplitter = laneInputs[0].Event.Type == GanttEntityType.Splitter;
            var isSpacer = laneInputs[0].Event.Type == GanttEntityType.Spacer;
            LaneGeometry? lane =
                isSplitter || isSpacer
                    ? BuildFixedLane(group.Key, laneOrder, laneTop, laneInputs, isSplitter, isSpacer, metrics)
                    : BuildEventLane(group.Key, laneOrder, laneTop, laneInputs, metrics, warnings);

            // A typed refusal, not an exception: an unrepresentable hierarchy is a
            // reportable condition, and this builder's whole surface is TryBuild
            // returning a reason. Propagated as the layout's outcome so the caller
            // sees WHICH lane could not be laid out rather than a thrown
            // KeyNotFoundException from deep inside a private helper.
            if (lane is null)
            {
                return Refused(LaneLayoutRefusal.UnresolvedRenderLaneOwner);
            }

            lanes.Add(lane);
            laneTop += lane.Height;
            laneOrder++;
        }

        return new LaneLayoutCreationOutcome(new LaneLayoutResult(lanes, warnings), null);
    }

    private static LaneGeometry BuildFixedLane(
        string laneKey,
        int laneOrder,
        double laneTop,
        IReadOnlyList<LaneEventInput> inputs,
        bool isSplitter,
        bool isSpacer,
        LaneLayoutMetrics metrics
    )
    {
        var height = isSplitter ? metrics.SplitterHeightPt : metrics.SpacerHeightPt;
        GanttRowId[] eventIds = [.. inputs.Select(input => input.Event.Id)];
        SlotGeometry slot = new(0, 0, laneTop, laneTop + (height / 2), laneTop + height, height, eventIds);
        return new LaneGeometry(laneKey, laneOrder, laneTop, height, [slot], eventIds, isSplitter, isSpacer);
    }

    /// <summary>
    /// Builds one event lane's slots, or returns <see langword="null"/> when a
    /// projected child's render-lane owner cannot be resolved in this lane.
    /// </summary>
    /// <param name="laneKey">The lane's key, for the overflow warning owner.</param>
    /// <param name="laneOrder">The lane's zero-based order.</param>
    /// <param name="laneTop">The lane's top edge in points.</param>
    /// <param name="inputs">The lane's events.</param>
    /// <param name="metrics">The resolved lane metrics.</param>
    /// <param name="warnings">The scene warning sink.</param>
    /// <returns>The lane, or <see langword="null"/> for an unresolved render-lane owner.</returns>
    private static LaneGeometry? BuildEventLane(
        string laneKey,
        int laneOrder,
        double laneTop,
        LaneEventInput[] inputs,
        LaneLayoutMetrics metrics,
        List<SceneWarning> warnings
    )
    {
        LaneEventInput[] rowOrdered =
        [
            .. inputs.OrderBy(input => input.Event.RowNumber).ThenBy(input => input.Event.Id.Value, StringComparer.Ordinal),
        ];

        // R4.7B: a projected child takes its lane owner's effective stack index, so
        // it lands in the SAME slot and overlaps. Assigning it its own index by
        // position would give it a second slot, stacking it below the parent and
        // growing the lane -- which ADR-0026 D7 forbids and which is the whole
        // reason a child is projected rather than given a lane of its own.
        //
        // Two passes, because the owner is not guaranteed to precede the child in
        // row order: an authoring UI may place a child above its parent. Owners are
        // therefore indexed first, and only then do children look one up.
        var effectiveById = new Dictionary<GanttRowId, int>();
        var ownIndex = 0;

        foreach (LaneEventInput input in rowOrdered)
        {
            if (input.IsProjected)
            {
                continue;
            }

            effectiveById[input.Event.Id] = ownIndex;
            ownIndex++;
        }
        foreach (LaneEventInput input in rowOrdered)
        {
            if (!input.IsProjected)
            {
                continue;
            }

            // TryGetValue, not the indexer. A projected child whose owner is not in
            // this lane has no slot to inherit; the indexer threw a
            // KeyNotFoundException straight out of a builder whose entire contract is
            // to return a typed refusal, so one unrepresentable hierarchy crashed the
            // caller instead of being reported. The owner may legitimately be absent
            // when a caller lays out a subset of the table.
            if (input.RenderLaneOwner is { } owner
                && effectiveById.TryGetValue(owner.Id, out var ownerIndex))
            {
                effectiveById[input.Event.Id] = ownerIndex;
            }
            else if (input.EffectiveStackIndex is null)
            {
                // No owner to inherit from and no compatibility value supplied, so the
                // effective slot is genuinely unknown. Surfaced by TryBuild as
                // UnresolvedRenderLaneOwner.
                return null;
            }
        }

        Dictionary<int, List<LaneEventInput>> effectiveSlots = [];
        foreach (LaneEventInput input in rowOrdered)
        {
            // A supplied compatibility stack value is authoritative and is used as-is:
            // the caller assigned it, so no lookup is needed or wanted.
            var effective = input.EffectiveStackIndex ?? effectiveById[input.Event.Id];
            if (!effectiveSlots.TryGetValue(effective, out List<LaneEventInput>? slotEvents))
            {
                slotEvents = [];
                effectiveSlots.Add(effective, slotEvents);
            }

            slotEvents.Add(input);
        }

        int[] effectiveValues = [.. effectiveSlots.Keys.OrderBy(value => value)];

        // R4.7D / ADR-0026: the lane height is FIXED. The previous rule was
        // `laneHeight = max(LaneHeightPt, contentHeight)` -- "the lane grows;
        // events are never silently compressed" -- which is removed rather than
        // capped. A lane that grows disagrees with the Excel row it is meant to
        // sit in, and a projected child must not be able to grow the lane its
        // parent owns. LaneHeightPt is now the height, not a minimum.
        var contentHeight = metrics.LanePaddingTopPt + metrics.LanePaddingBottomPt;
        contentHeight += Math.Max(0, effectiveValues.Length - 1) * metrics.StackGapPt;
        contentHeight += effectiveValues.Sum(value => effectiveSlots[value].Max(input => input.ResolvedHeightPt));

        if (contentHeight > metrics.LaneHeightPt + GeometryMath.Epsilon)
        {
            // Reported, never accommodated. Growing the lane is the defect this
            // replaces; silently compressing the content would hide it instead.
            //
            // The owner is a synthetic row keyed by the LANE, not by one of the
            // lane's own rows. SceneValidator keys duplicate warnings by
            // (owner, code), so a row that owns several lanes would make this
            // code repeat for the same owner and SceneValidator would report a
            // DuplicateWarning finding against a correct scene. The identical
            // collision was hit and fixed for AmbiguousStackOverlap in R4.7B.
            warnings.Add(
                new SceneWarning(
                    SceneOwnerId.ForRow(GanttRowId.Parse("G-" + StableLaneWarningSuffix(laneKey))),
                    LaneContentExceedsLaneHeightCode,

                    // Formatted with InvariantCulture deliberately. A SceneWarning is
                    // part of the scene model, and the committed golden snapshot
                    // serialises these messages: under a locale whose decimal
                    // separator is a comma the same layout produced different text
                    // on a different machine. The numbers are diagnostics, but the
                    // snapshot that carries them must not depend on the host locale.
                    string.Create(
                        CultureInfo.InvariantCulture,
                        $"Lane content needs {contentHeight:0.##}pt but the row height is {metrics.LaneHeightPt:0.##}pt; the content is not compressed."))
            );
        }

        var laneHeight = metrics.LaneHeightPt;
        var slots = new List<SlotGeometry>();
        var slotTop = laneTop + metrics.LanePaddingTopPt;
        for (var visualIndex = 0; visualIndex < effectiveValues.Length; visualIndex++)
        {
            var effective = effectiveValues[visualIndex];
            List<LaneEventInput> slotEvents = effectiveSlots[effective];
            var height = slotEvents.Max(input => input.ResolvedHeightPt);
            GanttRowId[] eventIds =
            [
                .. slotEvents
                    .OrderBy(input => input.Event.SortOrder ?? int.MaxValue)
                    .ThenBy(input => input.Event.Id.Value, StringComparer.Ordinal)
                    .Select(input => input.Event.Id),
            ];
            slots.Add(new SlotGeometry(visualIndex, effective, slotTop, slotTop + (height / 2), slotTop + height, height, eventIds));
            AddAmbiguityWarning(slotEvents, warnings);
            slotTop += height + (visualIndex < effectiveValues.Length - 1 ? metrics.StackGapPt : 0);
        }

        GanttRowId[] laneEventIds =
        [
            .. inputs
                .OrderBy(input => input.Event.SortOrder ?? int.MaxValue)
                .ThenBy(input => input.Event.Id.Value, StringComparer.Ordinal)
                .Select(input => input.Event.Id),
        ];
        return new LaneGeometry(laneKey, laneOrder, laneTop, laneHeight, slots, laneEventIds, false, false);
    }

    /// <summary>
    /// Emits the same-stack overlap warning for a slot, once per slot.
    /// </summary>
    /// <remarks>
    /// <para>
    /// R4.7B: a projected child shares its parent's slot <b>by design</b>, and
    /// overlapping its parent is the intended presentation — REV6 §8 says several
    /// children on one parent lane "may overlap", distinguished by date geometry,
    /// Type, style, z-order and label. Emitting <c>AmbiguousStackOverlap</c> for
    /// that case would report the product working, and it would also collide with
    /// any genuine ambiguity on the same owner, because <c>SceneValidator</c> keys
    /// duplicate warnings by (owner, code) and a parent can own more than one
    /// slot.
    /// </para>
    /// <para>
    /// So a pair where one event is projected onto the other is not ambiguous and
    /// warns nothing; genuine same-slot overlap between unrelated rows still does.
    /// </para>
    /// </remarks>
    private static void AddAmbiguityWarning(List<LaneEventInput> slotEvents, List<SceneWarning> warnings)
    {
        for (var firstIndex = 0; firstIndex < slotEvents.Count; firstIndex++)
        {
            for (var secondIndex = firstIndex + 1; secondIndex < slotEvents.Count; secondIndex++)
            {
                if (IsProjectionPair(slotEvents[firstIndex], slotEvents[secondIndex]))
                {
                    continue;
                }

                if (Overlaps(slotEvents[firstIndex].Event, slotEvents[secondIndex].Event))
                {
                    warnings.Add(
                        new SceneWarning(
                            SceneOwnerId.ForRow(slotEvents[firstIndex].Event.Id),
                            _ambiguousStackOverlapCode,
                            "Events share an effective stack slot and overlap in time."
                        )
                    );
                    return;
                }
            }
        }
    }

    /// <summary>
    /// Whether the two events share a slot because the hierarchy put them there,
    /// rather than because placement happened to collide.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Two cases, both of which are the product working rather than an ambiguity:
    /// </para>
    /// <list type="number">
    /// <item>
    /// <description>
    /// One event is projected onto the lane the other owns — the owner/child pair,
    /// in either direction.
    /// </description>
    /// </item>
    /// <item>
    /// <description>
    /// <b>Both events are projected onto the same third row</b> — two children of
    /// one parent. This case was previously unguarded, so two time-overlapping
    /// children of one parent each emitted <c>AmbiguousStackOverlap</c> and warned
    /// that the product was broken. REV6 §8 states children on one parent lane "may
    /// overlap", distinguished by date geometry, Type, style, z-order and label;
    /// that is the intended presentation, exactly as the owner/child overlap above
    /// is. Two children of one parent are as unambiguous as a parent and its child.
    /// </description>
    /// </item>
    /// </list>
    /// <para>
    /// Genuine same-slot overlap between unrelated rows — neither projected, or
    /// projected onto different owners — still warns, because nothing in the
    /// hierarchy explains why they share a slot.
    /// </para>
    /// </remarks>
    private static bool IsProjectionPair(LaneEventInput first, LaneEventInput second)
    {
        GanttRowId? firstParent = first.RenderLaneOwner?.Id;
        GanttRowId? secondParent = second.RenderLaneOwner?.Id;

        // Siblings under one parent: both name a real owner and they agree.
        // Non-null on BOTH sides is required, so two unrelated top-level events
        // (neither projected, both null) do not take this branch.
        var shareOneOwner = firstParent is not null && firstParent == secondParent;

        return shareOneOwner
            || (firstParent == second.Event.Id)
            || (secondParent == first.Event.Id);
    }

    private static bool Overlaps(GanttEvent first, GanttEvent second)
    {
        if (first.Start is not { } firstStart || second.Start is not { } secondStart)
        {
            return false;
        }

        DateOnly firstFinish = first.Finish ?? firstStart;
        DateOnly secondFinish = second.Finish ?? secondStart;
        return firstStart <= secondFinish && secondStart <= firstFinish;
    }

    private static bool ValidMetrics(LaneLayoutMetrics metrics) =>
        IsFiniteNonNegative(metrics.LaneHeightPt)
        && metrics.LaneHeightPt > 0
        && IsFiniteNonNegative(metrics.LanePaddingTopPt)
        && IsFiniteNonNegative(metrics.LanePaddingBottomPt)
        && IsFiniteNonNegative(metrics.StackGapPt)
        && IsFiniteNonNegative(metrics.SplitterHeightPt)
        && IsFiniteNonNegative(metrics.SpacerHeightPt);

    private static bool IsFiniteNonNegative(double value) => double.IsFinite(value) && value >= 0;

    private static LaneLayoutCreationOutcome Refused(LaneLayoutRefusal refusal) => new(null, refusal);
}
