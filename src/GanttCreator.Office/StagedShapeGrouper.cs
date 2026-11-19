using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office;

/// <summary>Why the staged grouping stage produced no verified group (R6.4).</summary>
public enum StagedGroupRefusal
{
    /// <summary>No staged composition outcome was supplied.</summary>
    NullOutcome = 0,

    /// <summary>No grouping port was supplied.</summary>
    NullPort = 1,

    /// <summary>
    /// The staged composition itself refused, so there is no verified set to group. The
    /// composition's own refusal rides <see cref="StagingCompositionOutcome.Refusal"/>.
    /// </summary>
    StagedSetRefused = 2,

    /// <summary>
    /// The composition succeeded but staged nothing, and Excel cannot group an empty set.
    /// A supported empty composition is the caller's to interpret; grouping is not.
    /// </summary>
    NothingToGroup = 3,

    /// <summary>
    /// The host refused the grouping or the stamp. <see cref="StagedGroupOutcome.GroupRefusal"/>
    /// carries the port's reason, and the children were restored.
    /// </summary>
    HostRejected = 4,

    /// <summary>
    /// The host produced a group whose name is not <see cref="ShapeStagingTag.GroupName"/>,
    /// so the staging identity did not land (work-item D1).
    /// </summary>
    GroupIdentityMismatch = 5,

    /// <summary>
    /// The group's children are not exactly the staged set (work-item D1).
    /// <see cref="StagedGroupOutcome.ChildOrder"/> carries what the host reported.
    /// </summary>
    MembershipMismatch = 6,

    /// <summary>
    /// The group's child order is not the scene's creation order (work-item D2). The host
    /// re-ordered the children; the group was restored rather than accepted.
    /// </summary>
    ChildOrderMismatch = 7,

    /// <summary>
    /// The group's envelope does not equal the R6.3 staged bounds within the host's
    /// storage tolerance - the architecture's step-3 bounds check (work-item acceptance).
    /// </summary>
    BoundsMismatch = 8,
}

/// <summary>The typed result of grouping the staged export set (R6.4).</summary>
/// <param name="GroupName">The verified group's name, or <see langword="null"/> when none was produced.</param>
/// <param name="ChildOrder">
/// The child names the host reported. On success they are the verified scene order; on a
/// membership or order refusal they are the host's actual order, so the caller can name the
/// offending child rather than only reporting that the group disagreed.
/// </param>
/// <param name="Bounds">The group envelope the host reported, or <see langword="null"/>.</param>
/// <param name="Refusal">The stage refusal, or <see langword="null"/> on success.</param>
/// <param name="GroupRefusal">The port's refusal for <see cref="StagedGroupRefusal.HostRejected"/>.</param>
public sealed record StagedGroupOutcome(
    string? GroupName,
    IReadOnlyList<string> ChildOrder,
    RectD? Bounds,
    StagedGroupRefusal? Refusal,
    ShapeWriteRefusal? GroupRefusal)
{
    /// <summary>Gets whether the stage produced a verified group.</summary>
    public bool Succeeded => Refusal is null;
}

/// <summary>
/// R6.4: groups the staged editable-export set into one shape group with the scene's
/// deterministic child order, and verifies what the host produced before reporting success.
/// </summary>
/// <remarks>
/// <para>
/// <b>Scene-driven, never assumed (D1, D2).</b> The staged composition outcome already
/// carries the primitive identifiers in the scene's back-to-front order and the bounds the
/// set must occupy. The stage never re-derives either: it hands the port the expected names
/// in that order and then asserts what the host reports - membership exactly the staged
/// set, child order equal to the scene order, and the group envelope equal to the R6.3
/// bounds through <see cref="StagedEnvelopeCheck"/>, the same comparison the R6.3 envelope
/// pin uses.
/// </para>
/// <para>
/// <b>Host tolerance, not a widened scene tolerance.</b> The envelope compared here is a
/// host read-back, and Excel stores shape geometry in single precision (measured
/// 2026-11-18: 514.3622047pt requested, 514.3621826pt read back). The comparison therefore
/// uses <see cref="HostOwnedGeometry.GeometryMatchTolerancePt"/> - above float rounding,
/// below display resolution - exactly as the R6.3 live read-back does. The exact
/// <see cref="GeometryMath.Epsilon"/> equality remains pinned by the R6.3 composer tests on
/// translated requests.
/// </para>
/// <para>
/// <b>Refusal restores the sheet (D1, D3).</b> A group that fails any verification is
/// ungrouped before the refusal is returned, so a failed grouping leaves the staged set
/// exactly as the composer left it. This is the stage's own restore, not lifecycle
/// ownership: R6.5 owns deletion, and the group is registered into that lifecycle by the
/// staging identity the adapter stamps on it at creation.
/// </para>
/// <para>
/// <b>Nothing is exposed through a command yet.</b> R6.6 copies the group and R6.8 owns the
/// click; this row only makes the group exist, deterministically and verifiably.
/// </para>
/// </remarks>
public static class StagedShapeGrouper
{
    /// <summary>Groups and verifies the staged set.</summary>
    /// <param name="staged">
    /// The composition outcome produced by
    /// <see cref="EditablePanelStagingComposer.ComposeFullScene"/>.
    /// </param>
    /// <param name="port">The grouping port.</param>
    /// <returns>The verified group, or a typed refusal with the children restored.</returns>
    public static StagedGroupOutcome Group(
        StagingCompositionOutcome? staged,
        IShapeGroupingPort? port)
    {
        if (staged is null)
        {
            return Refused(StagedGroupRefusal.NullOutcome);
        }

        if (port is null)
        {
            return Refused(StagedGroupRefusal.NullPort);
        }

        if (!staged.Succeeded)
        {
            return Refused(StagedGroupRefusal.StagedSetRefused);
        }

        if (staged.StagedPrimitiveIds.Count == 0 || staged.StagedBounds is null)
        {
            return Refused(StagedGroupRefusal.NothingToGroup);
        }

        string[] expected =
        [
            .. staged.StagedPrimitiveIds.Select(id => ShapeStagingTag.NamePrefix + id),
        ];

        ShapeGroupReadback readback = port.GroupStaged(expected);
        if (!readback.Succeeded)
        {
            return new StagedGroupOutcome(
                null,
                [],
                null,
                StagedGroupRefusal.HostRejected,
                readback.Refusal);
        }

        // D1: the group itself is a temporary staging shape with the deterministic identity.
        if (!string.Equals(readback.GroupName, ShapeStagingTag.GroupName, StringComparison.Ordinal))
        {
            return Restore(port, readback, StagedGroupRefusal.GroupIdentityMismatch);
        }

        // D1: membership is every staged shape and nothing else.
        if (!HasExactMembership(readback.ChildNames, expected))
        {
            return Restore(port, readback, StagedGroupRefusal.MembershipMismatch);
        }

        // D2: child order mirrors the scene's creation order.
        if (!HasSceneOrder(readback.ChildNames, expected))
        {
            return Restore(port, readback, StagedGroupRefusal.ChildOrderMismatch);
        }

        // Step-3 bounds check, reused rather than restated.
        if (readback.Bounds is null)
        {
            return Restore(port, readback, StagedGroupRefusal.BoundsMismatch);
        }

        StagedEnvelopeResult check = StagedEnvelopeCheck.Evaluate(
            [readback.Bounds.Value],
            staged.StagedBounds.Value,
            HostOwnedGeometry.GeometryMatchTolerancePt);

        return check.Matches(staged.StagedBounds.Value)
            ? new StagedGroupOutcome(
                readback.GroupName,
                readback.ChildNames,
                readback.Bounds,
                null,
                null)
            : Restore(port, readback, StagedGroupRefusal.BoundsMismatch);
    }

    /// <summary>
    /// Whether the host's children are exactly the staged names - no more, no fewer, no
    /// duplicates.
    /// </summary>
    /// <param name="children">The child names the host reported.</param>
    /// <param name="expected">The expected staged names, in scene order.</param>
    /// <returns><see langword="true"/> when the sets are equal.</returns>
    private static bool HasExactMembership(IReadOnlyList<string> children, string[] expected) =>
        children.Count == expected.Length
        && new HashSet<string>(children, StringComparer.Ordinal).SetEquals(expected);

    /// <summary>Whether the host's children are in the scene's creation order.</summary>
    /// <param name="children">The child names the host reported.</param>
    /// <param name="expected">The expected staged names, in scene order.</param>
    /// <returns><see langword="true"/> when every position matches ordinally.</returns>
    private static bool HasSceneOrder(IReadOnlyList<string> children, string[] expected)
    {
        for (int index = 0; index < expected.Length; index++)
        {
            if (!string.Equals(children[index], expected[index], StringComparison.Ordinal))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Restores the pre-group state (the children become top-level staged shapes again) and
    /// returns the typed refusal, keeping the host's reported children for diagnostics.
    /// </summary>
    /// <param name="port">The grouping port.</param>
    /// <param name="readback">The host's read-back being rejected.</param>
    /// <param name="refusal">The verification that failed.</param>
    /// <returns>The refusal outcome.</returns>
    private static StagedGroupOutcome Restore(
        IShapeGroupingPort port,
        ShapeGroupReadback readback,
        StagedGroupRefusal refusal)
    {
        string? groupName = readback.GroupName;
        if (!string.IsNullOrWhiteSpace(groupName))
        {
            // CA1031: a cleanup failure must never replace the verification being reported.
#pragma warning disable CA1031
            try
            {
                _ = port.UngroupStaged(groupName);
            }
            catch (Exception)
            {
                // Intentionally empty: the verification refusal is the report.
            }
#pragma warning restore CA1031
        }

        return new StagedGroupOutcome(
            groupName,
            readback.ChildNames,
            readback.Bounds,
            refusal,
            null);
    }

    private static StagedGroupOutcome Refused(StagedGroupRefusal refusal) =>
        new(null, [], null, refusal, null);
}
