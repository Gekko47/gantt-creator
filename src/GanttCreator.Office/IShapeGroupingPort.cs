using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office;

/// <summary>
/// What the host reported after grouping the staged shape set (R6.4).
/// </summary>
/// <param name="GroupName">
/// The group shape's <c>Name</c> as the host reports it after the stamp. The grouping stage
/// asserts it equals <see cref="ShapeStagingTag.GroupName"/>; a different value means the
/// stamp did not land.
/// </param>
/// <param name="ChildNames">
/// The child shape names in the host's <c>GroupItems</c> order, 1..Count. This is the only
/// surface Excel exposes for child order, and the stage pins it against the scene order
/// (work-item D2) rather than accepting whatever the host produced.
/// </param>
/// <param name="Bounds">
/// The group's envelope in scene points, or <see langword="null"/> when the operation
/// refused or the host reported a non-finite member. The stage compares it against the
/// R6.3 staged bounds through <see cref="StagedEnvelopeCheck"/>.
/// </param>
/// <param name="Refusal">The port's refusal, or <see langword="null"/> on success.</param>
public sealed record ShapeGroupReadback(
    string? GroupName,
    IReadOnlyList<string> ChildNames,
    RectD? Bounds,
    ShapeWriteRefusal? Refusal)
{
    /// <summary>Gets whether the host produced and stamped a group.</summary>
    public bool Succeeded => Refusal is null;
}

/// <summary>
/// Narrow port over the one set-wide shape operation the editable export needs: grouping
/// the staged set into a single shape group, and ungrouping it again for rollback or
/// inspection (R6.4).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why a separate port.</b> <see cref="IShapeWritePort"/> is deliberately single-shape
/// operations by identifier; grouping is an export-staging-only, set-wide operation whose
/// only caller is the editable-export pipeline. Keeping it separate leaves the live
/// writer's contract and every existing fake untouched.
/// </para>
/// <para>
/// <b>Ownership is written by the adapter.</b> Exactly as <see cref="IShapeWritePort"/>
/// stamps the name and alternative text of every shape it creates, this port stamps the
/// group with <see cref="ShapeStagingTag.GroupName"/> and
/// <see cref="ShapeStagingTag.ForStagedName"/> of that name, plus <c>xlMove</c> placement
/// and the locked flag every created shape carries. The caller never composes either
/// member, so the group is a temporary staging shape in exactly the sense R6.5's lifecycle
/// registry requires (work-item D1 amendment).
/// </para>
/// <para>
/// <b>Preflight, then commit.</b> A staged name that cannot be resolved refuses
/// <see cref="ShapeWriteRefusal.NotFound"/> before the first mutation. Excel's
/// <c>Shapes.Range</c> raises for a missing name rather than returning an empty range, so
/// the preflight is a correctness requirement, not a nicety: without it a single absent
/// shape would fail the group mid-commit.
/// </para>
/// <para>
/// <b>Ungroup proves ownership by tag, not by name.</b> The same rule every other write
/// operation follows: a group whose alternative text is not the staging tag for its live
/// name is refused as <see cref="ShapeWriteRefusal.NotFound"/> and never touched, so a
/// user's drawing that happens to share the name survives.
/// </para>
/// </remarks>
public interface IShapeGroupingPort
{
    /// <summary>
    /// Groups the named staged shapes into one shape group, stamps the staging identity on
    /// it, and reports the child names and envelope the host holds.
    /// </summary>
    /// <param name="shapeNames">
    /// The staged shape names, in the scene's back-to-front creation order. This order is
    /// the expected child order; the host's reported order is not assumed to match it.
    /// </param>
    /// <returns>The host's read-back, or a typed refusal with nothing mutated.</returns>
    ShapeGroupReadback GroupStaged(IReadOnlyList<string> shapeNames);

    /// <summary>
    /// Ungroups one staging-owned group, restoring its children as top-level shapes.
    /// </summary>
    /// <param name="groupName">The group's live shape name.</param>
    /// <returns>
    /// The typed result. An absent group, or one that does not carry the staging tag for
    /// its live name, is <see cref="ShapeWriteRefusal.NotFound"/> and is never touched.
    /// </returns>
    /// <remarks>
    /// Used only for the grouping stage's rollback and the live gate's D3 inspection; it is
    /// deliberately not exposed through any command (work-item exclusion).
    /// </remarks>
    ShapeWriteOutcome UngroupStaged(string groupName);
}
