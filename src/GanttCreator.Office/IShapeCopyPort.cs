namespace GanttCreator.Office;

/// <summary>
/// Narrow port over the one set-wide shape operation the editable export needs beyond
/// grouping: copying the verified staged group to the clipboard (R6.6).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why a separate port.</b> <see cref="IShapeWritePort"/> is single-shape operations by
/// identifier and <see cref="IShapeGroupingPort"/> is the set-wide grouping operation.
/// Copying the staged group is a third, export-staging-only operation whose only caller is
/// the editable-export pipeline. Keeping it separate leaves the live writer's and the
/// grouper's contracts untouched.
/// </para>
/// <para>
/// <b>Ownership is proved by recomputation, never by the name.</b> The adapter resolves the
/// group's worksheet through the shared <see cref="StagingShapeAdapter"/> target resolution
/// and accepts the shape only when its alternative text equals
/// <see cref="ShapeStagingTag.ForStagedName"/> recomputed from its live name. A user's
/// shape that happens to share the name is refused <see cref="ShapeWriteRefusal.NotFound"/>
/// and never touched.
/// </para>
/// <para>
/// <b>The copy has no worksheet effect.</b> <c>Shape.Copy</c> writes the clipboard, never
/// the workbook; the adapter nevertheless queries the protection guard first, matching the
/// staging family's guard-first discipline (ADR-0008 D4), so a protected target refuses
/// before the operation rather than after it.
/// </para>
/// </remarks>
public interface IShapeCopyPort
{
    /// <summary>
    /// Copies one staging-owned group shape to the clipboard.
    /// </summary>
    /// <param name="groupName">
    /// The group's live shape name (<see cref="ShapeStagingTag.GroupName"/> after a verified
    /// grouping).
    /// </param>
    /// <returns>
    /// The typed result. An absent group, or one that does not carry the staging tag for its
    /// live name, is <see cref="ShapeWriteRefusal.NotFound"/> and is never touched.
    /// </returns>
    ShapeWriteOutcome CopyStaged(string groupName);
}
