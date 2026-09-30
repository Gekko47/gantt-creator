using GanttCreator.Core;

namespace GanttCreator.Office;

/// <summary>Why the worksheet outline could not be rebuilt.</summary>
public enum OutlineGroupRefusalReason
{
    /// <summary>The application object or active workbook was absent.</summary>
    NoActiveWorkbook = 0,

    /// <summary>The Gantt worksheet or its <c>tblGanttData</c> table was missing.</summary>
    TableMissing = 1,

    /// <summary>The target worksheet is protected.</summary>
    TargetProtected = 2,
}

/// <summary>The typed result of rebuilding the worksheet outline.</summary>
/// <param name="GroupsApplied">How many outline groups were created.</param>
/// <param name="RowsUngrouped">How many stale groups were cleared.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/> on success.</param>
public sealed record OutlineGroupOutcome(int GroupsApplied, int RowsUngrouped, OutlineGroupRefusalReason? Refusal)
{
    /// <summary>Gets whether the outline was rebuilt.</summary>
    public bool Succeeded => Refusal is null;

    /// <summary>Creates a successful outcome.</summary>
    /// <param name="groupsApplied">How many groups were created.</param>
    /// <param name="rowsUngrouped">How many rows were ungrouped.</param>
    /// <returns>The successful outcome.</returns>
    public static OutlineGroupOutcome Ok(int groupsApplied, int rowsUngrouped) =>
        new(groupsApplied, rowsUngrouped, null);

    /// <summary>Creates a refusal.</summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <returns>The refusal outcome.</returns>
    public static OutlineGroupOutcome Refused(OutlineGroupRefusalReason refusal) => new(0, 0, refusal);
}

/// <summary>
/// Rebuilds the native Excel outline of the Gantt table from the validated
/// hierarchy, so the <c>-</c>/<c>+</c> controls are a presentation of
/// <c>ParentId</c> and never a second source of truth (R4.7D, ADR-0026 D5).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why the outline is rebuilt rather than read.</b> The hierarchy lives in the
/// hidden <c>ParentId</c> column. Reading the outline back would make collapsing a
/// group a way to change the model, and expanding one a way to silently redefine
/// the tree. The caller supplies the validated events; this adapter applies the
/// plan <see cref="OutlineGroupPlanner"/> derives from them.
/// </para>
/// <para>
/// <b>Why stale groups are cleared first.</b> Excel refuses to regroup a sheet
/// whose outline disagrees with the requested grouping, so
/// <see cref="OutlineGroupPlan.RowsToUngroup"/> is applied before any group is
/// created. The two lists never overlap, so the order of the two phases cannot
/// change the result.
/// </para>
/// <para>
/// <b>Collapse state is preserved.</b> Only the grouping is written. Whether a group
/// is currently expanded or collapsed is the user's presentation state and is never
/// touched, so a Refresh does not re-expand a collapsed parent.
/// </para>
/// <para>
/// <b>COM ownership.</b> The proxies reached here are Excel-owned shared roots; this
/// adapter takes no ownership and never calls <c>FinalReleaseComObject</c>. The
/// <c>internal virtual</c> seams isolate the COM parameterised properties for
/// contract tests.
/// </para>
/// </remarks>
public interface IOutlineGroupPort
{
    /// <summary>Rebuilds the table's outline groups from a validated hierarchy.</summary>
    /// <param name="events">The validated events, from <c>GanttRowValidator</c>.</param>
    /// <returns>How many groups were applied, or the refusal reason.</returns>
    OutlineGroupOutcome Apply(IReadOnlyList<GanttEvent> events);
}
