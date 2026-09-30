namespace GanttCreator.Core;

/// <summary>
/// The current version of the workbook schema contract owned by this
/// assembly: the <c>tblGanttData</c> column schema, the entity-type display
/// names, the entity-type contract metadata, and the exact
/// <c>tblGanttSettings</c> key/value contract.
/// </summary>
/// <remarks>
/// <para>
/// The schema version is stored on <c>tblGanttConfig</c> (R2.7) and checked
/// on initialise/open/Refresh. Bump <see cref="CurrentSchemaVersion"/> when
/// any of the following changes: a column name or column order, a catalogue
/// display name, catalogue contract metadata (date mode, kind, default style
/// key, colour capability, required-style flag, allowed label positions), or
/// the exact <c>tblGanttSettings</c> key/value contract. Renaming or
/// localising a display name is a workbook schema migration, not a change to
/// the durable <see cref="GanttEntityType"/> member name or numeric value.
/// </para>
/// </remarks>
public static class GanttSchemaVersion
{
    /// <summary>
    /// The current workbook schema version. Starts at 1 (R2.1), advances
    /// monotonically, and is currently 6 for the removal of the <c>Above</c> and
    /// <c>Below</c> label positions (owner ruling 2026-09-30, ADR-0028 D1).
    /// </summary>
    /// <remarks>
    /// <para>
    /// ADR-0029 D5 requires the bump to be the <em>last</em> change in the row's
    /// commit sequence: every other change lands first, so the integrity checker
    /// — not a manual review — is what notices a partial landing.
    /// </para>
    /// <para>
    /// Version 5 was R4.7C (the <c>Duration</c> column, column classification,
    /// <c>GanttRowHeightPt</c>, the retired <c>CriticalLinePt</c>, and
    /// <c>CriticalFill</c>). Allowed label positions are named in the version's own
    /// bump rule, so retiring two of them is a contract change and takes its own
    /// number rather than riding on 5.
    /// </para>
    /// </remarks>
    public const int CurrentSchemaVersion = 6;
}
