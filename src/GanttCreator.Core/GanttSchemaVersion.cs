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
    /// monotonically, and is currently 4 for R4.7A's <c>SiblingOrder</c> column
    /// (ADR-0029 D5, step 1). R4.7C takes it to 5 for the <c>Duration</c> column
    /// and the token/label-set changes; each schema-changing row takes its own bump,
    /// so two rows can never claim one version number.
    /// </summary>
    public const int CurrentSchemaVersion = 4;
}
