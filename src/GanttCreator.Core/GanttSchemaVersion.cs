namespace GanttCreator.Core;

/// <summary>
/// The current version of the workbook schema contract owned by this
/// assembly: the <c>tblGanttData</c> column schema, the entity-type display
/// names, and the entity-type contract metadata.
/// </summary>
/// <remarks>
/// <para>
/// The schema version is stored on <c>tblGanttConfig</c> (R2.7) and checked
/// on initialise/open/Refresh. Bump <see cref="CurrentSchemaVersion"/> when
/// any of the following changes: a column name or column order, a catalogue
/// display name, or catalogue contract metadata (date mode, kind, default
/// style key, colour capability, required-style flag, allowed label
/// positions). Renaming or localising a display name is a workbook schema
/// migration, not a change to the durable <see cref="GanttEntityType"/>
/// member name or numeric value.
/// </para>
/// </remarks>
public static class GanttSchemaVersion
{
    /// <summary>
    /// The current workbook schema version. Starts at 1 (R2.1) and increases
    /// monotonically; never reused, never reset.
    /// </summary>
    public const int CurrentSchemaVersion = 1;
}
