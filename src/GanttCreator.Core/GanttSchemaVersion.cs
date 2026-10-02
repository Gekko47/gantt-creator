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
    /// monotonically, and is currently 10 for the reserved bottom padding row and
    /// the retirement of the external label-width cap (owner ruling, ADR-0035 D1/D2).
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>Version 10 is ADR-0035.</b> Two things move together, and both are
    /// contract changes rather than behaviour the add-in merely tolerates. (1) The
    /// metric catalogue loses <c>MaximumExternalLabelWidthPt</c>, so an existing
    /// workbook's <c>tblGanttMetrics</c> carries a row the running add-in no longer
    /// knows and its stored catalogue hash no longer matches. (2) The chart's
    /// bottom margin becomes a <em>reserved</em> worksheet row below the table
    /// rather than whatever row happened to follow it, so a workbook written by
    /// version 9 may have user content sitting in the row version 10 treats as
    /// the margin. As with versions 8 and 9 there is no migration (ADR-0029 D6):
    /// a version-9 workbook reports a mismatch and the remedy is Initialise, which
    /// is also what reserves the padding row.
    /// </para>
    /// <para>
    /// ADR-0029 D5 requires the bump to be the <em>last</em> change in the row's
    /// commit sequence: every other change lands first, so the integrity checker
    /// — not a manual review — is what notices a partial landing.
    /// </para>
    /// <para>
    /// <b>Version 10 is ADR-0035, and it changes what a bottom margin <em>is</em>.</b>
    /// Version 9 named the row below the last activity row as the chart's bottom
    /// padding but never reserved it, so the row was an ordinary worksheet row the
    /// user could type into and the add-in would silently resize it. Version 10
    /// reserves it: Initialise guarantees the row directly below the table exists
    /// and is empty, and both the row-height normaliser and the panel measurement
    /// refuse rather than write to a row that is not the reserved one. The metric
    /// catalogue simultaneously loses <c>MaximumExternalLabelWidthPt</c>, so a
    /// version-9 workbook also reports a catalogue-hash mismatch. There is no
    /// migration (ADR-0029 D6); the remedy is Initialise, which is what performs
    /// the reservation.
    /// </para>
    /// <para>
    /// <b>Version 9 adds the chart padding rows.</b> A row is reserved above the
    /// title row so the chart's top margin is a real worksheet row, and the row
    /// below the last activity row is its bottom counterpart (ADR-0031 D1). The
    /// header therefore moves from row 2 to row 3 and the plot anchor moves with it,
    /// which is the same class of change as version 8 and carries the same
    /// consequence: a version-8 workbook's header is still on row 2, so the anchor
    /// repair would write an address that points at the wrong row. There is no
    /// migration (ADR-0029 D6); the remedy is Initialise, which
    /// <c>SchemaVersionMismatchMessage</c> already tells the user to run.
    /// </para>
    /// <para>
    /// <b>Version 8 moves the table down one row.</b> The reserved row above
    /// <c>tblGanttData</c> carries the table title and the year band, so the header
    /// row moves from worksheet row 1 to row 2 and the plot anchor moves with it.
    /// A version-7 workbook reports a mismatch rather than being adjusted, because
    /// the anchor repair writes the <em>expected</em> address — on a version-7
    /// workbook that would point the anchor at row 2 while the header is still on
    /// row 1, producing a workbook that passes its own integrity check and is
    /// wrong. There is no migration (ADR-0029 D6); the remedy is Initialise, which
    /// <c>SchemaVersionMismatchMessage</c> already tells the user to run.
    /// </para>
    /// <para>
    /// Version 5 was R4.7C (the <c>Duration</c> column, column classification,
    /// <c>GanttRowHeightPt</c>, the retired <c>CriticalLinePt</c>, and
    /// <c>CriticalFill</c>). Allowed label positions are named in the version's own
    /// bump rule, so retiring two of them is a contract change and takes its own
    /// number rather than riding on 5.
    /// </para>
    /// <para>
    /// <b>Version 7 adds <c>SizePreset</c> and <c>RangePaddingDays</c> to
    /// <c>tblGanttSettings</c>.</b> Both were read by the scene-request factory from
    /// the first version of R4.8A but were never part of the approved key set, so
    /// <c>ValidateSettings</c> could not return them and both permanently fell back
    /// — the R4.7H size presets were unreachable in production while their own test
    /// passed on an injected dictionary. The exact key set is named in the bump rule
    /// above, which is why adding one is a schema change rather than a code detail.
    /// There is no migration: a workbook written against version 6 reports a version
    /// mismatch and is reported, never coerced (ADR-0029 D6).
    /// </para>
    /// </remarks>
    public const int CurrentSchemaVersion = 10;
}
