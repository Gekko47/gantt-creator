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
    /// monotonically, and is currently 17 for ADR-0043's live title as an owned
    /// textbox shape.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>Version 17 is ADR-0043.</b> The live chart title moves from a cell
    /// placeholder in the reserved row to an owned textbox shape created by
    /// Refresh. The cell-placeholder write is removed; the reserved row no
    /// longer carries title text. A version-16 workbook's stored catalogue
    /// hash no longer matches because the initialiser no longer writes title
    /// text — there is no migration (ADR-0029 D6): a version-16 workbook
    /// reports a mismatch and the remedy is Initialise.
    /// </para>
    /// <para>
    /// <b>Version 16 is R5.6a.</b> The type catalogue gains
    /// <c>CustomMilestone</c> (named custom milestone; the row must supply a
    /// <c>StyleKey</c>), so a version-15 workbook's <c>tblGanttTypes</c> lacks
    /// a row the running add-in reads and its stored catalogue hash no longer
    /// matches. There is no migration (ADR-0029 D6): a version-15 workbook
    /// reports a mismatch and the remedy is Initialise.
    /// </para>
    /// <para>
    /// <b>Version 15 is R5.14.</b> The <c>tblGanttData</c> column schema gains
    /// the engine-hidden <c>DateLabelPosition</c> column (the per-row date-label
    /// selector: blank/<c>Auto</c>, <c>Start</c>, <c>Finish</c>, or
    /// <c>None</c>), so a
    /// version-14 workbook's table lacks a column the running add-in reads and its
    /// stored schema version no longer matches. The date-label capability is
    /// derived from the existing <c>EntityDateMode</c> catalogue metadata, so no
    /// catalogue row or stored hash changes — only the column. There is no
    /// migration (ADR-0029 D6): a version-14 workbook reports a mismatch and the
    /// remedy is Initialise, which materialises the column.
    /// </para>
    /// <para>
    /// <b>Version 14 is R5.2.</b> The settings catalogue gains <c>Margin</c> and
    /// <c>MarginCm</c>, so a version-13 workbook's <c>tblGanttSettings</c> lacks
    /// two rows the running add-in reads and its stored catalogue hash no longer
    /// matches. There is no migration (ADR-0029 D6): a version-13 workbook reports
    /// a mismatch and the remedy is Initialise.
    /// </para>
    /// <para>
    /// <b>Version 13 is R5.1.</b> The settings catalogue gains
    /// <c>PlotStartDate</c> and <c>PlotFinishDate</c>, so a version-12
    /// workbook's <c>tblGanttSettings</c> lacks two rows the running add-in
    /// reads and its stored catalogue hash no longer matches. There is no
    /// migration (ADR-0029 D6): a version-12 workbook reports a mismatch and
    /// the remedy is Initialise.
    /// </para>
    /// <para>
    /// <b>Version 12 is ADR-0038.</b> Two things move together, and neither is a
    /// behaviour the add-in merely tolerates. (1) The metric catalogue gains
    /// <c>ChartAnchorRowHeightPt</c>, so a version-11 workbook's
    /// <c>tblGanttMetrics</c> lacks a row the running add-in reads and its stored
    /// catalogue hash no longer matches. (2) A NEW RESERVED ROW appears between the
    /// last body row and the bottom padding row, so the padding row the add-in writes
    /// to and measures as the chart's bottom margin is no longer the row directly
    /// below the table. Both are the same class of change as version 9's and
    /// version 10's, which moved the row the padding lives in. There is no migration
    /// (ADR-0029 D6): a version-11 workbook reports a mismatch and the remedy is
    /// Initialise, which is also what reserves the anchor row.
    /// </para>
    /// <para>
    /// <b>Version 11 is ADR-0037.</b> The metric catalogue gains
    /// <c>PlotBandHeaderOverlapPt</c>, so an existing workbook's
    /// <c>tblGanttMetrics</c> lacks a row the running add-in reads and its stored
    /// catalogue hash no longer matches -- the same class of change as version 10's
    /// retirement, in the opposite direction. The value is what the plot-spanning
    /// shapes use as their sub-row overlap into the header row, so it is a stored
    /// measurement rather than a display name. There is no migration (ADR-0029 D6): a
    /// version-10 workbook reports a mismatch and the remedy is Initialise.
    /// </para>
    /// <para>
    /// <b>Version 10 is ADR-0035.</b> Two things move together, and both are
    /// contract changes rather than behaviour the add-in merely tolerates. (1) The
    /// metric catalogue loses <c>MaximumExternalLabelWidthPt</c>, so an existing
    /// workbook's <c>tblGanttMetrics</c> carries a row the running add-in no longer
    /// knows and its stored catalogue hash no longer matches. (2) The chart's
    /// bottom margin becomes a <em>reserved</em> worksheet row below the table
    /// rather than whatever row happened to follow it — a change to what a bottom
    /// margin <em>is</em>. Version 9 merely *named* the row below the last activity
    /// row without reserving it, so it stayed an ordinary worksheet row the user
    /// could type into and the add-in would silently resize it. Version 10 reserves
    /// it: Initialise guarantees the row directly below the table exists and is
    /// empty, and both the row-height normaliser and the panel measurement refuse
    /// rather than write to a row that is not the reserved one. So a workbook
    /// written by version 9 may have user content sitting in the row version 10
    /// treats as the margin. As with versions 8 and 9 there is no migration
    /// (ADR-0029 D6): a version-9 workbook reports a mismatch and the remedy is
    /// Initialise, which is also what reserves the padding row.
    /// </para>
    /// <para>
    /// ADR-0029 D5 requires the bump to be the <em>last</em> change in the row's
    /// commit sequence: every other change lands first, so the integrity checker
    /// — not a manual review — is what notices a partial landing.
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
    /// <para>
    /// <b>Version 18 is R5.6a's custom-style fallback.</b> The two style-key
    /// types carry a code-owned default style key (<c>CustomActivity</c> and
    /// <c>CustomMilestone</c>) and a matching built-in preset, so a row whose
    /// user-created style was deleted resolves to a renderable style instead
    /// of refusing the whole scene. Their allowed label positions become the
    /// span and milestone sets, and the built-in presets' default label
    /// position becomes <c>Auto</c>, so a healed row renders a label rather
    /// than a bare bar. The catalogue contents move through the type rows and
    /// the style rows; the hash moves with them. A version-17 workbook carries
    /// the older hash and is reported as a mismatch; the remedy is Initialise
    /// (ADR-0029 D6, no migration).
    /// </para>
    /// </remarks>
    public const int CurrentSchemaVersion = 18;
}
