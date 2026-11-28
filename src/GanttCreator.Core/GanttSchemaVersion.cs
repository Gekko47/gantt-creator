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
    /// monotonically, and is currently 30 for the plot-grid retirement.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>Version 30 retires the plot grid and its controls (R11.1, owner ruling 2026-11-27).</b>
    /// Two tables move. <c>tblGanttSettings</c> loses <c>ShowMinorGrid</c>,
    /// <c>ShowMajorGrid</c> and <c>MinorGridColour</c>, so a version-29 workbook's settings
    /// table carries three rows the running add-in no longer knows. <c>tblGanttMetrics</c>
    /// loses <c>GridLinePt</c>, so it carries one row the add-in no longer reads.
    /// <c>AdditiveSchemaUpgrade</c> inspects settings keys only and refuses a stored list
    /// longer than the approved one as <c>ExtraKey</c>, so the workbook reports the
    /// confirmation-required <c>SchemaVersionOlder</c>; the metrics removal reports as a
    /// <c>CatalogueHashMismatch</c>. Both are repaired by Repair, which rewrites the
    /// catalogues from the code-owned definitions (ADR-0029 D6: no migration).
    /// </para>
    /// <para>
    /// <b>Version 29 is R8.9's custom export canvas.</b> <c>tblGanttSettings</c>
    /// gains <c>SizeWidthCm</c> and <c>SizeHeightCm</c>, appended after the existing
    /// keys so a version-28 workbook's key list is an exact prefix.
    /// </para>
    /// <para>
    /// <b>Version 28 is R8.8's PNG-export settings.</b> <c>tblGanttSettings</c> gains
    /// <c>PngExportTarget</c> and <c>LastPngExportFolder</c>, so a version-27 workbook's
    /// settings table lacks two rows the running add-in reads and its stored catalogue hash
    /// no longer matches. There is no migration (ADR-0029 D6): a version-27 workbook reports
    /// a mismatch and the remedy is Repair, which rewrites the catalogues from the code-owned
    /// definitions and stamps the new version and hash.
    /// </para>
    /// <para>
    /// <b>Version 27 is the period-label reference width (ADR-0041 amendment).</b>
    /// The metric catalogue gains <c>PeriodLabelReferencePlotWidthPt</c>, so a
    /// version-26 workbook's <c>tblGanttMetrics</c> lacks a row the running add-in
    /// reads and its stored catalogue hash no longer matches. There is no
    /// migration (ADR-0029 D6): a version-26 workbook reports a mismatch and the
    /// remedy is <b>Repair</b>, which rewrites the catalogues from the code-owned
    /// definitions and stamps the new version and hash. Repair is the remedy
    /// rather than Initialise because the ribbon no longer carries an Initialise
    /// command, and a version bump that could only be cleared by a command the
    /// user cannot reach would strand every existing workbook.
    /// </para>
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
    /// <para>
    /// <b>Version 20 is R5.6a's chart-chrome colours.</b> The
    /// <c>tblGanttSettings</c> key set gains six chart-chrome colour keys
    /// (<c>ChartBackgroundColour</c>, <c>AlternateBandColour</c>,
    /// <c>MinorGridColour</c>, <c>MajorGridColour</c>,
    /// <c>YearHeaderColour</c>, and <c>PeriodHeaderColour</c>), so a
    /// version-19 workbook's settings table lacks six rows the running add-in
    /// reads and its stored catalogue hash no longer matches. The colours are
    /// the chart's own frame — the background, the alternating bands, the
    /// grid, and the header bands — rather than any entity's style, and each
    /// default is the approved colour-token value, so the catalogue and the
    /// token table cannot disagree. The table's extent grew from 19 keys to
    /// 25, so the config table's anchor moved from A92 to A98 to keep a blank
    /// row between the two (the 2026-10-09 defect). There is no migration
    /// (ADR-0029 D6): a version-19 workbook reports a mismatch and the remedy
    /// is Initialise.
    /// </para>
    /// <para>
    /// <b>Version 19 is R5.6a's stroke dash type.</b> The
    /// <c>tblGanttStyles</c> column schema gains the engine-owned
    /// <c>StrokeDash</c> column (the style's line dash pattern:
    /// <c>Solid</c>, <c>SquareDot</c>, <c>RoundDot</c>, <c>Dash</c>,
    /// <c>DashDot</c>, or <c>LongDash</c>), so a version-18 workbook's style
    /// table lacks a column the running add-in reads and its stored catalogue
    /// hash no longer matches. The column is appended at the end of the row, so
    /// every existing column index is unchanged. There is no migration
    /// (ADR-0029 D6): a version-18 workbook reports a mismatch and the remedy
    /// is Initialise.
    /// </para>
    /// <para>
    /// <b>Version 26 is R5.21's user-authored <c>Activity ID</c> column (ADR-0064).</b>
    /// <c>tblGanttData</c> gains a required, user-visible <c>Activity ID</c> column
    /// directly after <c>Type</c>, so the analyst has a place to record their own
    /// reference for a row without hand-editing the hidden, locked engine <c>Id</c>.
    /// The column is classified <c>Authoring</c>, is measured as part of the data
    /// panel, and is offered by the ribbon's column dropdown as
    /// <c>colActivity ID</c>. No catalogue row, style, or setting changes, so the
    /// settings table's extent and the config anchor do not move; the hash moves
    /// through the schema seed alone. The column count moves 17 → 18, so
    /// <see cref="GanttSheetLayout.PlotAnchorColumnIndex"/> moves from column S to
    /// column T. A version-25 workbook's table lacks the column, so its stored
    /// catalogue hash no longer matches and it reports a mismatch; the remedy is
    /// Initialise (ADR-0029 D6, no migration).
    /// </para>
    /// <para>
    /// <b>Version 25 is R5.20's persisted automatic-refresh switch.</b>
    /// <c>tblGanttSettings</c> gains <c>AutoRefreshEnabled</c>, which records whether a
    /// chart-affecting table edit or a successful chart-setting command may rebuild the
    /// chart automatically. The key is required rather than session state because the
    /// user switches the automatic refresh off to work without COM churn, and a reopening
    /// that switched it back on would silently resume the passes they turned off. The
    /// stored value is read fail-open: absent or unreadable means enabled, which is the
    /// behaviour every earlier release had. The extra row moves the settings table's extent
    /// from row 97 to row 98, so the config table's anchor moves from A99 to A100 to keep
    /// the blank row Excel's table auto-expand requires. No table column, order or
    /// entity-style value changes. A version-24 workbook's settings table lacks the row, so
    /// its stored catalogue hash no longer matches and it reports a mismatch; the remedy is
    /// Initialise (ADR-0029 D6, no migration).
    /// </para>
    /// <para>
    /// <b>Version 24 is R5.18h's persisted column visibility.</b>
    /// <c>tblGanttSettings</c> gains <c>HiddenColumns</c>, which stores the schema display
    /// names of the table columns the user hid from the ribbon's column dropdown. The key
    /// is required rather than optional because the choice must survive opening the
    /// workbook in another instance, and a missing row would make every reopening look
    /// like a fresh workbook and un-hide the user's columns. The extra row moves the
    /// settings table's extent from row 95 to row 97, so the config table's anchor moves
    /// from A98 to A99 to keep the blank row Excel's table auto-expand requires. A
    /// version-23 workbook's settings table lacks the row, so its stored catalogue hash no
    /// longer matches and it reports a mismatch; the remedy is Initialise (ADR-0029 D6, no
    /// migration).
    /// </para>
    /// <para>
    /// <b>Version 23 is R5.18k's visible-columns-first table (owner-approved).</b>
    /// <c>tblGanttData</c> now leads with <c>Type</c>, <c>Description</c>, <c>Start</c>,
    /// <c>Finish</c> and <c>Duration</c>, and the engine columns follow them instead of
    /// preceding them. The table also starts in column B so column A can hold the
    /// add-in's spacer gutter, which is not a table column. Both changes are
    /// schema-visible: the column order is part of the schema, and the origin moves
    /// <see cref="GanttSheetLayout.PlotAnchorColumnIndex"/> with it. There is no
    /// migration (ADR-0029 D6): a version-22 workbook reports a mismatch and the remedy
    /// is Initialise, which is the only creation path and has no active users to migrate.
    /// </para>
    /// <para>
    /// <b>Version 22 is ADR-0055's add-in-owned worksheet protection.</b>
    /// <c>tblGanttSettings</c> gains <c>ProtectGanttSheet</c>, which records
    /// that the add-in owns the protection of the visible Gantt worksheet and
    /// doubles as the user's switch for the feature. The key is required
    /// because the guard refuses every write to a protected worksheet
    /// (ADR-0008 D4): without a record of ownership, protecting the sheet to
    /// lock the chart shapes and refreshing it would be mutually exclusive. A
    /// version-21 workbook's settings table lacks the row, so its stored
    /// catalogue hash no longer matches and it reports a mismatch; the remedy
    /// is Initialise (ADR-0029 D6, no migration).
    /// </para>
    /// </remarks>
    public const int CurrentSchemaVersion = 30;
}
