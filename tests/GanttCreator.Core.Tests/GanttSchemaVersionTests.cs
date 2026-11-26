using System.Globalization;

namespace GanttCreator.Core.Tests;

public class GanttSchemaVersionTests
{
    /// <summary>
    /// The schema version is the current one, and each bump is named with the row
    /// that took it.
    /// </summary>
    /// <remarks>
    /// The name of this test used to say "version six" while asserting 7, which is
    /// the kind of stale title that survives because nothing checks it. It now
    /// describes what it pins rather than repeating a number that lives in the
    /// comment below.
    /// </remarks>
    [Fact]
    public void Current_schema_version_is_the_current_version()
    {
        // Schema v2 added the nonblank title/default and PeriodLabelFormat contract (ADR-0014).
        // Schema v3 adds the DateDisplayFormat contract (ADR-0016).
        // Schema v4 adds the SiblingOrder column (R4.7A D2). ADR-0029 D5 splits the
        // version across two rows on purpose -- R4.7A took 3 -> 4 and R4.7C took
        // 4 -> 5 -- because two rows claiming one version number would collide or
        // skip a version in a way ConfigIntegrity cannot detect, since the version
        // is the signal it compares against.
        // Schema v5 adds the Duration column, the column classification, the
        // GanttRowHeightPt rename, the retired CriticalLinePt, and CriticalFill.
        // Schema v8 is R4.7I: the reserved title row moves the header from row 1 to
        // row 2 and the plot anchor with it (ADR-0030 D4/D5).
        // Schema v9 is ADR-0031 D1: the chart's top and bottom padding become real
        // worksheet rows, so the header moves from row 2 to row 3 and the plot anchor
        // moves again. It also adds the ChartPaddingRowHeightPt token.
        // Schema v10 is ADR-0035: the BOTTOM padding row becomes a genuinely
        // reserved row rather than whatever row followed the table, and the metric
        // catalogue loses MaximumExternalLabelWidthPt. No header move this time --
        // the reservation is below the table -- but the stored catalogue hash moves,
        // so a version-9 workbook reports a mismatch and Initialise is the remedy.
        // Schema v11 is ADR-0037: the metric catalogue GAINS PlotBandHeaderOverlapPt,
        // the sub-row amount the plot-spanning shapes extend up into the header row so
        // Excel's cell anchoring stretches them when a row is added at the top. Same
        // class as v10's retirement, opposite direction: the stored catalogue hash
        // moves because a workbook written before it lacks a row the add-in reads.
        // Schema v12 is ADR-0038: the metric catalogue GAINS ChartAnchorRowHeightPt AND a
        // NEW RESERVED ROW appears between the body and the bottom padding row, so the
        // padding row the add-in writes to and measures is no longer the row directly
        // below the table. Same class as v9's and v10's, both of which moved the row
        // the padding lives in. No migration (ADR-0029 D6); Initialise is the remedy
        // and is also what reserves the anchor row.
        // Schema v13 is R5.1: the settings catalogue gains PlotStartDate and
        // PlotFinishDate for the explicit plot-range modes.
        // Schema v14 is R5.2: the settings catalogue gains Margin and MarginCm for the
        // plot page margin. The margin is subtracted twice from the width budget to
        // determine the usable plotting area; it never touches the measured text panel.
        // Schema v15 is R5.14: the tblGanttData column schema gains the engine-hidden
        // DateLabelPosition column (the per-row date-label selector). The capability is
        // derived from the existing EntityDateMode metadata, so no catalogue row or hash
        // changes -- only the column.
        // Schema v17 is ADR-0043: the live chart title moves from a cell
        // placeholder in the reserved row (written by Initialise) to an owned
        // textbox shape created by Refresh. WriteTitleRow, RollBackTitleRow,
        // and GetTitleRange are removed; the reserved row no longer carries
        // title text. A version-16 workbook's stored catalogue hash no longer
        // matches because the initialiser no longer writes title text. No
        // migration (ADR-0029 D6); Initialise is the remedy.
        // Schema v18 is R5.6a: the two style-key types (Custom Activity,
        // Custom Milestone) carry a code-owned default style key and a matching
        // built-in preset, so a row whose user-created style was deleted
        // resolves to a renderable style instead of refusing the scene. Their
        // allowed label positions become the span and milestone sets, and the
        // built-in presets' default label position becomes Auto. The catalogue
        // hash moves with the type and style rows. No migration (ADR-0029 D6);
        // Initialise is the remedy.
        // Schema v19 is R5.6a: the tblGanttStyles column schema gains the
        // StrokeDash column (the style's line dash pattern). The column is
        // appended at the end of the row, so every existing column index is
        // unchanged. No migration (ADR-0029 D6); Initialise is the remedy.
        // 22 adds `ProtectGanttSheet` (ADR-0055).
        // 23 reorders tblGanttData to lead with the visible columns and moves the
        // table to column B so column A can hold the spacer gutter (R5.18k).
        // 24 adds `HiddenColumns` to tblGanttSettings so the R5.18h column-
        // visibility choice survives reopening the workbook. A blank value means
        // nothing is hidden; a value that does not parse is refused, never
        // defaulted. No migration (ADR-0029 D6); Initialise is the remedy.
        // 25 adds `AutoRefreshEnabled` to tblGanttSettings so the R5.20 switch
        // survives reopening the workbook. It is read fail-open (absent or
        // unreadable means enabled), and the config table's anchor moves from
        // A99 to A100 to keep the blank row between the tables.
        // 26 adds the user-authored `Activity ID` column directly after Type
        // (R5.21, ADR-0064). It is visible and editable, is measured into the data
        // panel, and moves the plot anchor from column S to column T because the
        // anchor is the origin plus the column count. No migration (ADR-0029 D6);
        // Repair is the remedy.
        // 27 adds the PeriodLabelReferencePlotWidthPt metric token (ADR-0041's
        // 2026-11-20 amendment). No migration (ADR-0029 D6); Repair is the remedy,
        // because the ribbon no longer carries an Initialise command.
        // 28 adds `PngExportTarget` and `LastPngExportFolder` to tblGanttSettings
        // (R8.8), so the PNG export target and the remembered folder survive
        // reopening the workbook. The config anchor moves from A100 to A102 to keep
        // the blank row. No migration (ADR-0029 D6); Repair is the remedy.
        // 29 adds `SizeWidthCm` and `SizeHeightCm` (R8.9), appended after the
        // existing keys so a version-28 workbook's list is an exact prefix. That
        // prefix is what lets the additive repair upgrade it in place (ADR-0073)
        // without Initialise and without rewriting tblGanttMetrics.
        Assert.Equal(29, GanttSchemaVersion.CurrentSchemaVersion);
    }

    [Fact]
    public void Current_schema_version_is_positive_and_invariant_culture_stable()
    {
        Assert.True(GanttSchemaVersion.CurrentSchemaVersion > 0);

        var text = GanttSchemaVersion.CurrentSchemaVersion.ToString(CultureInfo.InvariantCulture);
        Assert.Equal(GanttSchemaVersion.CurrentSchemaVersion, int.Parse(text, CultureInfo.InvariantCulture));
    }
}
