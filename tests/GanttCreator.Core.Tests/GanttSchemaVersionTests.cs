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
        Assert.Equal(10, GanttSchemaVersion.CurrentSchemaVersion);
    }

    [Fact]
    public void Current_schema_version_is_positive_and_invariant_culture_stable()
    {
        Assert.True(GanttSchemaVersion.CurrentSchemaVersion > 0);

        var text = GanttSchemaVersion.CurrentSchemaVersion.ToString(CultureInfo.InvariantCulture);
        Assert.Equal(GanttSchemaVersion.CurrentSchemaVersion, int.Parse(text, CultureInfo.InvariantCulture));
    }
}
