using System.Globalization;

namespace GanttCreator.Core.Tests;

public class GanttSchemaVersionTests
{
    [Fact]
    public void Current_schema_version_is_the_current_version_five()
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
        Assert.Equal(5, GanttSchemaVersion.CurrentSchemaVersion);
    }

    [Fact]
    public void Current_schema_version_is_positive_and_invariant_culture_stable()
    {
        Assert.True(GanttSchemaVersion.CurrentSchemaVersion > 0);

        var text = GanttSchemaVersion.CurrentSchemaVersion.ToString(CultureInfo.InvariantCulture);
        Assert.Equal(GanttSchemaVersion.CurrentSchemaVersion, int.Parse(text, CultureInfo.InvariantCulture));
    }
}
