using System.Globalization;

namespace GanttCreator.Core.Tests;

public class GanttSchemaVersionTests
{
    [Fact]
    public void Current_schema_version_is_the_current_version_two()
    {
        // Schema v2 adds the nonblank title/default and PeriodLabelFormat contract (ADR-0014).
        Assert.Equal(2, GanttSchemaVersion.CurrentSchemaVersion);
    }

    [Fact]
    public void Current_schema_version_is_positive_and_invariant_culture_stable()
    {
        Assert.True(GanttSchemaVersion.CurrentSchemaVersion > 0);

        var text = GanttSchemaVersion.CurrentSchemaVersion.ToString(CultureInfo.InvariantCulture);
        Assert.Equal(GanttSchemaVersion.CurrentSchemaVersion, int.Parse(text, CultureInfo.InvariantCulture));
    }
}
