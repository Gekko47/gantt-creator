using GanttCreator.Core;

namespace GanttCreator.Core.Tests;

/// <summary>Contract tests for the canonical R2.8 scaffold rows.</summary>
public class GanttRowDefaultsTests
{
    public static TheoryData<GanttEntityType, string, string> ExpectedScaffolds => new()
    {
        { GanttEntityType.AsPlannedActivity, "As-Planned Activity", "AsPlannedActivity" },
        { GanttEntityType.AsPlannedMilestone, "As-Planned Milestone", "AsPlannedMilestone" },
        { GanttEntityType.Delineator, "Delineator", "DefaultDelineator" },
    };

    [Fact]
    public void Build_rejects_a_null_generator()
    {
        Assert.Throws<ArgumentNullException>(
            () => GanttRowDefaults.Build(GanttEntityType.AsPlannedActivity, null!));
    }

    [Fact]
    public void Build_rejects_a_type_outside_the_catalogue()
    {
        Assert.Throws<ArgumentOutOfRangeException>(
            () => GanttRowDefaults.Build((GanttEntityType)999, GanttRowId.New));
    }

    [Theory]
    [MemberData(nameof(ExpectedScaffolds))]
    public void Build_uses_the_exact_catalogue_scaffold(
        GanttEntityType type,
        string expectedType,
        string expectedStyleKey)
    {
        GanttRowId id = GanttRowId.Parse("G-0123456789abcdef0123456789abcdef");

        IReadOnlyList<object?> actual = GanttRowDefaults.Build(type, () => id);

        Assert.Equal(14, actual.Count);
        Assert.Equal(id.Value, actual[0]);
        Assert.Null(actual[1]);
        Assert.Null(actual[2]);
        Assert.Equal(expectedType, actual[3]);
        Assert.Null(actual[4]);
        Assert.Null(actual[5]);
        Assert.Null(actual[6]);
        Assert.Null(actual[7]);
        Assert.Equal(expectedStyleKey, actual[8]);
        Assert.Null(actual[9]);
        Assert.Null(actual[10]);
        Assert.Null(actual[11]);
        Assert.Null(actual[12]);
        Assert.Null(actual[13]);
    }
}
