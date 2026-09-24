namespace GanttCreator.Core.Tests;

public class GanttEntityTypeTests
{
    [Theory]
    [InlineData(GanttEntityType.Splitter, 0)]
    [InlineData(GanttEntityType.Spacer, 1)]
    [InlineData(GanttEntityType.AsBuiltActivity, 2)]
    [InlineData(GanttEntityType.AsPlannedActivity, 3)]
    [InlineData(GanttEntityType.BaselineActivity, 4)]
    [InlineData(GanttEntityType.CriticalInterval, 5)]
    [InlineData(GanttEntityType.DelayEvent, 6)]
    [InlineData(GanttEntityType.AsBuiltProcurement, 7)]
    [InlineData(GanttEntityType.AsPlannedProcurement, 8)]
    [InlineData(GanttEntityType.BaselineProcurement, 9)]
    [InlineData(GanttEntityType.CustomActivity, 10)]
    [InlineData(GanttEntityType.AsBuiltMilestone, 11)]
    [InlineData(GanttEntityType.AsPlannedMilestone, 12)]
    [InlineData(GanttEntityType.BaselineMilestone, 13)]
    [InlineData(GanttEntityType.CriticalMilestone, 14)]
    [InlineData(GanttEntityType.Delineator, 15)]
    public void Machine_identity_name_and_value_are_stable(GanttEntityType type, int expectedValue)
    {
        Assert.Equal(expectedValue, (int)type);
        Assert.Equal(Enum.GetName(type), type.ToString());
    }
}
