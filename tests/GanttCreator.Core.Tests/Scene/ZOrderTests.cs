using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

public sealed class ZOrderTests
{
    [Fact]
    public void Layers_match_the_entity_guide()
    {
        Assert.Equal(0, (int)ZLayer.Background);
        Assert.Equal(10, (int)ZLayer.AlternateBand);
        Assert.Equal(20, (int)ZLayer.Grid);
        Assert.Equal(25, (int)ZLayer.Delineator);
        Assert.Equal(30, (int)ZLayer.Section);
        Assert.Equal(40, (int)ZLayer.ActivityBody);
        Assert.Equal(50, (int)ZLayer.CriticalOverlay);
        Assert.Equal(60, (int)ZLayer.Milestone);
        Assert.Equal(70, (int)ZLayer.Label);
        Assert.Equal(75, (int)ZLayer.DelineatorLabel);
        Assert.Equal(80, (int)ZLayer.Frame);
        Assert.Equal(90, (int)ZLayer.Title);
    }

    [Fact]
    public void Activity_subtype_priority_matches_back_to_front_contract()
    {
        Assert.Equal(0, ActivitySubtypePriority.For(GanttEntityType.BaselineActivity));
        Assert.Equal(1, ActivitySubtypePriority.For(GanttEntityType.BaselineProcurement));
        Assert.Equal(2, ActivitySubtypePriority.For(GanttEntityType.AsPlannedActivity));
        Assert.Equal(3, ActivitySubtypePriority.For(GanttEntityType.AsPlannedProcurement));
        Assert.Equal(4, ActivitySubtypePriority.For(GanttEntityType.AsBuiltActivity));
        Assert.Equal(5, ActivitySubtypePriority.For(GanttEntityType.AsBuiltProcurement));
        Assert.Equal(6, ActivitySubtypePriority.For(GanttEntityType.CustomActivity));
        Assert.Equal(7, ActivitySubtypePriority.For(GanttEntityType.DelayEvent));
        Assert.Equal(int.MaxValue, ActivitySubtypePriority.For(GanttEntityType.AsBuiltMilestone));
    }

    [Fact]
    public void Scene_orders_layers_then_subtype_then_lane_stack_sort_and_id()
    {
        var owner = SceneOwnerId.ForRow(GanttRowId.New());
        SceneRect background = new("row:background", owner, ZLayer.Background, new RectD(0, 0, 10, 10), new SceneStyle("Default"));
        SceneRect delay = new(
            "row:delay",
            owner,
            ZLayer.ActivityBody,
            new RectD(0, 0, 10, 10),
            new SceneStyle("Delay"),
            GanttEntityType.DelayEvent,
            1,
            0,
            1
        );
        SceneRect actual = new(
            "row:actual",
            owner,
            ZLayer.ActivityBody,
            new RectD(0, 0, 10, 10),
            new SceneStyle("Actual"),
            GanttEntityType.AsBuiltActivity,
            1,
            0,
            1
        );
        SceneRect planned = new(
            "row:planned",
            owner,
            ZLayer.ActivityBody,
            new RectD(0, 0, 10, 10),
            new SceneStyle("Planned"),
            GanttEntityType.AsPlannedActivity,
            0,
            0,
            2
        );

        SceneCreationOutcome outcome = GanttScene.TryCreate(
            new RectD(0, 0, 100, 100),
            new RectD(10, 10, 90, 90),
            [delay, background, actual, planned],
            []
        );

        Assert.True(outcome.Succeeded);
        Assert.Equal(
            ["row:background", "row:planned", "row:actual", "row:delay"],
            outcome.Scene!.Primitives.Select(primitive => primitive.PrimitiveId)
        );
    }
}
