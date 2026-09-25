using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

public sealed class GanttSceneTests
{
    [Fact]
    public void TryCreate_orders_primitives_and_preserves_warnings()
    {
        var owner = SceneOwnerId.ForRow(GanttRowId.New());
        var frame = new SceneRect("row:frame", owner, ZLayer.Frame, new RectD(0, 0, 100, 100), new SceneStyle("Frame"));
        var bar = new SceneRect(
            "row:bar",
            owner,
            ZLayer.ActivityBody,
            new RectD(1, 2, 3, 4),
            new SceneStyle("Planned"),
            GanttEntityType.AsPlannedActivity
        );
        var warning = new SceneWarning(owner, "W1", "A warning.");

        var outcome = GanttScene.TryCreate(new RectD(0, 0, 100, 100), new RectD(10, 10, 90, 90), [bar, frame], [warning]);

        Assert.True(outcome.Succeeded);
        Assert.Equal(["row:bar", "row:frame"], outcome.Scene!.Primitives.Select(primitive => primitive.PrimitiveId));
        Assert.Equal([warning], outcome.Scene.Warnings);
    }

    [Fact]
    public void RectD_rejects_invalid_bounds_before_scene_creation()
    {
        _ = Assert.Throws<ArgumentOutOfRangeException>(() => new RectD(double.NaN, 0, 10, 10));
        _ = Assert.Throws<ArgumentOutOfRangeException>(() => new RectD(0, 0, double.PositiveInfinity, 10));
    }

    [Fact]
    public void TryCreate_refuses_duplicate_primitive_ids()
    {
        var owner = SceneOwnerId.ForRow(GanttRowId.New());
        var first = new SceneRect("row:bar", owner, ZLayer.ActivityBody, new RectD(0, 0, 1, 1), new SceneStyle("A"));
        var second = new SceneRect("row:bar", owner, ZLayer.ActivityBody, new RectD(1, 0, 1, 1), new SceneStyle("B"));

        var outcome = GanttScene.TryCreate(new RectD(0, 0, 10, 10), new RectD(0, 0, 10, 10), [first, second], []);

        Assert.False(outcome.Succeeded);
        Assert.Equal(SceneCreationRefusal.DuplicatePrimitiveId, outcome.Refusal);
    }

    [Fact]
    public void TryCreate_refuses_unresolved_group_child()
    {
        var owner = SceneOwnerId.ForRow(GanttRowId.New());
        var group = new SceneGroup("row:group", owner, ZLayer.Label, ["row:missing"]);

        var outcome = GanttScene.TryCreate(new RectD(0, 0, 10, 10), new RectD(0, 0, 10, 10), [group], []);

        Assert.False(outcome.Succeeded);
        Assert.Equal(SceneCreationRefusal.UnresolvedGroupChild, outcome.Refusal);
    }
}
