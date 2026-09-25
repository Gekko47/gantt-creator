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

    [Fact]
    public void TryCreate_refuses_a_cycle_spanning_two_groups()
    {
        // Neither group names itself, so SceneGroup accepts both; every child ID
        // also resolves, so the unresolved-child check cannot catch it. Only the
        // whole-set cycle check can see that expanding row:a revisits row:b and
        // never terminates.
        var owner = SceneOwnerId.ForRow(GanttRowId.New());
        var groupA = new SceneGroup("row:a", owner, ZLayer.Label, ["row:b"]);
        var groupB = new SceneGroup("row:b", owner, ZLayer.Label, ["row:a"]);

        var outcome = GanttScene.TryCreate(
            new RectD(0, 0, 10, 10),
            new RectD(0, 0, 10, 10),
            [groupA, groupB],
            []);

        Assert.False(outcome.Succeeded);
        Assert.Equal(SceneCreationRefusal.CyclicGroupChild, outcome.Refusal);
    }

    [Fact]
    public void TryCreate_refuses_a_cycle_spanning_three_groups()
    {
        // A longer cycle must be refused for the same reason; a two-group cycle
        // alone would not prove the walk follows more than one hop.
        var owner = SceneOwnerId.ForRow(GanttRowId.New());
        var groupA = new SceneGroup("row:a", owner, ZLayer.Label, ["row:b"]);
        var groupB = new SceneGroup("row:b", owner, ZLayer.Label, ["row:c"]);
        var groupC = new SceneGroup("row:c", owner, ZLayer.Label, ["row:a"]);

        var outcome = GanttScene.TryCreate(
            new RectD(0, 0, 10, 10),
            new RectD(0, 0, 10, 10),
            [groupA, groupB, groupC],
            []);

        Assert.False(outcome.Succeeded);
        Assert.Equal(SceneCreationRefusal.CyclicGroupChild, outcome.Refusal);
    }

    [Fact]
    public void TryCreate_accepts_a_diamond_of_groups_sharing_one_descendant()
    {
        // Two groups naming the same descendant is a shared child, not a cycle.
        // The walk colours completed nodes precisely so this legal shape is not
        // mistaken for one, and the scene still builds.
        var owner = SceneOwnerId.ForRow(GanttRowId.New());
        var leaf = new SceneRect("row:leaf", owner, ZLayer.Label, new RectD(0, 0, 1, 1), new SceneStyle("A"));
        var shared = new SceneGroup("row:shared", owner, ZLayer.Label, ["row:leaf"]);
        var left = new SceneGroup("row:left", owner, ZLayer.Label, ["row:shared"]);
        var right = new SceneGroup("row:right", owner, ZLayer.Label, ["row:shared"]);

        var outcome = GanttScene.TryCreate(
            new RectD(0, 0, 10, 10),
            new RectD(0, 0, 10, 10),
            [left, right, shared, leaf],
            []);

        Assert.True(outcome.Succeeded);
        Assert.Equal(4, outcome.Scene!.Primitives.Count);
    }

    [Fact]
    public void SceneGroup_refuses_a_direct_self_reference()
    {
        // The single-group half of the cycle rule, refused at the primitive
        // boundary before a scene ever exists.
        var owner = SceneOwnerId.ForRow(GanttRowId.New());

        _ = Assert.Throws<ArgumentException>(
            () => new SceneGroup("row:self", owner, ZLayer.Label, ["row:self"]));
    }
}
