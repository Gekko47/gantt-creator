using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

public sealed class ScenePrimitiveTests
{
    [Fact]
    public void CreateId_combines_owner_and_role_deterministically()
    {
        var owner = GanttRowId.New();
        Assert.Equal($"{owner.Value}:bar", ScenePrimitive.CreateId(owner, "bar"));
        _ = Assert.Throws<ArgumentException>(() => ScenePrimitive.CreateId(owner, " "));
    }

    [Fact]
    public void Rect_preserves_geometry_and_style()
    {
        var owner = GanttRowId.New();
        var style = new SceneStyle("Planned", fillColour: ColourHex.Parse("#92D050"));
        var rectangle = new SceneRect(
            "row:bar",
            owner,
            ZLayer.ActivityBody,
            new RectD(1, 2, 3, 4),
            style,
            GanttEntityType.AsPlannedActivity,
            1,
            0,
            4);

        Assert.Equal("row:bar", rectangle.PrimitiveId);
        Assert.Equal(owner, rectangle.OwnerId);
        Assert.Equal(new RectD(1, 2, 3, 4), rectangle.Bounds);
        Assert.Equal(style, rectangle.Style);
    }

    [Fact]
    public void Primitive_rejects_blank_id()
    {
        _ = Assert.Throws<ArgumentException>(() => new SceneRect(
            " ",
            GanttRowId.New(),
            ZLayer.Background,
            new RectD(0, 0, 1, 1),
            new SceneStyle("Default")));
    }

    [Fact]
    public void Primitive_rejects_negative_order_values()
    {
        _ = Assert.Throws<ArgumentOutOfRangeException>(() => new SceneRect(
            "row:bar",
            GanttRowId.New(),
            ZLayer.ActivityBody,
            new RectD(0, 0, 1, 1),
            new SceneStyle("Default"),
            laneOrder: -1));
        _ = Assert.Throws<ArgumentOutOfRangeException>(() => new SceneRect(
            "row:bar",
            GanttRowId.New(),
            ZLayer.ActivityBody,
            new RectD(0, 0, 1, 1),
            new SceneStyle("Default"),
            stackIndex: -1));
    }

    [Fact]
    public void Polygon_preserves_point_order_and_rejects_short_or_non_finite_points()
    {
        var owner = GanttRowId.New();
        PointD[] points = [new(0, 0), new(1, 0), new(0, 1)];
        var polygon = new ScenePolygon(
            "row:diamond",
            owner,
            ZLayer.Milestone,
            points,
            new SceneStyle("Milestone"),
            GanttEntityType.AsBuiltMilestone);

        Assert.Equal(points, polygon.Points);
        _ = Assert.Throws<ArgumentOutOfRangeException>(() => new ScenePolygon(
            "row:short",
            owner,
            ZLayer.Milestone,
            [new(0, 0), new(1, 0)],
            new SceneStyle("Milestone")));
        _ = Assert.Throws<ArgumentOutOfRangeException>(() => new ScenePolygon(
            "row:nan",
            owner,
            ZLayer.Milestone,
            [new(double.NaN, 0), new(1, 0), new(0, 1)],
            new SceneStyle("Milestone")));
    }

    [Fact]
    public void Group_preserves_child_order_and_rejects_blank_child()
    {
        var owner = GanttRowId.New();
        var group = new SceneGroup("row:group", owner, ZLayer.Label, ["row:a", "row:b"]);
        Assert.Equal(["row:a", "row:b"], group.ChildPrimitiveIds);
        _ = Assert.Throws<ArgumentException>(() => new SceneGroup(
            "row:group",
            owner,
            ZLayer.Label,
            ["row:a", " "]));
    }
}
