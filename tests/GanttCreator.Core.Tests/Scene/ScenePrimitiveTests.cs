using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

public sealed class ScenePrimitiveTests
{
    [Fact]
    public void CreateId_combines_owner_and_role_deterministically()
    {
        var owner = SceneOwnerId.ForRow(GanttRowId.New());
        Assert.Equal($"{owner.Value}:bar", ScenePrimitive.CreateId(owner, "bar"));
        _ = Assert.Throws<ArgumentException>(() => ScenePrimitive.CreateId(owner, " "));
    }

    [Fact]
    public void Chart_owner_is_distinct_from_row_owner_and_creates_chart_ids()
    {
        var row = SceneOwnerId.ForRow(GanttRowId.New());

        Assert.Equal(SceneOwnerKind.Row, row.Kind);
        Assert.Equal(SceneOwnerKind.Chart, SceneOwnerId.Chart.Kind);
        Assert.NotEqual(row, SceneOwnerId.Chart);
        Assert.Equal("chart:frame", ScenePrimitive.CreateId(SceneOwnerId.Chart, "frame"));
    }

    [Fact]
    public void TryParse_accepts_only_the_closed_owner_pairs()
    {
        Assert.True(SceneOwnerId.TryParse("chart", "chart", out SceneOwnerId? chart));
        Assert.Same(SceneOwnerId.Chart, chart);
        Assert.True(SceneOwnerId.TryParse("row", GanttRowId.New().Value, out SceneOwnerId? row));
        Assert.Equal(SceneOwnerKind.Row, row!.Kind);
        Assert.False(SceneOwnerId.TryParse("chart", "row", out _));
        Assert.False(SceneOwnerId.TryParse("row", "chart", out _));
        Assert.False(SceneOwnerId.TryParse("other", "chart", out _));
    }

    [Fact]
    public void Rect_preserves_geometry_and_style()
    {
        var owner = SceneOwnerId.ForRow(GanttRowId.New());
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
            4
        );

        Assert.Equal("row:bar", rectangle.PrimitiveId);
        Assert.Equal(owner, rectangle.OwnerId);
        Assert.Equal(new RectD(1, 2, 3, 4), rectangle.Bounds);
        Assert.Equal(style, rectangle.Style);
    }

    [Fact]
    public void Primitive_rejects_blank_id()
    {
        _ = Assert.Throws<ArgumentException>(() =>
            new SceneRect(" ", SceneOwnerId.ForRow(GanttRowId.New()), ZLayer.Background, new RectD(0, 0, 1, 1), new SceneStyle("Default"))
        );
    }

    [Fact]
    public void Primitive_rejects_negative_order_values()
    {
        _ = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new SceneRect(
                "row:bar",
                SceneOwnerId.ForRow(GanttRowId.New()),
                ZLayer.ActivityBody,
                new RectD(0, 0, 1, 1),
                new SceneStyle("Default"),
                laneOrder: -1
            )
        );
        _ = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new SceneRect(
                "row:bar",
                SceneOwnerId.ForRow(GanttRowId.New()),
                ZLayer.ActivityBody,
                new RectD(0, 0, 1, 1),
                new SceneStyle("Default"),
                stackIndex: -1
            )
        );
    }

    [Fact]
    public void Polygon_preserves_point_order_and_rejects_short_or_non_finite_points()
    {
        var owner = SceneOwnerId.ForRow(GanttRowId.New());
        PointD[] points = [new(0, 0), new(1, 0), new(0, 1)];
        var polygon = new ScenePolygon(
            "row:diamond",
            owner,
            ZLayer.Milestone,
            points,
            new SceneStyle("Milestone"),
            GanttEntityType.AsBuiltMilestone
        );

        Assert.Equal(points, polygon.Points);
        _ = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ScenePolygon("row:short", owner, ZLayer.Milestone, [new(0, 0), new(1, 0)], new SceneStyle("Milestone"))
        );
        _ = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ScenePolygon("row:nan", owner, ZLayer.Milestone, [new(double.NaN, 0), new(1, 0), new(0, 1)], new SceneStyle("Milestone"))
        );
    }

    [Fact]
    public void Primitive_rejects_a_negative_sort_order()
    {
        // SortOrder is a non-negative user ordering key (BadSortOrder in the row
        // validator); a negative value must not silently reorder the scene.
        var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new SceneRect(
                "row:bar",
                SceneOwnerId.ForRow(GanttRowId.New()),
                ZLayer.ActivityBody,
                new RectD(0, 0, 1, 1),
                new SceneStyle("Default"),
                sortOrder: -1
            )
        );
        Assert.Equal("sortOrder", exception.ParamName);
    }

    [Fact]
    public void Group_rejects_duplicate_child_primitive_ids()
    {
        // A repeated child would render twice and make the snapshot non-canonical.
        var exception = Assert.Throws<ArgumentException>(() =>
            new SceneGroup("row:group", SceneOwnerId.ForRow(GanttRowId.New()), ZLayer.Label, ["row:a", "row:a"]));
        Assert.Contains("unique", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Group_rejects_a_direct_self_reference()
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            new SceneGroup("row:group", SceneOwnerId.ForRow(GanttRowId.New()), ZLayer.Label, ["row:group"]));
        Assert.Contains("itself", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Group_preserves_child_order_and_rejects_blank_child()
    {
        var owner = SceneOwnerId.ForRow(GanttRowId.New());
        var group = new SceneGroup("row:group", owner, ZLayer.Label, ["row:a", "row:b"]);
        Assert.Equal(["row:a", "row:b"], group.ChildPrimitiveIds);
        _ = Assert.Throws<ArgumentException>(() => new SceneGroup("row:group", owner, ZLayer.Label, ["row:a", " "]));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void Polygon_rejects_fewer_than_two_points(int count)
    {
        // The count guard is the polygon contract's own boundary. Only the two-point
        // case was covered; zero and one prove the guard is a count check and not a
        // special case for "a segment", which would be legal for a line primitive.
        var owner = SceneOwnerId.ForRow(GanttRowId.New());
        PointD[] points = [.. Enumerable.Range(0, count).Select(i => new PointD(i, 0))];

        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => new ScenePolygon("row:short", owner, ZLayer.Milestone, points, new SceneStyle("Milestone")));
    }

    [Theory]
    [InlineData(double.PositiveInfinity, 0)]
    [InlineData(double.NegativeInfinity, 0)]
    [InlineData(0, double.PositiveInfinity)]
    [InlineData(0, double.NegativeInfinity)]
    public void Polygon_rejects_an_infinite_point_on_either_axis(double x, double y)
    {
        // Only NaN was covered. Infinity reaches the same guard by a different
        // arithmetic path, and an infinite coordinate would serialise as a
        // non-finite snapshot value, so it must be refused the same way.
        var owner = SceneOwnerId.ForRow(GanttRowId.New());

        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => new ScenePolygon(
                "row:infinite",
                owner,
                ZLayer.Milestone,
                [new PointD(x, y), new(1, 0), new(0, 1)],
                new SceneStyle("Milestone")));
    }

    [Fact]
    public void Polygon_reports_the_count_guard_before_the_finiteness_guard()
    {
        // Both guards throw the same exception type, so a refactor that swapped
        // their order would be invisible to a type-only assertion. The message is
        // what distinguishes them, so it is pinned here.
        var owner = SceneOwnerId.ForRow(GanttRowId.New());

        ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(
            () => new ScenePolygon(
                "row:short-and-nan",
                owner,
                ZLayer.Milestone,
                [new PointD(double.NaN, 0), new(1, 0)],
                new SceneStyle("Milestone")));

        Assert.Contains("at least three points", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Polygon_copies_its_point_list_so_a_later_mutation_cannot_reach_it()
    {
        // The constructor copies with a collection expression, but nothing proved
        // it: a polygon is an immutable scene primitive, and a caller mutating its
        // source list afterwards must not be able to change a built scene.
        var owner = SceneOwnerId.ForRow(GanttRowId.New());
        List<PointD> points = [new(0, 0), new(1, 0), new(0, 1)];

        ScenePolygon polygon = new("row:diamond", owner, ZLayer.Milestone, points, new SceneStyle("Milestone"));
        points[0] = new PointD(99, 99);

        Assert.Equal(new PointD(0, 0), polygon.Points[0]);
    }

    [Fact]
    public void PointD_itself_performs_no_validation_which_is_why_the_polygon_guard_exists()
    {
        // The contrast control for the two tests above. PointD is a bare struct, so
        // a non-finite coordinate reaches a SceneLine and is caught later by
        // SceneValidator, but a ScenePolygon refuses it at construction. This is
        // asserted so a future change to PointD is not mistaken for a redundancy.
        PointD point = new(double.NaN, 0);

        Assert.True(double.IsNaN(point.X));
        var owner = SceneOwnerId.ForRow(GanttRowId.New());
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => new ScenePolygon(
                "row:nan",
                owner,
                ZLayer.Milestone,
                [point, new(1, 0), new(0, 1)],
                new SceneStyle("Milestone")));
    }
}
