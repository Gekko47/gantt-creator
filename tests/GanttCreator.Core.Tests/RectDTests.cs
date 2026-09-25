namespace GanttCreator.Core.Tests;

public sealed class RectDTests
{
    [Fact]
    public void Constructor_preserves_values_and_derives_edges()
    {
        var rectangle = new RectD(10, 20, 30, 40);
        Assert.Equal(10, rectangle.X);
        Assert.Equal(20, rectangle.Y);
        Assert.Equal(30, rectangle.Width);
        Assert.Equal(40, rectangle.Height);
        Assert.Equal(10, rectangle.Left);
        Assert.Equal(40, rectangle.Right);
        Assert.Equal(20, rectangle.Top);
        Assert.Equal(60, rectangle.Bottom);
    }

    [Fact]
    public void Zero_extents_are_legal()
    {
        var rectangle = new RectD(4, 5, 0, 0);
        Assert.Equal(4, rectangle.Left);
        Assert.Equal(4, rectangle.Right);
        Assert.Equal(5, rectangle.Top);
        Assert.Equal(5, rectangle.Bottom);
    }

    [Fact]
    public void Equal_values_have_equal_equality_and_hashes()
    {
        var left = new RectD(1, 2, 3, 4);
        var right = new RectD(1, 2, 3, 4);
        Assert.Equal(left, right);
        Assert.True(left == right);
        Assert.Equal(left.GetHashCode(), right.GetHashCode());
    }

    [Fact]
    public void Contains_accepts_inside_and_boundary_points()
    {
        var rectangle = new RectD(10, 20, 30, 40);
        Assert.True(rectangle.Contains(new PointD(20, 30)));
        Assert.True(rectangle.Contains(new PointD(10, 20)));
        Assert.True(rectangle.Contains(new PointD(40, 60)));
        Assert.False(rectangle.Contains(new PointD(9.9, 30)));
    }

    [Fact]
    public void Contains_uses_epsilon_at_boundary()
    {
        var rectangle = new RectD(10, 20, 30, 40);
        Assert.True(rectangle.Contains(new PointD(10 - (GeometryMath.Epsilon / 2), 20)));
    }

    [Fact]
    public void IntersectsWith_accepts_overlap_touching_edges_and_corners()
    {
        var rectangle = new RectD(10, 10, 10, 10);
        Assert.True(rectangle.IntersectsWith(new RectD(15, 15, 10, 10)));
        Assert.True(rectangle.IntersectsWith(new RectD(20, 10, 10, 10)));
        Assert.True(rectangle.IntersectsWith(new RectD(20, 20, 10, 10)));
    }

    [Fact]
    public void IntersectsWith_rejects_disjoint_rectangles()
    {
        var rectangle = new RectD(0, 0, 10, 10);
        Assert.False(rectangle.IntersectsWith(new RectD(10.001, 0, 10, 10)));
    }

    [Fact]
    public void Negative_width_is_rejected()
    {
        ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(
            () => new RectD(0, 0, -1, 1));
        Assert.Equal("width", exception.ParamName);
    }

    [Fact]
    public void Negative_height_is_rejected()
    {
        ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(
            () => new RectD(0, 0, 1, -1));
        Assert.Equal("height", exception.ParamName);
    }

    [Theory]
    [InlineData(double.NaN, 1, 1)]
    [InlineData(double.PositiveInfinity, 1, 1)]
    [InlineData(1, double.NegativeInfinity, 1)]
    [InlineData(1, 1, double.NaN)]
    public void Non_finite_values_are_rejected(double x, double y, double width)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new RectD(x, y, width, 1));
    }

    [Fact]
    public void Negative_coordinates_remain_legal()
    {
        // A plot may start left of and above the origin; only the derived edges
        // must stay finite, so negative coordinates are preserved as-is.
        var rectangle = new RectD(-30, -40, 10, 10);

        Assert.Equal(-30, rectangle.Left);
        Assert.Equal(-20, rectangle.Right);
        Assert.Equal(-40, rectangle.Top);
        Assert.Equal(-30, rectangle.Bottom);
    }

    [Fact]
    public void Right_edge_overflow_is_rejected()
    {
        // Two large but individually finite values overflow when summed; a
        // non-finite edge would make Contains/IntersectsWith meaningless.
        // (MaxValue + 10 would merely round, so the additive case is required.)
        var exception = Assert.Throws<ArgumentOutOfRangeException>(
            () => new RectD(double.MaxValue, 0, double.MaxValue, 10));
        Assert.Equal("width", exception.ParamName);
    }

    [Fact]
    public void Bottom_edge_overflow_is_rejected()
    {
        var exception = Assert.Throws<ArgumentOutOfRangeException>(
            () => new RectD(0, double.MaxValue, 10, double.MaxValue));
        Assert.Equal("height", exception.ParamName);
    }
}
