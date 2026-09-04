namespace GanttCreator.Core.Tests;

public class PointDTests
{
    [Fact]
    public void Zero_returns_origin()
    {
        var p = PointD.Zero;
        Assert.Equal(0.0, p.X);
        Assert.Equal(0.0, p.Y);
    }

    [Fact]
    public void Constructor_preserves_values()
    {
        var p = new PointD(3.5, 7.25);
        Assert.Equal(3.5, p.X);
        Assert.Equal(7.25, p.Y);
    }

    [Fact]
    public void Offset_returns_new_point()
    {
        var p = new PointD(10.0, 20.0);
        var q = p.Offset(1.5, -0.5);

        // Original is unchanged (immutable)
        Assert.Equal(10.0, p.X);
        Assert.Equal(20.0, p.Y);

        // New point has offset applied
        Assert.Equal(11.5, q.X);
        Assert.Equal(19.5, q.Y);
    }

    [Fact]
    public void Equals_same_values_returns_true()
    {
        var a = new PointD(1.0, 2.0);
        var b = new PointD(1.0, 2.0);
        Assert.True(a.Equals(b));
        Assert.True(a == b);
    }

    [Fact]
    public void Equals_different_values_returns_false()
    {
        var a = new PointD(1.0, 2.0);
        var b = new PointD(1.0, 3.0);
        Assert.False(a.Equals(b));
        Assert.True(a != b);
    }

    [Fact]
    public void GetHashCode_equal_points_have_equal_hashes()
    {
        var a = new PointD(99.9, 1.0);
        var b = new PointD(99.9, 1.0);
        Assert.Equal(a.GetHashCode(), b.GetHashCode());
    }

    [Fact]
    public void ToString_contains_both_coordinates()
    {
        var p = new PointD(1.5, 2.5);
        var s = p.ToString();
        Assert.Contains("1.50", s);
        Assert.Contains("2.50", s);
    }
}
