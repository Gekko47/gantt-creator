namespace GanttCreator.Core.Tests;

public sealed class SizeDTests
{
    [Fact]
    public void Constructor_preserves_values()
    {
        var size = new SizeD(12.5, 8.0);
        Assert.Equal(12.5, size.Width);
        Assert.Equal(8.0, size.Height);
    }

    [Fact]
    public void Zero_extents_are_legal()
    {
        var size = new SizeD(0, 0);
        Assert.Equal(0, size.Width);
        Assert.Equal(0, size.Height);
    }

    [Fact]
    public void Equal_values_have_equal_equality_and_hashes()
    {
        var left = new SizeD(3.5, 7.25);
        var right = new SizeD(3.5, 7.25);
        Assert.Equal(left, right);
        Assert.True(left == right);
        Assert.Equal(left.GetHashCode(), right.GetHashCode());
    }

    [Fact]
    public void Different_values_are_not_equal()
    {
        var left = new SizeD(3.5, 7.25);
        var right = new SizeD(3.5, 7.5);
        Assert.NotEqual(left, right);
        Assert.True(left != right);
    }

    [Fact]
    public void Negative_width_is_rejected()
    {
        ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(
            () => new SizeD(-0.1, 1));
        Assert.Equal("width", exception.ParamName);
    }

    [Fact]
    public void Negative_height_is_rejected()
    {
        ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(
            () => new SizeD(1, -0.1));
        Assert.Equal("height", exception.ParamName);
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void Non_finite_width_is_rejected(double width)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new SizeD(width, 1));
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void Non_finite_height_is_rejected(double height)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new SizeD(1, height));
    }
}
