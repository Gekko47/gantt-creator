namespace GanttCreator.Core.Tests;

public sealed class GeometryMathTests
{
    [Fact]
    public void ApproximatelyEqual_accepts_values_within_epsilon()
    {
        Assert.True(GeometryMath.ApproximatelyEqual(10.0, 10.0 + GeometryMath.Epsilon / 2));
    }

    [Fact]
    public void ApproximatelyEqual_rejects_values_outside_epsilon()
    {
        Assert.False(GeometryMath.ApproximatelyEqual(10.0, 10.0 + (GeometryMath.Epsilon * 2)));
    }

    [Theory]
    [InlineData(double.NaN, 1.0)]
    [InlineData(1.0, double.NaN)]
    [InlineData(double.PositiveInfinity, double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity, double.NegativeInfinity)]
    public void ApproximatelyEqual_rejects_non_finite_values(double left, double right)
    {
        Assert.False(GeometryMath.ApproximatelyEqual(left, right));
    }

    [Fact]
    public void SnapToDisplayPrecision_converts_finite_value_once()
    {
        Assert.Equal(12.5f, GeometryMath.SnapToDisplayPrecision(12.5));
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void SnapToDisplayPrecision_rejects_non_finite_values(double value)
    {
        ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(
            () => GeometryMath.SnapToDisplayPrecision(value));
        Assert.Equal("value", exception.ParamName);
    }

    [Fact]
    public void SnapToDisplayPrecision_rejects_values_outside_float_range()
    {
        Assert.Throws<ArgumentOutOfRangeException>(
            () => GeometryMath.SnapToDisplayPrecision(double.MaxValue));
    }
}
