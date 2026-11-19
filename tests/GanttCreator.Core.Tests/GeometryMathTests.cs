using GanttCreator.Core.Scene;

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
    public void ApproximatelyEqual_with_a_tolerance_accepts_a_host_rounded_value() =>
        // R6.4: the group envelope is a host read-back stored in single precision
        // (measured 2026-11-18: 514.3622047pt requested, 514.3621826pt read back).
        Assert.True(GeometryMath.ApproximatelyEqual(
            514.3622047,
            514.3621826,
            HostOwnedGeometry.GeometryMatchTolerancePt));

    [Fact]
    public void ApproximatelyEqual_with_a_tolerance_rejects_a_value_beyond_it() =>
        Assert.False(GeometryMath.ApproximatelyEqual(
            10.0,
            10.0 + (HostOwnedGeometry.GeometryMatchTolerancePt * 2),
            HostOwnedGeometry.GeometryMatchTolerancePt));

    [Fact]
    public void ApproximatelyEqual_with_a_tolerance_accepts_the_exact_boundary() =>
        // The comparison is inclusive at the tolerance: the difference equals the
        // tolerance exactly, and a strict '<' would reject it.
        Assert.True(GeometryMath.ApproximatelyEqual(
            0.0,
            HostOwnedGeometry.GeometryMatchTolerancePt,
            HostOwnedGeometry.GeometryMatchTolerancePt));

    [Theory]
    [InlineData(double.NaN, 1.0)]
    [InlineData(1.0, double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity, double.NegativeInfinity)]
    public void ApproximatelyEqual_with_a_tolerance_still_rejects_non_finite_values(
        double left,
        double right) =>
        Assert.False(GeometryMath.ApproximatelyEqual(
            left,
            right,
            HostOwnedGeometry.GeometryMatchTolerancePt));

    [Fact]
    public void SnapToDisplayPrecision_converts_finite_value_once()
    {
        Assert.Equal(12.5f, GeometryMath.SnapToDisplayPrecision(12.5));
    }

    [Theory]
    [InlineData(float.MaxValue)]
    [InlineData(-float.MaxValue)]
    public void SnapToDisplayPrecision_accepts_the_float_range_boundaries(float boundary)
    {
        // Both range checks are inclusive at the boundary: a strict comparison would
        // refuse a value that is representable as a finite float.
        Assert.Equal(boundary, GeometryMath.SnapToDisplayPrecision(boundary));
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
        Assert.Contains("Geometry must be finite", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void SnapToDisplayPrecision_rejects_values_outside_float_range()
    {
        Assert.Throws<ArgumentOutOfRangeException>(
            () => GeometryMath.SnapToDisplayPrecision(double.MaxValue));
    }
}
