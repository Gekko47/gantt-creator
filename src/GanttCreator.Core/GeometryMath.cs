namespace GanttCreator.Core;

/// <summary>
/// Central numeric policies for point-based scene geometry.
/// </summary>
public static class GeometryMath
{
    /// <summary>The tolerance used when comparing geometry with host display precision.</summary>
    public const double Epsilon = 1e-6;

    /// <summary>Determines whether two finite values are equal within <see cref="Epsilon"/>.</summary>
    /// <param name="left">The first value in points.</param>
    /// <param name="right">The second value in points.</param>
    /// <returns><see langword="true"/> when the values are within the documented tolerance.</returns>
    public static bool ApproximatelyEqual(double left, double right) =>
        double.IsFinite(left) && double.IsFinite(right) && Math.Abs(left - right) <= Epsilon;

    /// <summary>Converts a point value to the single display precision used by Excel COM.</summary>
    /// <param name="value">The point value to convert.</param>
    /// <returns>The value represented at display precision.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="value"/> is not finite or cannot be represented as a finite <see cref="float"/>.</exception>
    public static float SnapToDisplayPrecision(double value) =>
        double.IsFinite(value) && value >= -float.MaxValue && value <= float.MaxValue
            ? (float)value
            : throw new ArgumentOutOfRangeException(
                nameof(value),
                value,
                "Geometry must be finite and representable as a finite float.");
}
