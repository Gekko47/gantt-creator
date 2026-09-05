using System.Globalization;

namespace GanttCreator.Core;

/// <summary>
/// An immutable point in 2-D space using double-precision coordinates.
/// All scene geometry in this project is defined in points
/// (1 inch = 72 points).
/// </summary>
public readonly struct PointD : IEquatable<PointD>
{
    /// <summary>The horizontal coordinate in points.</summary>
    public double X { get; }

    /// <summary>The vertical coordinate in points.</summary>
    public double Y { get; }

    /// <summary>Initialises a point from x/y coordinates.</summary>
    /// <param name="x">The horizontal coordinate in points.</param>
    /// <param name="y">The vertical coordinate in points.</param>
    public PointD(double x, double y)
    {
        (X, Y) = (x, y);
    }

    /// <summary>The origin (0, 0).</summary>
    public static PointD Zero => new(0, 0);

    /// <summary>Returns a new point offset by <paramref name="dx"/> and <paramref name="dy"/>.</summary>
    /// <param name="dx">The horizontal offset in points.</param>
    /// <param name="dy">The vertical offset in points.</param>
    public PointD Offset(double dx, double dy) => new(X + dx, Y + dy);

    /// <inheritdoc cref="IEquatable{T}.Equals(T)"/>
    public bool Equals(PointD other) =>
        X.Equals(other.X) && Y.Equals(other.Y);

    /// <inheritdoc />
    public override bool Equals(object? obj) =>
        obj is PointD other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => HashCode.Combine(X, Y);

    /// <summary>Equality operator.</summary>
    public static bool operator ==(PointD left, PointD right) => left.Equals(right);

    /// <summary>Inequality operator.</summary>
    public static bool operator !=(PointD left, PointD right) => !left.Equals(right);

    /// <summary>
    /// Returns a coordinate string in the form "(X, Y)" with both
    /// components formatted to two decimal places. Formatting is
    /// culture-invariant so coordinate output is deterministic
    /// across locales; see ExportSize for the same discipline
    /// applied to width parsing.
    /// </summary>
    public override string ToString() =>
        string.Create(CultureInfo.InvariantCulture, $"({X:F2}, {Y:F2})");
}
