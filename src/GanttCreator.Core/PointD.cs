namespace GanttCreator.Core;

/// <summary>
/// An immutable point in 2-D space using double-precision coordinates.
/// All scene geometry in this project is defined in points
/// (1 inch = 72 points).
/// </summary>
public readonly struct PointD : IEquatable<PointD>
{
    public double X { get; }
    public double Y { get; }

    public PointD(double x, double y) => (X, Y) = (x, y);

    public static PointD Zero => new(0, 0);

    public PointD Offset(double dx, double dy) => new(X + dx, Y + dy);

    public bool Equals(PointD other) =>
        X.Equals(other.X) && Y.Equals(other.Y);

    public override bool Equals(object? obj) =>
        obj is PointD other && Equals(other);

    public override int GetHashCode() => HashCode.Combine(X, Y);

    public static bool operator ==(PointD left, PointD right) => left.Equals(right);
    public static bool operator !=(PointD left, PointD right) => !left.Equals(right);

    public override string ToString() => $"({X:F2}, {Y:F2})";
}
