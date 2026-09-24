namespace GanttCreator.Core;

/// <summary>An immutable axis-aligned rectangle in points.</summary>
public readonly record struct RectD
{
    /// <summary>Initialises a rectangle from its origin and extents.</summary>
    /// <param name="x">The left coordinate in points.</param>
    /// <param name="y">The top coordinate in points.</param>
    /// <param name="width">The horizontal extent in points.</param>
    /// <param name="height">The vertical extent in points.</param>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when a coordinate is non-finite or an extent is negative/non-finite.</exception>
    public RectD(double x, double y, double width, double height)
    {
        ValidateCoordinate(x, nameof(x));
        ValidateCoordinate(y, nameof(y));
        ValidateExtent(width, nameof(width));
        ValidateExtent(height, nameof(height));
        X = x;
        Y = y;
        Width = width;
        Height = height;
    }

    /// <summary>The left coordinate in points.</summary>
    public double X { get; }

    /// <summary>The top coordinate in points.</summary>
    public double Y { get; }

    /// <summary>The horizontal extent in points.</summary>
    public double Width { get; }

    /// <summary>The vertical extent in points.</summary>
    public double Height { get; }

    /// <summary>The left edge in points.</summary>
    public double Left => X;

    /// <summary>The exclusive right edge in points.</summary>
    public double Right => X + Width;

    /// <summary>The top edge in points.</summary>
    public double Top => Y;

    /// <summary>The exclusive bottom edge in points.</summary>
    public double Bottom => Y + Height;

    /// <summary>Determines whether a point is within the rectangle, including its edges within tolerance.</summary>
    /// <param name="point">The point to test.</param>
    /// <returns><see langword="true"/> when the point is inside or on the boundary.</returns>
    public bool Contains(PointD point) =>
        point.X >= Left - GeometryMath.Epsilon &&
        point.X <= Right + GeometryMath.Epsilon &&
        point.Y >= Top - GeometryMath.Epsilon &&
        point.Y <= Bottom + GeometryMath.Epsilon;

    /// <summary>Determines whether this rectangle intersects another rectangle, including touching edges and corners.</summary>
    /// <param name="other">The other rectangle.</param>
    /// <returns><see langword="true"/> when the closed rectangles intersect.</returns>
    public bool IntersectsWith(RectD other) =>
        Left <= other.Right + GeometryMath.Epsilon &&
        Right + GeometryMath.Epsilon >= other.Left &&
        Top <= other.Bottom + GeometryMath.Epsilon &&
        Bottom + GeometryMath.Epsilon >= other.Top;

    private static void ValidateCoordinate(double value, string parameterName)
    {
        if (!double.IsFinite(value))
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                value,
                "Coordinate must be finite.");
        }
    }

    private static void ValidateExtent(double value, string parameterName)
    {
        if (!double.IsFinite(value) || value < 0)
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                value,
                "Extent must be finite and non-negative.");
        }
    }
}
