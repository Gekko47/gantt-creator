namespace GanttCreator.Core;

/// <summary>An immutable two-dimensional size in points.</summary>
public readonly record struct SizeD
{
    /// <summary>Initialises a size from width and height.</summary>
    /// <param name="width">The horizontal extent in points.</param>
    /// <param name="height">The vertical extent in points.</param>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when either extent is negative or non-finite.</exception>
    public SizeD(double width, double height)
    {
        ValidateExtent(width, nameof(width));
        ValidateExtent(height, nameof(height));
        Width = width;
        Height = height;
    }

    /// <summary>The horizontal extent in points.</summary>
    public double Width { get; }

    /// <summary>The vertical extent in points.</summary>
    public double Height { get; }

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
