namespace GanttCreator.Core.Scene;

/// <summary>Immutable resolved style information carried by a scene primitive.</summary>
public sealed record SceneStyle
{
    /// <summary>Initialises a validated resolved style value.</summary>
    /// <param name="styleKey">The resolved style key.</param>
    /// <param name="fillColour">The resolved fill colour, if any.</param>
    /// <param name="strokeColour">The resolved stroke colour, if any.</param>
    /// <param name="outlineWidthPt">The resolved outline or line width, if applicable.</param>
    /// <param name="hatchPattern">The resolved hatch pattern, if applicable.</param>
    /// <param name="fontFamily">The resolved font family, if applicable.</param>
    /// <param name="fontSizePt">The resolved font size, if applicable.</param>
    /// <param name="bold">Whether the resolved typography is bold, if applicable.</param>
    /// <param name="alignment">The resolved text alignment, if applicable.</param>
    public SceneStyle(
        string styleKey,
        ColourHex? fillColour = null,
        ColourHex? strokeColour = null,
        double? outlineWidthPt = null,
        GanttHatchPattern hatchPattern = GanttHatchPattern.None,
        string? fontFamily = null,
        double? fontSizePt = null,
        bool? bold = null,
        GanttLabelPosition? alignment = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleKey);
        if (outlineWidthPt is { } outline && (!double.IsFinite(outline) || outline < 0))
        {
            throw new ArgumentOutOfRangeException(nameof(outlineWidthPt));
        }

        if (fontSizePt is { } fontSize && (!double.IsFinite(fontSize) || fontSize <= 0))
        {
            throw new ArgumentOutOfRangeException(nameof(fontSizePt));
        }

        StyleKey = styleKey;
        FillColour = fillColour;
        StrokeColour = strokeColour;
        OutlineWidthPt = outlineWidthPt;
        HatchPattern = hatchPattern;
        FontFamily = fontFamily;
        FontSizePt = fontSizePt;
        Bold = bold;
        Alignment = alignment;
    }

    /// <summary>Gets the resolved style key.</summary>
    public string StyleKey { get; }

    /// <summary>Gets the resolved fill colour, if any.</summary>
    public ColourHex? FillColour { get; }

    /// <summary>Gets the resolved stroke colour, if any.</summary>
    public ColourHex? StrokeColour { get; }

    /// <summary>Gets the resolved outline or line width, if applicable.</summary>
    public double? OutlineWidthPt { get; }

    /// <summary>Gets the resolved hatch pattern, if applicable.</summary>
    public GanttHatchPattern HatchPattern { get; }

    /// <summary>Gets the resolved font family, if applicable.</summary>
    public string? FontFamily { get; }

    /// <summary>Gets the resolved font size, if applicable.</summary>
    public double? FontSizePt { get; }

    /// <summary>Gets whether the resolved typography is bold, if applicable.</summary>
    public bool? Bold { get; }

    /// <summary>Gets the resolved text alignment, if applicable.</summary>
    public GanttLabelPosition? Alignment { get; }
}
