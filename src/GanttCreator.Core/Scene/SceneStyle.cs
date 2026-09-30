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
    /// <param name="textColour">
    /// The resolved label text colour, if any. Null means the scene resolved no
    /// text colour for this entity, which is not the same fact as black.
    /// </param>
    public SceneStyle(
        string styleKey,
        ColourHex? fillColour = null,
        ColourHex? strokeColour = null,
        double? outlineWidthPt = null,
        GanttHatchPattern hatchPattern = GanttHatchPattern.None,
        string? fontFamily = null,
        double? fontSizePt = null,
        bool? bold = null,
        GanttTextAlignment? alignment = null,
        ColourHex? textColour = null)
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
        TextColour = textColour;
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
    public GanttTextAlignment? Alignment { get; }

    /// <summary>Gets the resolved label text colour, if applicable.</summary>
    /// <remarks>
    /// Entity guide section 17's inside/outside switch is a statement about this
    /// member, so it lives on the resolved style rather than in a renderer: the
    /// scene names the colour for each label it emits and the adapter writes it
    /// verbatim. A renderer that chose a text colour would be the substitution the
    /// architecture's scene-first rule forbids.
    /// </remarks>
    public ColourHex? TextColour { get; }

    /// <summary>Returns this style with the supplied label text colour.</summary>
    /// <param name="textColour">The text colour the copy should carry.</param>
    /// <returns>A copy carrying <paramref name="textColour"/>.</returns>
    /// <remarks>
    /// <para>
    /// A named copy rather than a <c>with</c> expression because every member of
    /// this type is get-only: it is a validated value, not a mutable bag, and
    /// <c>with</c> would need an <c>init</c> accessor to reach. Rebuilding through
    /// the constructor is also what keeps the validation in force, so a copy cannot
    /// be a way to smuggle in an outline width or font size the constructor
    /// would have rejected.
    /// </para>
    /// <para>
    /// The one caller that needs it is entity guide §17's inside/outside switch,
    /// which varies a single style member across the labels of one build.
    /// </para>
    /// </remarks>
    public SceneStyle WithTextColour(ColourHex? textColour) =>
        new(
            StyleKey,
            FillColour,
            StrokeColour,
            OutlineWidthPt,
            HatchPattern,
            FontFamily,
            FontSizePt,
            Bold,
            Alignment,
            textColour);
}
