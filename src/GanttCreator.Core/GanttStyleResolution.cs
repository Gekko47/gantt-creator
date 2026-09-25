namespace GanttCreator.Core;

/// <summary>Why a named-style base could not be resolved.</summary>
public enum GanttStyleResolutionRefusal
{
    /// <summary>The Type is outside the code-owned catalogue.</summary>
    UnknownType = 0,
    /// <summary>The Type has no named-style default and no valid custom style was supplied.</summary>
    NoDefaultStyle = 1,
    /// <summary>The selected style key is not available as a formatted style.</summary>
    StyleUnavailable = 2,
}

/// <summary>Resolved named-style defaults with validated per-row overrides applied.</summary>
public sealed record GanttResolvedStyle
{
    /// <summary>Initialises a resolved style value.</summary>
    /// <param name="styleKey">The resolved base named style.</param>
    /// <param name="fillColour">The resolved fill colour, or null when none.</param>
    /// <param name="strokeColour">The resolved stroke colour, or null when none.</param>
    /// <param name="labelPosition">The resolved label position.</param>
    /// <param name="usedFallback">Whether the Type default replaced an unavailable style.</param>
    public GanttResolvedStyle(
        string styleKey,
        ColourHex? fillColour,
        ColourHex? strokeColour,
        GanttLabelPosition labelPosition,
        bool usedFallback)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleKey);
        StyleKey = styleKey;
        FillColour = fillColour;
        StrokeColour = strokeColour;
        LabelPosition = labelPosition;
        UsedFallback = usedFallback;
    }

    /// <summary>Gets the resolved base named style.</summary>
    public string StyleKey { get; }

    /// <summary>Gets the resolved fill colour.</summary>
    public ColourHex? FillColour { get; }

    /// <summary>Gets the resolved stroke colour.</summary>
    public ColourHex? StrokeColour { get; }

    /// <summary>Gets the resolved label position.</summary>
    public GanttLabelPosition LabelPosition { get; }

    /// <summary>Gets whether an unavailable style fell back to the Type default.</summary>
    public bool UsedFallback { get; }
}

/// <summary>Pure named-style and per-row override resolution.</summary>
public static class GanttStyleResolver
{
    /// <summary>Resolves a style base and applies explicit per-row overrides.</summary>
    /// <param name="type">The validated entity Type.</param>
    /// <param name="styleKey">The selected named style, or null.</param>
    /// <param name="fillColour">The explicit fill override, or null.</param>
    /// <param name="strokeColour">The explicit stroke override, or null.</param>
    /// <param name="labelPosition">The explicit label override, or null.</param>
    /// <param name="resolved">The resolved style when successful.</param>
    /// <param name="refusal">The refusal when unsuccessful.</param>
    /// <returns><see langword="true"/> when a style was resolved.</returns>
    public static bool TryResolve(
        GanttEntityType type,
        string? styleKey,
        string? fillColour,
        string? strokeColour,
        GanttLabelPosition? labelPosition,
        out GanttResolvedStyle? resolved,
        out GanttStyleResolutionRefusal? refusal)
    {
        resolved = null;
        EntityTypeDefinition? definition = EntityTypeCatalog.GetDefinition(type);
        if (definition is null)
        {
            refusal = GanttStyleResolutionRefusal.UnknownType;
            return false;
        }

        GanttStylePreset? selected = FindStyle(styleKey);
        var usedFallback = false;
        if (selected is null)
        {
            selected = FindStyle(definition.DefaultStyleKey);
            usedFallback = styleKey is not null;
            if (selected is null)
            {
                refusal = GanttStyleResolutionRefusal.NoDefaultStyle;
                return false;
            }
        }

        if (!TryParseOptionalColour(selected.FillColour, out ColourHex? baseFill)
            || !TryParseOptionalColour(selected.StrokeColour, out ColourHex? baseStroke)
            || !TryParseOptionalColour(fillColour, out ColourHex? overrideFill)
            || !TryParseOptionalColour(strokeColour, out ColourHex? overrideStroke))
        {
            refusal = GanttStyleResolutionRefusal.StyleUnavailable;
            return false;
        }

        resolved = new GanttResolvedStyle(
            selected.StyleKey,
            overrideFill ?? baseFill,
            overrideStroke ?? baseStroke,
            labelPosition ?? selected.DefaultLabelPosition,
            usedFallback);
        refusal = null;
        return true;
    }

    private static GanttStylePreset? FindStyle(string? styleKey) =>
        styleKey is null
            ? null
            : GanttCatalogues.StylePresets.FirstOrDefault(preset =>
                string.Equals(preset.StyleKey, styleKey, StringComparison.Ordinal));

    private static bool TryParseOptionalColour(string? text, out ColourHex? colour)
    {
        colour = null;
        return text is null
            || text.Length == 0
            || ColourHex.TryParse(text, out colour);
    }
}

/// <summary>Style-change operations that preserve the approved per-row override rule.</summary>
public static class GanttStyleChange
{
    /// <summary>Returns the row cells after changing StyleKey and clearing all per-row formatting overrides.</summary>
    /// <param name="row">The existing row.</param>
    /// <param name="newStyleKey">The new named style key.</param>
    /// <returns>The row with new style and empty override cells.</returns>
    public static GanttRowDto ChangeStyleKey(GanttRowDto row, string? newStyleKey)
    {
        ArgumentNullException.ThrowIfNull(row);
        return row with
        {
            StyleKeyCell = newStyleKey is null ? GanttCells.Empty<string>() : GanttCells.Value(newStyleKey),
            LabelPositionCell = GanttCells.Empty<string>(),
            FillColourCell = GanttCells.Empty<string>(),
            StrokeColourCell = GanttCells.Empty<string>(),
        };
    }
}
