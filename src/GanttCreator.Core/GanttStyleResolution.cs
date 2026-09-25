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
        out GanttStyleResolutionRefusal? refusal) =>
        TryResolve(
            type,
            GanttStyleRegistry.Empty,
            styleKey,
            fillColour,
            strokeColour,
            labelPosition,
            out resolved,
            out refusal);

    /// <summary>
    /// Resolves a style base from the catalogue registry, applying explicit per-row
    /// overrides.
    /// </summary>
    /// <remarks>
    /// R2.7c: a style present in <paramref name="registry"/> resolves from the
    /// registry, which is the only path that can see a user-authored style. A key
    /// absent from the registry falls back to the code-owned built-in preset, which
    /// preserves the behaviour of the compatibility overload.
    /// </remarks>
    /// <param name="type">The validated entity Type.</param>
    /// <param name="registry">The catalogue style registry.</param>
    /// <param name="styleKey">The selected named style, or null.</param>
    /// <param name="fillColour">The explicit fill override, or null.</param>
    /// <param name="strokeColour">The explicit stroke override, or null.</param>
    /// <param name="labelPosition">The explicit label override, or null.</param>
    /// <param name="resolved">The resolved style when successful.</param>
    /// <param name="refusal">The refusal when unsuccessful.</param>
    /// <returns><see langword="true"/> when a style was resolved.</returns>
    public static bool TryResolve(
        GanttEntityType type,
        GanttStyleRegistry registry,
        string? styleKey,
        string? fillColour,
        string? strokeColour,
        GanttLabelPosition? labelPosition,
        out GanttResolvedStyle? resolved,
        out GanttStyleResolutionRefusal? refusal)
    {
        ArgumentNullException.ThrowIfNull(registry);
        resolved = null;
        EntityTypeDefinition? definition = EntityTypeCatalog.GetDefinition(type);
        if (definition is null)
        {
            refusal = GanttStyleResolutionRefusal.UnknownType;
            return false;
        }

        if (!TrySelect(styleKey, definition, registry, out StyleBase selected))
        {
            refusal = styleKey is null
                ? GanttStyleResolutionRefusal.NoDefaultStyle
                : GanttStyleResolutionRefusal.StyleUnavailable;
            return false;
        }

        var usedFallback = styleKey is not null && !string.Equals(
            selected.StyleKey,
            styleKey,
            StringComparison.Ordinal);

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

    /// <summary>
    /// A style base flattened from either source, so the resolution maths does not
    /// branch on where the style came from.
    /// </summary>
    private readonly record struct StyleBase(
        string StyleKey,
        string FillColour,
        string StrokeColour,
        GanttLabelPosition DefaultLabelPosition);

    private static bool TrySelect(
        string? styleKey,
        EntityTypeDefinition definition,
        GanttStyleRegistry registry,
        out StyleBase selected)
    {
        if (styleKey is not null
            && registry.TryGet(styleKey, out GanttStyleDefinition? registered)
            && registered is { HasFormatting: true })
        {
            selected = new StyleBase(
                registered.StyleKey,
                registered.FillColour ?? string.Empty,
                registered.StrokeColour ?? string.Empty,
                registered.DefaultLabelPosition!.Value);
            return true;
        }

        if (styleKey is null
            && registry.TryGet(definition.DefaultStyleKey, out GanttStyleDefinition? typeDefault)
            && typeDefault is { HasFormatting: true })
        {
            selected = new StyleBase(
                typeDefault.StyleKey,
                typeDefault.FillColour ?? string.Empty,
                typeDefault.StrokeColour ?? string.Empty,
                typeDefault.DefaultLabelPosition!.Value);
            return true;
        }

        // Built-in preset fallback: keeps the compatibility overload and any
        // registry that predates R2.7c resolving exactly as before. A supplied but
        // unknown key falls back to the Type default preset, which is the
        // pre-R2.7c behaviour; it refuses only when the Type has no preset either.
        GanttStylePreset? preset = FindStyle(styleKey);
        if (preset is null && styleKey is not null)
        {
            preset = FindStyle(definition.DefaultStyleKey);
        }

        if (preset is null)
        {
            selected = default;
            return false;
        }

        selected = new StyleBase(
            preset.StyleKey,
            preset.FillColour,
            preset.StrokeColour,
            preset.DefaultLabelPosition);
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
    /// <summary>
    /// Returns the row cells after changing StyleKey and clearing all per-row
    /// formatting overrides. Re-selecting the row's current style key is a no-op:
    /// the overrides are kept, because the style did not change and clearing them
    /// would discard the user's explicit per-row formatting.
    /// </summary>
    /// <param name="row">The existing row.</param>
    /// <param name="newStyleKey">The new named style key.</param>
    /// <returns>The row with new style and empty override cells, or the unchanged
    /// row when the style key is the same.</returns>
    public static GanttRowDto ChangeStyleKey(GanttRowDto row, string? newStyleKey)
    {
        ArgumentNullException.ThrowIfNull(row);
        return string.Equals(newStyleKey, row.StyleKey, StringComparison.Ordinal)
            ? row
            : ApplyStyleKey(row, newStyleKey);
    }

    /// <summary>Writes the new style key and empties every per-row override cell.</summary>
    /// <param name="row">The existing row.</param>
    /// <param name="newStyleKey">The new named style key.</param>
    /// <returns>The row with new style and empty override cells.</returns>
    private static GanttRowDto ApplyStyleKey(GanttRowDto row, string? newStyleKey) =>
        row with
        {
            StyleKeyCell = newStyleKey is null ? GanttCells.Empty<string>() : GanttCells.Value(newStyleKey),
            LabelPositionCell = GanttCells.Empty<string>(),
            FillColourCell = GanttCells.Empty<string>(),
            StrokeColourCell = GanttCells.Empty<string>(),
        };
}
