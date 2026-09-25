namespace GanttCreator.Core;

/// <summary>
/// The projection of one named style used by row validation: its capabilities and,
/// when the catalogue supplied them, its resolved formatting.
/// </summary>
/// <remarks>
/// <para>
/// R2.7b added the capability projection (allowed label positions, colour
/// capability). R2.7c adds the resolved formatting the entity guide requires
/// before rendering — a blank <c>FillColour</c> "uses resolved style", and renderers
/// must never substitute colours. The formatting parameters are optional so the
/// capability-only compatibility path stays legal.
/// </para>
/// </remarks>
/// <param name="StyleKey">The exact style key.</param>
/// <param name="AllowedLabelPositions">The label positions permitted by the style.</param>
/// <param name="ColourCapability">The colour override capability of the style.</param>
/// <param name="DefaultLabelPosition">The label position used when the row leaves LabelPosition blank, or null when unknown.</param>
/// <param name="FillColour">The resolved fill <c>#RRGGBB</c>, empty when the entity has no fill.</param>
/// <param name="StrokeColour">The resolved stroke <c>#RRGGBB</c>, empty when the entity has no stroke.</param>
/// <param name="TextColour">The resolved label text <c>#RRGGBB</c>, empty when the style has no text colour.</param>
/// <param name="HatchPattern">The resolved hatch pattern.</param>
/// <param name="HatchPitchPt">The hatch pitch in points.</param>
/// <param name="HatchLinePt">The hatch stroke width in points.</param>
/// <param name="StandardOutlinePt">The outline width in points.</param>
/// <param name="ActivityHeightPt">The entity height in points.</param>
/// <param name="MilestoneSizePt">The diamond tip-to-tip size in points.</param>
public sealed record GanttStyleDefinition(
    string StyleKey,
    IReadOnlySet<GanttLabelPosition> AllowedLabelPositions,
    EntityColourCapability ColourCapability,
    GanttLabelPosition? DefaultLabelPosition = null,
    string? FillColour = null,
    string? StrokeColour = null,
    string? TextColour = null,
    GanttHatchPattern? HatchPattern = null,
    double HatchPitchPt = 0,
    double HatchLinePt = 0,
    double StandardOutlinePt = 0,
    double ActivityHeightPt = 0,
    double MilestoneSizePt = 0)
{
    /// <summary>
    /// Gets whether the catalogue supplied resolved formatting for this style.
    /// A capability-only definition cannot be rendered and must resolve as
    /// unavailable rather than as a blank style.
    /// </summary>
    public bool HasFormatting => DefaultLabelPosition is not null;

    /// <summary>
    /// Validates the resolved formatting. Optional when
    /// <see cref="DefaultLabelPosition"/> is null (capability-only definition).
    /// </summary>
    /// <exception cref="ArgumentException">Thrown when a colour is neither empty nor uppercase <c>#RRGGBB</c>.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when a metric is negative or not finite, or an enum value is undefined.</exception>
    public void ValidateFormatting()
    {
        if (DefaultLabelPosition is not { } defaultPosition)
        {
            return;
        }

        if (!Enum.IsDefined(defaultPosition))
        {
            throw new ArgumentOutOfRangeException(nameof(DefaultLabelPosition));
        }

        if (HatchPattern is not { } hatch)
        {
            throw new ArgumentOutOfRangeException(nameof(HatchPattern));
        }

        if (!Enum.IsDefined(hatch))
        {
            throw new ArgumentOutOfRangeException(nameof(HatchPattern));
        }

        ValidateColour(FillColour, nameof(FillColour));
        ValidateColour(StrokeColour, nameof(StrokeColour));
        ValidateColour(TextColour, nameof(TextColour));
        ValidateMetric(HatchPitchPt, nameof(HatchPitchPt));
        ValidateMetric(HatchLinePt, nameof(HatchLinePt));
        ValidateMetric(StandardOutlinePt, nameof(StandardOutlinePt));
        ValidateMetric(ActivityHeightPt, nameof(ActivityHeightPt));
        ValidateMetric(MilestoneSizePt, nameof(MilestoneSizePt));
    }

    private static void ValidateColour(string? text, string parameterName)
    {
        if (text is null)
        {
            return;
        }

        if (text.Length > 0 && !GanttColourToken.IsValidHex(text))
        {
            throw new ArgumentException(
                $"Colour must be empty or uppercase #RRGGBB. Parameter name: {parameterName}",
                parameterName);
        }
    }

    private static void ValidateMetric(double value, string parameterName)
    {
        if (!double.IsFinite(value) || value < 0)
        {
            throw new ArgumentOutOfRangeException(parameterName, value, "Metric must be finite and non-negative.");
        }
    }
}

/// <summary>
/// Immutable registry of named-style capabilities projected from the
/// configuration catalogue. It contains no Office or worksheet concepts.
/// </summary>
public sealed class GanttStyleRegistry
{
    private readonly Dictionary<string, GanttStyleDefinition> _styles;

    /// <summary>Initialises a registry from style definitions.</summary>
    /// <param name="styles">The style definitions in catalogue order.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="styles"/> is null.</exception>
    /// <exception cref="ArgumentException">Thrown when a key is blank or duplicated.</exception>
    public GanttStyleRegistry(IEnumerable<GanttStyleDefinition> styles)
    {
        ArgumentNullException.ThrowIfNull(styles);
        var map = new Dictionary<string, GanttStyleDefinition>(StringComparer.Ordinal);
        foreach (GanttStyleDefinition style in styles)
        {
            ArgumentNullException.ThrowIfNull(style);
            if (style.AllowedLabelPositions is null
                || style.AllowedLabelPositions.Count == 0
                || style.AllowedLabelPositions.Any(position => !Enum.IsDefined(position)))
            {
                throw new ArgumentException("Style label capabilities must contain defined positions.", nameof(styles));
            }

            const EntityColourCapability knownCapabilities =
                EntityColourCapability.Fill | EntityColourCapability.Stroke |
                EntityColourCapability.Hatch | EntityColourCapability.StyleDefined;
            if ((style.ColourCapability & ~knownCapabilities) != EntityColourCapability.None)
            {
                throw new ArgumentOutOfRangeException(nameof(styles));
            }

            if (string.IsNullOrWhiteSpace(style.StyleKey) || !map.TryAdd(style.StyleKey, style))
            {
                throw new ArgumentException("Style keys must be nonblank and unique.", nameof(styles));
            }

            // R2.7c: resolved formatting is validated where the registry is built,
            // so a malformed value is refused once rather than surfacing later as
            // a render-time surprise.
            style.ValidateFormatting();
        }

        _styles = map;
    }

    /// <summary>Gets an empty registry for the compatibility validator path.</summary>
    public static GanttStyleRegistry Empty { get; } = new([]);

    /// <summary>Gets the number of registered styles.</summary>
    public int Count => _styles.Count;

    /// <summary>Finds a style by exact key.</summary>
    /// <param name="styleKey">The exact style key.</param>
    /// <param name="style">The matching definition, or null.</param>
    /// <returns>Whether the key was found.</returns>
    public bool TryGet(string? styleKey, out GanttStyleDefinition? style)
    {
        if (styleKey is null)
        {
            style = null;
            return false;
        }

        return _styles.TryGetValue(styleKey, out style);
    }
}
