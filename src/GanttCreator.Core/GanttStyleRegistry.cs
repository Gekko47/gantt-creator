namespace GanttCreator.Core;

/// <summary>
/// The capability projection for one named style used by row validation.
/// </summary>
/// <param name="StyleKey">The exact style key.</param>
/// <param name="AllowedLabelPositions">The label positions permitted by the style.</param>
/// <param name="ColourCapability">The colour override capability of the style.</param>
public sealed record GanttStyleDefinition(
    string StyleKey,
    IReadOnlySet<GanttLabelPosition> AllowedLabelPositions,
    EntityColourCapability ColourCapability);

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
