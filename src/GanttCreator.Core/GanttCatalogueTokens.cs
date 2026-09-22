namespace GanttCreator.Core;

/// <summary>
/// The hatch fill pattern of a style preset, materialised to
/// <c>tblGanttStyles</c> as its exact name. Procurement entities hatch; the
/// pattern itself is part of the style preset, the pitch and stroke width
/// come from <see cref="GanttCatalogues.Metrics"/> (ADR-0007 D7).
/// </summary>
public enum GanttHatchPattern
{
    /// <summary>Solid fill; no hatch overlay.</summary>
    None = 0,

    /// <summary>Forward diagonal hatch (top-left to bottom-right).</summary>
    ForwardDiagonal = 1,

    /// <summary>Backward diagonal hatch (bottom-left to top-right).</summary>
    BackwardDiagonal = 2,

    /// <summary>Crossing diagonal hatch.</summary>
    Cross = 3,
}

/// <summary>
/// One named geometry metric token from the entity guide's "Shared metric
/// tokens" table (23 tokens). Immutable; the catalogue is code-owned and the
/// only source (ADR-0007 D2). Value equality enables round-trip comparison.
/// </summary>
public sealed record GanttMetricToken
{
    /// <summary>Gets the exact token name; part of the workbook schema.</summary>
    public string Name { get; }

    /// <summary>Gets the first-release default, in points.</summary>
    public double DefaultValue { get; }

    /// <summary>Gets the valid initial range minimum, inclusive.</summary>
    public double Minimum { get; }

    /// <summary>Gets the valid initial range maximum, inclusive.</summary>
    public double Maximum { get; }

    /// <summary>
    /// Initialises a metric token with validation.
    /// </summary>
    /// <param name="name">The exact token name; part of the workbook schema.</param>
    /// <param name="defaultValue">The first-release default, in points.</param>
    /// <param name="minimum">The valid initial range minimum, inclusive.</param>
    /// <param name="maximum">The valid initial range maximum, inclusive.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="name"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// Thrown when the name is empty, any value is not finite, the range is
    /// inverted, or the default lies outside the range.
    /// </exception>
    public GanttMetricToken(string name, double defaultValue, double minimum, double maximum)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (name.Length == 0)
        {
            throw new ArgumentException("Metric token name must not be empty.", nameof(name));
        }

        if (!IsFinite(defaultValue))
        {
            throw new ArgumentException($"Metric '{name}' default must be finite.", nameof(defaultValue));
        }

        if (!IsFinite(minimum))
        {
            throw new ArgumentException($"Metric '{name}' minimum must be finite.", nameof(minimum));
        }

        if (!IsFinite(maximum))
        {
            throw new ArgumentException($"Metric '{name}' maximum must be finite.", nameof(maximum));
        }

        if (minimum > maximum)
        {
            throw new ArgumentException($"Metric '{name}' range is inverted ({minimum} > {maximum}).", nameof(minimum));
        }

        if (defaultValue < minimum || defaultValue > maximum)
        {
            throw new ArgumentException($"Metric '{name}' default {defaultValue} lies outside [{minimum}, {maximum}].", nameof(defaultValue));
        }

        Name = name;
        DefaultValue = defaultValue;
        Minimum = minimum;
        Maximum = maximum;
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}

/// <summary>
/// One named colour token from the entity guide's "Shared colour and
/// typography tokens" table (21 tokens). Hex values are stored as uppercase
/// <c>#RRGGBB</c> with an explicit alpha of <c>FF</c> implied unless
/// transparency is named (ADR-0007 D2). Immutable; value equality enables
/// round-trip comparison.
/// </summary>
public sealed record GanttColourToken
{
    /// <summary>Gets the exact token name; part of the workbook schema.</summary>
    public string Name { get; }

    /// <summary>Gets the uppercase <c>#RRGGBB</c> initial value.</summary>
    public string HexValue { get; }

    /// <summary>
    /// Initialises a colour token with validation.
    /// </summary>
    /// <param name="name">The exact token name; part of the workbook schema.</param>
    /// <param name="hexValue">The uppercase <c>#RRGGBB</c> initial value.</param>
    /// <exception cref="ArgumentNullException">Thrown when an argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// Thrown when the name is empty or the hex value is not uppercase
    /// <c>#RRGGBB</c> (the <see cref="GanttValidationCodes.BadColourFormat"/>
    /// convention).
    /// </exception>
    public GanttColourToken(string name, string hexValue)
    {
        ArgumentNullException.ThrowIfNull(name);
        ArgumentNullException.ThrowIfNull(hexValue);
        if (name.Length == 0)
        {
            throw new ArgumentException("Colour token name must not be empty.", nameof(name));
        }

        if (!IsValidHex(hexValue))
        {
            throw new ArgumentException(
                $"Colour '{name}' value '{hexValue}' is not uppercase #RRGGBB.",
                nameof(hexValue));
        }

        Name = name;
        HexValue = hexValue;
    }

    /// <summary>
    /// Determines whether the text is exactly an uppercase <c>#RRGGBB</c>
    /// colour (the <see cref="GanttValidationCodes.BadColourFormat"/>
    /// convention). Shared by the catalogue and the style presets.
    /// </summary>
    /// <param name="text">The text to check.</param>
    /// <returns><see langword="true"/> when the text is uppercase <c>#RRGGBB</c>.</returns>
    public static bool IsValidHex(string text)
    {
        if (text is null || text.Length != 7 || text[0] != '#')
        {
            return false;
        }

        for (var index = 1; index < 7; index++)
        {
            var c = text[index];
            if (c is not ((>= '0' and <= '9') or (>= 'A' and <= 'F')))
            {
                return false;
            }
        }

        return true;
    }
}

/// <summary>
/// One named typography token from the entity guide (6 tokens). The value is
/// stored as text exactly as written to the configuration tables
/// (ADR-0007 D2/D7); numeric sizes use invariant culture. Immutable; value
/// equality enables round-trip comparison.
/// </summary>
public sealed record GanttTypographyToken
{
    /// <summary>Gets the exact token name; part of the workbook schema.</summary>
    public string Name { get; }

    /// <summary>Gets the token value as text (a font family name or a point size).</summary>
    public string Value { get; }

    /// <summary>Gets a value indicating whether the token's bold flag is set.</summary>
    public bool IsBold { get; }

    /// <summary>
    /// Initialises a typography token with validation.
    /// </summary>
    /// <param name="name">The exact token name; part of the workbook schema.</param>
    /// <param name="value">The token value as text (a font family name or a point size).</param>
    /// <param name="isBold">Whether the token's bold flag is set.</param>
    /// <exception cref="ArgumentNullException">Thrown when an argument is null.</exception>
    /// <exception cref="ArgumentException">Thrown when the name or value is empty.</exception>
    public GanttTypographyToken(string name, string value, bool isBold)
    {
        ArgumentNullException.ThrowIfNull(name);
        ArgumentNullException.ThrowIfNull(value);
        if (name.Length == 0)
        {
            throw new ArgumentException("Typography token name must not be empty.", nameof(name));
        }

        if (value.Length == 0)
        {
            throw new ArgumentException($"Typography token '{name}' value must not be empty.", nameof(value));
        }

        Name = name;
        Value = value;
        IsBold = isBold;
    }
}

/// <summary>
/// One first-release setting key and its default value, approved in
/// ADR-0007 D3. Values are text: booleans are <c>TRUE</c>/<c>FALSE</c>,
/// numerals are invariant. Immutable; value equality enables round-trip
/// comparison.
/// </summary>
public sealed record GanttSettingDefinition
{
    /// <summary>Gets the exact setting key; part of the workbook schema.</summary>
    public string Key { get; }

    /// <summary>Gets the first-release default (empty allowed, e.g. <c>ChartTitle</c>).</summary>
    public string DefaultValue { get; }

    /// <summary>
    /// Initialises a setting definition with validation.
    /// </summary>
    /// <param name="key">The exact setting key; part of the workbook schema.</param>
    /// <param name="defaultValue">The first-release default (empty allowed, e.g. <c>ChartTitle</c>).</param>
    /// <exception cref="ArgumentNullException">Thrown when an argument is null.</exception>
    /// <exception cref="ArgumentException">Thrown when the key is empty.</exception>
    public GanttSettingDefinition(string key, string defaultValue)
    {
        ArgumentNullException.ThrowIfNull(key);
        ArgumentNullException.ThrowIfNull(defaultValue);
        if (key.Length == 0)
        {
            throw new ArgumentException("Setting key must not be empty.", nameof(key));
        }

        Key = key;
        DefaultValue = defaultValue;
    }
}

/// <summary>
/// One built-in style preset: the resolved fill/stroke/hatch/text colours
/// and metric values for one style key, materialised to
/// <c>tblGanttStyles</c> (ADR-0007 D7). Resolved at catalogue-construction
/// time from the token defaults; the sheet stores resolved values, not
/// token references. Immutable; value equality enables round-trip
/// comparison.
/// </summary>
public sealed record GanttStylePreset
{
    /// <summary>Gets the exact style key; part of the workbook schema.</summary>
    public string StyleKey { get; }

    /// <summary>Gets the human-readable preset name.</summary>
    public string DisplayName { get; }

    /// <summary>Gets the resolved fill <c>#RRGGBB</c>, or empty when the entity has no fill.</summary>
    public string FillColour { get; }

    /// <summary>Gets the resolved stroke <c>#RRGGBB</c>, or empty when the entity has no stroke.</summary>
    public string StrokeColour { get; }

    /// <summary>Gets the hatch pattern.</summary>
    public GanttHatchPattern HatchPattern { get; }

    /// <summary>Gets the hatch pitch in points.</summary>
    public double HatchPitchPt { get; }

    /// <summary>Gets the hatch stroke width in points.</summary>
    public double HatchLinePt { get; }

    /// <summary>Gets the resolved label text <c>#RRGGBB</c>.</summary>
    public string TextColour { get; }

    /// <summary>Gets the outline width in points (0 when not applicable).</summary>
    public double StandardOutlinePt { get; }

    /// <summary>Gets the entity height in points (0 when not applicable, e.g. a delineator).</summary>
    public double ActivityHeightPt { get; }

    /// <summary>Gets the diamond tip-to-tip size in points (0 when not a milestone).</summary>
    public double MilestoneSizePt { get; }

    /// <summary>
    /// Initialises a style preset with validation.
    /// </summary>
    /// <param name="styleKey">The exact style key; part of the workbook schema.</param>
    /// <param name="displayName">The human-readable preset name.</param>
    /// <param name="fillColour">The resolved fill <c>#RRGGBB</c>, or empty when the entity has no fill.</param>
    /// <param name="strokeColour">The resolved stroke <c>#RRGGBB</c>, or empty when the entity has no stroke.</param>
    /// <param name="hatchPattern">The hatch pattern.</param>
    /// <param name="hatchPitchPt">The hatch pitch in points.</param>
    /// <param name="hatchLinePt">The hatch stroke width in points.</param>
    /// <param name="textColour">The resolved label text <c>#RRGGBB</c>.</param>
    /// <param name="standardOutlinePt">The outline width in points (0 when not applicable).</param>
    /// <param name="activityHeightPt">The entity height in points (0 when not applicable).</param>
    /// <param name="milestoneSizePt">The diamond tip-to-tip size in points (0 when not a milestone).</param>
    /// <exception cref="ArgumentNullException">Thrown when a text argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// Thrown when the style key or display name is empty, a colour is
    /// neither empty nor uppercase <c>#RRGGBB</c>, or a metric is negative
    /// or not finite.
    /// </exception>
    public GanttStylePreset(
        string styleKey,
        string displayName,
        string fillColour,
        string strokeColour,
        GanttHatchPattern hatchPattern,
        double hatchPitchPt,
        double hatchLinePt,
        string textColour,
        double standardOutlinePt,
        double activityHeightPt,
        double milestoneSizePt)
    {
        ArgumentNullException.ThrowIfNull(styleKey);
        ArgumentNullException.ThrowIfNull(displayName);
        ArgumentNullException.ThrowIfNull(fillColour);
        ArgumentNullException.ThrowIfNull(strokeColour);
        ArgumentNullException.ThrowIfNull(textColour);
        if (styleKey.Length == 0)
        {
            throw new ArgumentException("Style key must not be empty.", nameof(styleKey));
        }

        if (displayName.Length == 0)
        {
            throw new ArgumentException($"Style '{styleKey}' display name must not be empty.", nameof(displayName));
        }

        ValidateOptionalColour(styleKey, fillColour, nameof(fillColour));
        ValidateOptionalColour(styleKey, strokeColour, nameof(strokeColour));
        ValidateOptionalColour(styleKey, textColour, nameof(textColour));
        ValidateMetric(styleKey, hatchPitchPt, nameof(hatchPitchPt));
        ValidateMetric(styleKey, hatchLinePt, nameof(hatchLinePt));
        ValidateMetric(styleKey, standardOutlinePt, nameof(standardOutlinePt));
        ValidateMetric(styleKey, activityHeightPt, nameof(activityHeightPt));
        ValidateMetric(styleKey, milestoneSizePt, nameof(milestoneSizePt));

        StyleKey = styleKey;
        DisplayName = displayName;
        FillColour = fillColour;
        StrokeColour = strokeColour;
        HatchPattern = hatchPattern;
        HatchPitchPt = hatchPitchPt;
        HatchLinePt = hatchLinePt;
        TextColour = textColour;
        StandardOutlinePt = standardOutlinePt;
        ActivityHeightPt = activityHeightPt;
        MilestoneSizePt = milestoneSizePt;
    }

    private static void ValidateOptionalColour(string styleKey, string colour, string parameterName)
    {
        if (colour.Length > 0 && !GanttColourToken.IsValidHex(colour))
        {
            throw new ArgumentException(
                $"Style '{styleKey}' colour '{colour}' is neither empty nor uppercase #RRGGBB.",
                parameterName);
        }
    }

    private static void ValidateMetric(string styleKey, double value, string parameterName)
    {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0)
        {
            throw new ArgumentException(
                $"Style '{styleKey}' metric {parameterName} must be a finite non-negative number.",
                parameterName);
        }
    }
}

/// <summary>
/// One projected row of <c>tblGanttTypes</c>: exactly one
/// <see cref="EntityTypeCatalog"/> entry (ADR-0007 D2). No second type list
/// is authored anywhere. Immutable; value equality enables round-trip
/// comparison.
/// </summary>
/// <param name="TypeName">The code identity (the enum name).</param>
/// <param name="DisplayName">The exact workbook display name.</param>
/// <param name="Kind">The visual-entity class name.</param>
/// <param name="DateMode">The date-mode name.</param>
/// <param name="DefaultStyleKey">The default style key, or empty for <c>Custom Activity</c>.</param>
/// <param name="ColourCapability">The colour-capability flags as text.</param>
/// <param name="RequiresStyleKey">Whether the row must supply a <c>StyleKey</c> (<c>TRUE</c>/<c>FALSE</c>).</param>
/// <param name="AllowedLabelPositions">The permitted label positions, space-joined, or <c>"None"</c> alone when empty.</param>
public sealed record GanttTypeCatalogueRow(
    string TypeName,
    string DisplayName,
    string Kind,
    string DateMode,
    string DefaultStyleKey,
    string ColourCapability,
    string RequiresStyleKey,
    string AllowedLabelPositions);
