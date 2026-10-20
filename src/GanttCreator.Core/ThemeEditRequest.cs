namespace GanttCreator.Core;

/// <summary>
/// Why a theme edit was refused before any workbook mutation (R5.6 D2).
/// </summary>
public enum ThemeEditRefusal
{
    /// <summary>A style key is not a built-in editable preset or a valid custom style key.</summary>
    UnknownStyleKey = 0,

    /// <summary>A colour is neither empty nor uppercase <c>#RRGGBB</c>.</summary>
    BadColourFormat = 1,

    /// <summary>A metric is non-finite or outside its catalogue range.</summary>
    MetricOutOfRange = 2,

    /// <summary>A custom style operation is invalid: duplicate key, missing target, or built-in key used as custom.</summary>
    InvalidCustomStyle = 3,
}

/// <summary>
/// One style preset's edited colours: empty means the entity has no fill/stroke
/// (preserved as empty, never defaulted — the <c>OfficeStyleMapper</c> rule).
/// </summary>
/// <param name="StyleKey">The exact built-in preset key.</param>
/// <param name="FillColour">The edited fill <c>#RRGGBB</c>, or empty.</param>
/// <param name="StrokeColour">The edited stroke <c>#RRGGBB</c>, or empty.</param>
/// <param name="TextColour">The edited label text <c>#RRGGBB</c>, or empty when the preset default should be used.</param>
public sealed record ThemeStyleEdit(
    string StyleKey,
    string FillColour,
    string StrokeColour,
    string TextColour);

/// <summary>The operation a custom style edit performs.</summary>
public enum CustomStyleOperation
{
    /// <summary>Create a new user-named style row.</summary>
    Create = 0,

    /// <summary>
    /// Rename an existing custom style (and optionally edit its fields).
    /// <c>NewStyleKey</c> is required and must differ from <c>StyleKey</c>.
    /// </summary>
    Rename = 1,

    /// <summary>
    /// Update an existing custom style's fields in place (upsert by key).
    /// <c>StyleKey</c> is unchanged; <c>NewStyleKey</c> is null.
    /// </summary>
    Update = 2,

    /// <summary>Delete an existing custom style row.</summary>
    Delete = 3,
}

/// <summary>
/// One custom style edit: create, rename, update, or delete a user-authored
/// style row. Built-in preset keys are never custom and are excluded by
/// validation.
/// </summary>
/// <remarks>
/// <para>
/// A null optional member means <b>not carried by this edit</b>: the writer
/// preserves the stored value. That distinction is the reason the nullable
/// members exist at all. An earlier revision rebuilt the whole row from
/// hard-coded catalogue defaults and back-filled the gaps with
/// <c>??=</c>, but the fresh defaults were never null, so the back-fill never
/// fired and an update that changed one colour silently reset the style's
/// pattern, sizes, and label vocabulary to the catalogue defaults.
/// </para>
/// <para>
/// The three colour members are not nullable: an empty colour is a
/// meaningful value (the entity genuinely has no fill or stroke), so they are
/// always carried.
/// </para>
/// </remarks>
/// <param name="Operation">Create, rename, update, or delete.</param>
/// <param name="StyleKey">Target style key: new key for Create, old key for Rename, key to delete for Delete.</param>
/// <param name="NewStyleKey">For Rename, the new style key; null for Update and Delete.</param>
/// <param name="DisplayName">The human-readable style name; null preserves the stored name.</param>
/// <param name="FillColour">The fill <c>#RRGGBB</c>, or empty for no fill.</param>
/// <param name="StrokeColour">The stroke <c>#RRGGBB</c>, or empty for no stroke.</param>
/// <param name="TextColour">The label text <c>#RRGGBB</c>, or empty.</param>
/// <param name="HatchPattern">The fill pattern; null preserves the stored pattern.</param>
/// <param name="StandardOutlinePt">The outline width in points; null preserves the stored value.</param>
/// <param name="ActivityHeightPt">The bar height in points; null preserves the stored value.</param>
/// <param name="MilestoneSizePt">The milestone tip-to-tip size in points; null preserves the stored value.</param>
/// <param name="DefaultLabelPosition">The default label position; null preserves the stored position.</param>
/// <param name="AllowedLabelPositions">The permitted label positions; null preserves the stored set.</param>
/// <param name="StrokeDash">The stroke dash pattern; null preserves the stored dash.</param>
public sealed record CustomStyleEdit(
    CustomStyleOperation Operation,
    string StyleKey,
    string? NewStyleKey = null,
    string? DisplayName = null,
    string FillColour = "",
    string StrokeColour = "",
    string TextColour = "",
    GanttHatchPattern? HatchPattern = null,
    double? StandardOutlinePt = null,
    double? ActivityHeightPt = null,
    double? MilestoneSizePt = null,
    GanttLabelPosition? DefaultLabelPosition = null,
    IReadOnlySet<GanttLabelPosition>? AllowedLabelPositions = null,
    GanttStrokeDash? StrokeDash = null);

/// <summary>
/// One custom style's current state as read from the workbook registry.
/// </summary>
/// <param name="StyleKey">The exact style key.</param>
/// <param name="DisplayName">The human-readable style name.</param>
/// <param name="FillColour">The resolved fill <c>#RRGGBB</c>, or empty.</param>
/// <param name="StrokeColour">The resolved stroke <c>#RRGGBB</c>, or empty.</param>
/// <param name="TextColour">The resolved label text <c>#RRGGBB</c>, or empty.</param>
/// <param name="HatchPattern">The resolved fill pattern.</param>
/// <param name="StandardOutlinePt">The resolved outline width in points.</param>
/// <param name="ActivityHeightPt">The resolved bar height in points.</param>
/// <param name="MilestoneSizePt">The resolved milestone tip-to-tip size in points.</param>
/// <param name="DefaultLabelPosition">The resolved default label position.</param>
/// <param name="AllowedLabelPositions">
/// The resolved permitted label positions, or null when the caller does not
/// carry them.
/// </param>
/// <param name="StrokeDash">The resolved stroke dash pattern.</param>
public sealed record CustomStyleDef(
    string StyleKey,
    string DisplayName,
    string FillColour,
    string StrokeColour,
    string TextColour,
    GanttHatchPattern HatchPattern = GanttHatchPattern.None,
    double StandardOutlinePt = 0,
    double ActivityHeightPt = 0,
    double MilestoneSizePt = 0,
    GanttLabelPosition DefaultLabelPosition = GanttLabelPosition.Auto,
    IReadOnlySet<GanttLabelPosition>? AllowedLabelPositions = null,
    GanttStrokeDash StrokeDash = GanttStrokeDash.Solid);

/// <summary>
/// The pure R5.6 theme edit: per-preset colours plus metric values, validated
/// against the code-owned catalogue before any COM runs.
/// </summary>
/// <remarks>
/// No Office types: the Office writer consumes the validated request. Colour
/// warnings are advisory flags on the outcome, never silent replacements
/// (entity-guide rule); shipped green/blue/red defaults never warn on open.
/// </remarks>
public sealed record ThemeEditRequest(
    IReadOnlyList<ThemeStyleEdit> StyleEdits,
    IReadOnlyDictionary<string, double> MetricEdits,
    IReadOnlyList<CustomStyleEdit>? CustomStyleEdits = null,
    IReadOnlyList<CustomStyleDef>? CustomStyleDefs = null,
    IReadOnlyDictionary<string, string>? ChromeColourEdits = null)
{
    /// <summary>The metric tokens editable by the R5.6 dialog (work-item scope).</summary>
    public static readonly IReadOnlySet<string> EditableMetrics =
        new HashSet<string>(StringComparer.Ordinal)
        {
            "ActivityHeightPt", "MilestoneSizePt", "StandardOutlinePt",
            "GanttRowHeightPt", "SplitterHeightPt", "SpacerHeightPt",
            "GridLinePt", "MajorBoundaryPt", "DelineatorLinePt",
            "HatchPitchPt", "HatchLinePt",
        };

    /// <summary>
    /// Validates the request against the catalogue ranges and colour format.
    /// </summary>
    /// <param name="request">The request to validate.</param>
    /// <param name="refusal">The refusal when invalid.</param>
    /// <param name="message">The user-actionable message when invalid.</param>
    /// <returns>True when every edit is in range and well-formed.</returns>
    public static bool TryValidate(
        ThemeEditRequest? request,
        out ThemeEditRefusal? refusal,
        out string? message)
    {
        refusal = null;
        message = null;
        if (request is null)
        {
            refusal = ThemeEditRefusal.UnknownStyleKey;
            message = "Theme edit is missing.";
            return false;
        }

        foreach (ThemeStyleEdit edit in request.StyleEdits)
        {
            if (!IsBuiltInPreset(edit.StyleKey))
            {
                refusal = ThemeEditRefusal.UnknownStyleKey;
                message = $"Unknown style '{edit.StyleKey}'.";
                return false;
            }

            if (!IsOptionalColour(edit.FillColour)
                || !IsOptionalColour(edit.StrokeColour)
                || !IsOptionalColour(edit.TextColour))
            {
                refusal = ThemeEditRefusal.BadColourFormat;
                message = $"Style '{edit.StyleKey}' colour must be empty or uppercase #RRGGBB.";
                return false;
            }
        }

        // Validate custom style edits: keys must be non-empty, not collide with
        // built-in presets, and be unique within the request. Colours on
        // Create/Rename/Update must be well-formed. Rename/Delete targets are
        // checked for existence at write time (the Office writer holds the
        // registry). Update keeps its key, so it can never collide with itself.
        HashSet<string> seenCustomKeys = new HashSet<string>(StringComparer.Ordinal);
        foreach (CustomStyleEdit edit in request.CustomStyleEdits ?? [])
        {
            if (string.IsNullOrWhiteSpace(edit.StyleKey))
            {
                refusal = ThemeEditRefusal.InvalidCustomStyle;
                message = "Custom style key must not be blank.";
                return false;
            }

            string key = edit.StyleKey.Trim();
            if (IsBuiltInPreset(key))
            {
                refusal = ThemeEditRefusal.InvalidCustomStyle;
                message = $"Style key '{key}' is a built-in preset and cannot be used as a custom style.";
                return false;
            }

            if (!seenCustomKeys.Add(key))
            {
                refusal = ThemeEditRefusal.InvalidCustomStyle;
                message = $"Custom style key '{key}' is used more than once in the same request.";
                return false;
            }

            if (edit.Operation is CustomStyleOperation.Create or CustomStyleOperation.Rename or CustomStyleOperation.Update)
            {
                if ((edit.FillColour.Length > 0 && !IsOptionalColour(edit.FillColour))
                    || (edit.StrokeColour.Length > 0 && !IsOptionalColour(edit.StrokeColour))
                    || (edit.TextColour.Length > 0 && !IsOptionalColour(edit.TextColour)))
                {
                    refusal = ThemeEditRefusal.BadColourFormat;
                    message = $"Custom style '{key}' colour must be empty or uppercase #RRGGBB.";
                    return false;
                }

                // The fill pattern is a closed enum: an undefined member would
                // serialise to a name the reader cannot parse, and the reader
                // refuses the whole catalogue rather than skipping the row.
                if (edit.HatchPattern is { } hatch && !Enum.IsDefined(hatch))
                {
                    refusal = ThemeEditRefusal.InvalidCustomStyle;
                    message = $"Custom style '{key}' fill pattern is not a defined pattern.";
                    return false;
                }

                // The label vocabulary is a closed enum too, and the default
                // must be one of the permitted positions: the reader enforces
                // exactly this, so a row that breaks it would be unreadable.
                if (edit.DefaultLabelPosition is { } defaultPosition && !Enum.IsDefined(defaultPosition))
                {
                    refusal = ThemeEditRefusal.InvalidCustomStyle;
                    message = $"Custom style '{key}' default label position is not a defined position.";
                    return false;
                }

                // The stroke dash is a closed enum: an undefined member would
                // serialise to a name the reader cannot parse, and the reader
                // refuses the whole catalogue rather than skipping the row.
                if (edit.StrokeDash is { } dash && !Enum.IsDefined(dash))
                {
                    refusal = ThemeEditRefusal.InvalidCustomStyle;
                    message = $"Custom style '{key}' stroke dash is not a defined dash pattern.";
                    return false;
                }

                if (edit.AllowedLabelPositions is { } allowed)
                {
                    if (allowed.Count == 0 || allowed.Any(position => !Enum.IsDefined(position)))
                    {
                        refusal = ThemeEditRefusal.InvalidCustomStyle;
                        message = $"Custom style '{key}' must permit at least one defined label position.";
                        return false;
                    }

                    if (edit.DefaultLabelPosition is { } required
                        && !allowed.Contains(required))
                    {
                        refusal = ThemeEditRefusal.InvalidCustomStyle;
                        message =
                            $"Custom style '{key}' default label position '{required}' is not in its permitted set.";
                        return false;
                    }
                }

                // The style's own sizes are range-checked against the same
                // catalogue tokens the global metrics use, so a custom style
                // cannot carry a size the entity guide forbids.
                if (!IsValidCustomStyleMetric(edit.StandardOutlinePt, "StandardOutlinePt")
                    || !IsValidCustomStyleMetric(edit.ActivityHeightPt, "ActivityHeightPt")
                    || !IsValidCustomStyleMetric(edit.MilestoneSizePt, "MilestoneSizePt"))
                {
                    refusal = ThemeEditRefusal.MetricOutOfRange;
                    message = $"Custom style '{key}' has a size outside the catalogue range.";
                    return false;
                }

                if (edit.Operation == CustomStyleOperation.Rename)
                {
                    string? newKey = edit.NewStyleKey;
                    if (string.IsNullOrWhiteSpace(newKey))
                    {
                        refusal = ThemeEditRefusal.InvalidCustomStyle;
                        message = $"Rename of custom style '{key}' requires a new key.";
                        return false;
                    }
                    newKey = newKey.Trim();
                    if (IsBuiltInPreset(newKey))
                    {
                        refusal = ThemeEditRefusal.InvalidCustomStyle;
                        message = $"Custom style rename target '{newKey}' is a built-in preset key.";
                        return false;
                    }
                    if (!seenCustomKeys.Add(newKey))
                    {
                        refusal = ThemeEditRefusal.InvalidCustomStyle;
                        message = $"Custom style rename target '{newKey}' collides with another key in the same request.";
                        return false;
                    }
                }
            }
        }

        // The chart-chrome colours: the key must be one the catalogue defines
        // and the value an uppercase #RRGGBB. The reader refuses a malformed
        // colour rather than substituting one, so the write path must not
        // accept one either.
        foreach (KeyValuePair<string, string> pair in request.ChromeColourEdits
            ?? new Dictionary<string, string>(StringComparer.Ordinal))
        {
            if (!GanttChartChrome.IsColourSettingKey(pair.Key))
            {
                refusal = ThemeEditRefusal.UnknownStyleKey;
                message = $"'{pair.Key}' is not a chart-chrome colour setting.";
                return false;
            }

            if (!GanttColourToken.IsValidHex(pair.Value))
            {
                refusal = ThemeEditRefusal.BadColourFormat;
                message = $"Chart colour '{pair.Key}' must be uppercase #RRGGBB.";
                return false;
            }
        }

        foreach (KeyValuePair<string, double> pair in request.MetricEdits)
        {
            GanttMetricToken? token = FindMetric(pair.Key);
            if (token is null || !EditableMetrics.Contains(pair.Key))
            {
                refusal = ThemeEditRefusal.MetricOutOfRange;
                message = $"Metric '{pair.Key}' is not editable in theme settings.";
                return false;
            }

            if (!double.IsFinite(pair.Value) || pair.Value < token.Minimum || pair.Value > token.Maximum)
            {
                refusal = ThemeEditRefusal.MetricOutOfRange;
                message =
                    $"Metric '{pair.Key}' value {pair.Value} is outside [{token.Minimum}, {token.Maximum}].";
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Collects non-blocking contrast advisories for edited text/fill pairs.
    /// Empty when nothing edited crosses the policy; never blocks the write.
    /// </summary>
    /// <param name="request">The validated request.</param>
    /// <returns>Warning lines, or empty.</returns>
    public static IReadOnlyList<string> ContrastWarnings(ThemeEditRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        List<string> warnings = [];
        foreach (ThemeStyleEdit edit in request.StyleEdits)
        {
            if (edit.FillColour.Length == 0 || edit.TextColour.Length == 0)
            {
                continue;
            }

            if (!ColourHex.TryParse(edit.TextColour, out ColourHex? text)
                || !ColourHex.TryParse(edit.FillColour, out ColourHex? fill)
                || text is null || fill is null)
            {
                continue;
            }

            if (!ThemeContrast.MeetsMinimum(text, fill))
            {
                warnings.Add(
                    $"Style '{edit.StyleKey}' text {edit.TextColour} on fill {edit.FillColour} "
                    + "is below 4.5:1 contrast.");
            }

            if (ThemeContrast.IsRedGreenPair(text, fill))
            {
                warnings.Add(
                    $"Style '{edit.StyleKey}' is a red-green pair that hue-deficient vision may confuse.");
            }
        }

        return warnings;
    }

    private static bool IsBuiltInPreset(string styleKey)
    {
        try
        {
            _ = GanttCatalogues.GetPreset(styleKey);
            return true;
        }
        catch (ArgumentOutOfRangeException)
        {
            return false;
        }
    }

    /// <summary>
    /// Validates one optional custom-style size against its catalogue token's
    /// range. A null value is valid: it means the edit does not carry that
    /// size, and the writer preserves the stored one.
    /// </summary>
    /// <param name="value">The edited size, or null when not carried.</param>
    /// <param name="tokenName">The catalogue metric token name.</param>
    /// <returns>Whether the value is absent or inside the token's range.</returns>
    /// <remarks>
    /// Zero is valid for a style size and means <b>not applicable to this
    /// entity kind</b>: a style used by an activity carries a zero milestone
    /// size, and a style used by a milestone carries a zero bar height. The
    /// built-in presets use exactly that convention, so a style size is
    /// range-checked only when it is non-zero. The catalogue's minimum for
    /// <c>ActivityHeightPt</c> is 2 and for <c>MilestoneSizePt</c> is 3, so
    /// treating zero as a violation would refuse every style that names only
    /// one entity kind.
    /// </remarks>
    private static bool IsValidCustomStyleMetric(double? value, string tokenName)
    {
        if (value is not { } metric)
        {
            return true;
        }

        if (!double.IsFinite(metric))
        {
            return false;
        }

        if (metric == 0)
        {
            return true;
        }

        GanttMetricToken? token = FindMetric(tokenName);
        return token is not null
            && metric >= token.Minimum
            && metric <= token.Maximum;
    }

    private static GanttMetricToken? FindMetric(string name)
    {
        foreach (GanttMetricToken token in GanttCatalogues.Metrics)
        {
            if (string.Equals(token.Name, name, StringComparison.Ordinal))
            {
                return token;
            }
        }

        return null;
    }

    private static bool IsOptionalColour(string value) =>
        value.Length == 0 || GanttColourToken.IsValidHex(value);
}

