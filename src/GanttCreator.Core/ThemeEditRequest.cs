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
/// One custom style edit: create, rename, or delete a user-authored style row.
/// Built-in preset keys are never custom and are excluded by validation.
/// </summary>
/// <param name="Operation">Create, rename, update, or delete.</param>
/// <param name="StyleKey">Target style key: new key for Create, old key for Rename, key to delete for Delete.</param>
/// <param name="NewStyleKey">For Rename, the new style key; null for Update and Delete.</param>
/// <param name="DisplayName">For Create/Rename, the human-readable style name; otherwise null.</param>
/// <param name="FillColour">For Create/Rename, the edited fill <c>#RRGGBB</c> or empty; otherwise null.</param>
/// <param name="StrokeColour">For Create/Rename, the edited stroke <c>#RRGGBB</c> or empty; otherwise null.</param>
/// <param name="TextColour">For Create/Rename, the edited label text <c>#RRGGBB</c> or empty; otherwise null.</param>
public sealed record CustomStyleEdit(
    CustomStyleOperation Operation,
    string StyleKey,
    string? NewStyleKey = null,
    string? DisplayName = null,
    string FillColour = "",
    string StrokeColour = "",
    string TextColour = "");

/// <summary>
/// One custom style's current state as read from the workbook registry.
/// </summary>
/// <param name="StyleKey">The exact style key.</param>
/// <param name="DisplayName">The human-readable style name.</param>
/// <param name="FillColour">The resolved fill <c>#RRGGBB</c>, or empty.</param>
/// <param name="StrokeColour">The resolved stroke <c>#RRGGBB</c>, or empty.</param>
/// <param name="TextColour">The resolved label text <c>#RRGGBB</c>, or empty.</param>
public sealed record CustomStyleDef(
    string StyleKey,
    string DisplayName,
    string FillColour,
    string StrokeColour,
    string TextColour);

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
    IReadOnlyList<CustomStyleDef>? CustomStyleDefs = null)
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
        var seenCustomKeys = new HashSet<string>(StringComparer.Ordinal);
        foreach (CustomStyleEdit edit in request.CustomStyleEdits ?? [])
        {
            if (string.IsNullOrWhiteSpace(edit.StyleKey))
            {
                refusal = ThemeEditRefusal.InvalidCustomStyle;
                message = "Custom style key must not be blank.";
                return false;
            }

            var key = edit.StyleKey.Trim();
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

                if (edit.Operation == CustomStyleOperation.Rename)
                {
                    var newKey = edit.NewStyleKey;
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

        foreach (var pair in request.MetricEdits)
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
        var warnings = new List<string>();
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

