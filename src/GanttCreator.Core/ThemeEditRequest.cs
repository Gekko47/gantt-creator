namespace GanttCreator.Core;

/// <summary>
/// Why a theme edit was refused before any workbook mutation (R5.6 D2).
/// </summary>
public enum ThemeEditRefusal
{
    /// <summary>A style key is not a built-in editable preset.</summary>
    UnknownStyleKey = 0,

    /// <summary>A colour is neither empty nor uppercase <c>#RRGGBB</c>.</summary>
    BadColourFormat = 1,

    /// <summary>A metric is non-finite or outside its catalogue range.</summary>
    MetricOutOfRange = 2,
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
    IReadOnlyDictionary<string, double> MetricEdits)
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

