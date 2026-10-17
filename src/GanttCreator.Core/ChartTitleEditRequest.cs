namespace GanttCreator.Core;

/// <summary>
/// Why a chart-title edit was refused before any workbook mutation (R5.6b).
/// </summary>
public enum ChartTitleEditRefusal
{
    /// <summary>The title is blank after trimming leading and trailing whitespace.</summary>
    BlankTitle = 0,
}

/// <summary>
/// The pure R5.6b chart-title edit: the trimmed non-blank title plus the title-band
/// visibility flag, validated before any COM runs.
/// </summary>
/// <remarks>
/// No Office types: the Office writer consumes the validated request via
/// <c>WriteSettings</c> with exactly the <c>ChartTitle</c> and <c>ShowTitle</c> keys.
/// The title carries no length cap: available space is the constraint and the
/// renderer truncates with an ellipsis plus a warning (entity-guide §2), never by
/// silently shrinking the font. The title is always stored non-blank, even when
/// hidden, so <c>ShowTitle=true</c> restores the retained title.
/// </remarks>
/// <param name="Title">The trimmed non-blank chart title.</param>
/// <param name="ShowTitle">Whether the populated title band is visible.</param>
public sealed record ChartTitleEditRequest(
    string Title,
    bool ShowTitle)
{
    /// <summary>
    /// Validates the request: the title must be non-blank after trimming leading
    /// and trailing whitespace.
    /// </summary>
    /// <param name="request">The request to validate.</param>
    /// <param name="refusal">The refusal when invalid.</param>
    /// <param name="message">The user-actionable message when invalid.</param>
    /// <returns>True when the title is non-blank after trimming.</returns>
    public static bool TryValidate(
        ChartTitleEditRequest? request,
        out ChartTitleEditRefusal? refusal,
        out string? message)
    {
        refusal = null;
        message = null;
        if (request is null || string.IsNullOrWhiteSpace(request.Title))
        {
            refusal = ChartTitleEditRefusal.BlankTitle;
            message = "Chart title must not be blank.";
            return false;
        }

        return true;
    }

    /// <summary>
    /// Creates a validated request from raw dialog input, trimming leading and
    /// trailing whitespace before validation.
    /// </summary>
    /// <param name="rawTitle">The raw title text from the dialog.</param>
    /// <param name="showTitle">Whether the populated title band is visible.</param>
    /// <param name="request">The trimmed request when valid.</param>
    /// <param name="refusal">The refusal when invalid.</param>
    /// <param name="message">The user-actionable message when invalid.</param>
    /// <returns>True when the trimmed title is non-blank.</returns>
    public static bool TryCreate(
        string? rawTitle,
        bool showTitle,
        out ChartTitleEditRequest? request,
        out ChartTitleEditRefusal? refusal,
        out string? message)
    {
        request = rawTitle is null
            ? new ChartTitleEditRequest(string.Empty, showTitle)
            : new ChartTitleEditRequest(rawTitle.Trim(), showTitle);
        return TryValidate(request, out refusal, out message);
    }

    /// <summary>
    /// Projects the validated request to the exact settings-table key/value pairs.
    /// </summary>
    /// <param name="request">The validated request.</param>
    /// <returns>Exactly the <c>ChartTitle</c> and <c>ShowTitle</c> pairs.</returns>
    public static IReadOnlyDictionary<string, string> ToSettings(ChartTitleEditRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        return new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["ChartTitle"] = request.Title,
            ["ShowTitle"] = request.ShowTitle ? "TRUE" : "FALSE",
        };
    }
}
