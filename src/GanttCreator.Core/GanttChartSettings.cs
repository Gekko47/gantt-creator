namespace GanttCreator.Core;

/// <summary>
/// The closed time-scale values stored in workbook settings. Each selects the
/// calendar unit of the LOWER (period) band; the UPPER band is always the year.
/// </summary>
/// <remarks>
/// <para>
/// <c>Year</c> was removed (owner ruling, R5.2): the upper band is fixed as the
/// year, so a <c>Year</c> lower band would duplicate it. <c>Week</c> was added in
/// its place. The three values are the closed set; any other value is a typed
/// validation error, never a fallback.
/// </para>
/// </remarks>
public enum GanttTimeScale
{
    /// <summary>One calendar month per period.</summary>
    Month = 0,

    /// <summary>One calendar quarter per period.</summary>
    Quarter = 1,

    /// <summary>One ISO week per period.</summary>
    Week = 2,
}

/// <summary>The closed period-label formats stored in workbook settings.</summary>
/// <remarks>
/// The <c>Year</c> member was removed alongside the <c>Year</c> time scale (the
/// year band is the fixed upper band and never carries a period-format setting).
/// <c>Week</c> renders the two-digit ISO week number prefixed <c>W</c> (owner
/// ruling, R5.2).
/// </remarks>
public enum GanttPeriodLabelFormat
{
    /// <summary>Two-digit numeric month, such as <c>01</c>.</summary>
    MM = 0,

    /// <summary>Invariant three-character month, such as <c>Jan</c>.</summary>
    MMM = 1,

    /// <summary>Calendar quarter, such as <c>Q1</c>.</summary>
    Quarter = 2,

    /// <summary>Two-digit ISO week number prefixed <c>W</c>, such as <c>W07</c>.</summary>
    Week = 3,
}

/// <summary>Strict parsers and compatibility rules for the frame/band settings.</summary>
public static class GanttChartSettings
{
    /// <summary>Attempts to parse a stored time scale using exact ordinal text.</summary>
    /// <param name="text">The stored setting text.</param>
    /// <param name="scale">The parsed scale when successful.</param>
    /// <returns><see langword="true"/> when the text is a known scale.</returns>
    public static bool TryParseTimeScale(string? text, out GanttTimeScale scale)
    {
        switch (text)
        {
            case nameof(GanttTimeScale.Month):
                scale = GanttTimeScale.Month;
                return true;
            case nameof(GanttTimeScale.Quarter):
                scale = GanttTimeScale.Quarter;
                return true;
            case nameof(GanttTimeScale.Week):
                scale = GanttTimeScale.Week;
                return true;
            default:
                scale = default;
                return false;
        }
    }

    /// <summary>Attempts to parse a stored period format using exact ordinal text.</summary>
    /// <param name="text">The stored setting text.</param>
    /// <param name="format">The parsed format when successful.</param>
    /// <returns><see langword="true"/> when the text is a known format.</returns>
    public static bool TryParsePeriodLabelFormat(string? text, out GanttPeriodLabelFormat format)
    {
        switch (text)
        {
            case nameof(GanttPeriodLabelFormat.MM):
                format = GanttPeriodLabelFormat.MM;
                return true;
            case nameof(GanttPeriodLabelFormat.MMM):
                format = GanttPeriodLabelFormat.MMM;
                return true;
            case nameof(GanttPeriodLabelFormat.Quarter):
                format = GanttPeriodLabelFormat.Quarter;
                return true;
            case nameof(GanttPeriodLabelFormat.Week):
                format = GanttPeriodLabelFormat.Week;
                return true;
            default:
                format = default;
                return false;
        }
    }

    /// <summary>Attempts to parse a stored event-date format using exact ordinal text.</summary>
    /// <param name="text">The stored setting text.</param>
    /// <param name="format">The parsed format when successful.</param>
    /// <returns><see langword="true"/> when the text is a known format.</returns>
    /// <remarks>
    /// The set is closed (ADR-0016 D3), so an unrecognised stored value is
    /// refused rather than coerced or defaulted: a silently wrong date format
    /// would change every rendered date without any visible error.
    /// </remarks>
    public static bool TryParseDateDisplayFormat(string? text, out GanttDateDisplayFormat format)
    {
        switch (text)
        {
            case nameof(GanttDateDisplayFormat.DdMMyyyy):
                format = GanttDateDisplayFormat.DdMMyyyy;
                return true;
            case nameof(GanttDateDisplayFormat.DdMMMyy):
                format = GanttDateDisplayFormat.DdMMMyy;
                return true;
            default:
                format = default;
                return false;
        }
    }

    /// <summary>Determines whether a scale and period format are a permitted pair.</summary>
    /// <param name="scale">The stored time scale.</param>
    /// <param name="format">The stored period format.</param>
    /// <returns><see langword="true"/> when the pair is valid.</returns>
    public static bool IsCompatible(GanttTimeScale scale, GanttPeriodLabelFormat format) =>
        scale switch
        {
            GanttTimeScale.Month => format is GanttPeriodLabelFormat.MM or GanttPeriodLabelFormat.MMM,
            GanttTimeScale.Quarter => format == GanttPeriodLabelFormat.Quarter,
            GanttTimeScale.Week => format == GanttPeriodLabelFormat.Week,
            _ => false,
        };
}
