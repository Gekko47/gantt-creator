namespace GanttCreator.Core;

/// <summary>The closed time-scale values stored in workbook settings.</summary>
public enum GanttTimeScale
{
    /// <summary>One calendar month per period.</summary>
    Month = 0,

    /// <summary>One calendar quarter per period.</summary>
    Quarter = 1,

    /// <summary>One calendar year per period.</summary>
    Year = 2,
}

/// <summary>The closed period-label formats stored in workbook settings.</summary>
public enum GanttPeriodLabelFormat
{
    /// <summary>Two-digit numeric month, such as <c>01</c>.</summary>
    MM = 0,

    /// <summary>Invariant three-character month, such as <c>Jan</c>.</summary>
    MMM = 1,

    /// <summary>Calendar quarter, such as <c>Q1</c>.</summary>
    Quarter = 2,

    /// <summary>Four-digit calendar year.</summary>
    Year = 3,
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
            case nameof(GanttTimeScale.Year):
                scale = GanttTimeScale.Year;
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
            case nameof(GanttPeriodLabelFormat.Year):
                format = GanttPeriodLabelFormat.Year;
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
            GanttTimeScale.Year => format == GanttPeriodLabelFormat.Year,
            _ => false,
        };
}
