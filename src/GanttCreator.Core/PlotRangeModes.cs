namespace GanttCreator.Core;

/// <summary>
/// The closed plot-range mode values stored in workbook settings
/// (<c>PlotStartMode</c>/<c>PlotFinishMode</c>, R5.1).
/// </summary>
public enum PlotRangeMode
{
    /// <summary>Derive the range end from the validated visible events.</summary>
    DataRange = 0,

    /// <summary>Use the user-supplied explicit date for the range end.</summary>
    Explicit = 1,
}

/// <summary>Strict parsers for the plot-range mode settings.</summary>
public static class PlotRangeModes
{
    /// <summary>Attempts to parse a stored plot-range mode using exact ordinal text.</summary>
    /// <param name="text">The stored setting text.</param>
    /// <param name="mode">The parsed mode when successful.</param>
    /// <returns><see langword="true"/> when the text is a known mode.</returns>
    public static bool TryParse(string? text, out PlotRangeMode mode)
    {
        switch (text)
        {
            case nameof(PlotRangeMode.DataRange):
                mode = PlotRangeMode.DataRange;
                return true;
            case nameof(PlotRangeMode.Explicit):
                mode = PlotRangeMode.Explicit;
                return true;
            default:
                mode = default;
                return false;
        }
    }
}

/// <summary>Why a plot range could not be resolved.</summary>
public enum PlotRangeRefusal
{
    /// <summary>The event collection was null.</summary>
    NullEvents = 0,

    /// <summary>No event carries a start date, so there is no range to draw.</summary>
    NoPlotRange = 1,

    /// <summary>A mode setting names no known mode.</summary>
    UnknownMode = 2,

    /// <summary>Explicit mode names no date text for one end.</summary>
    MissingExplicitDate = 3,

    /// <summary>An explicit date does not parse as a calendar date.</summary>
    UnparsableExplicitDate = 4,

    /// <summary>An explicit date is the default <see cref="DateOnly"/> value.</summary>
    DefaultDates = 5,

    /// <summary>The resolved start is after the resolved finish.</summary>
    StartAfterFinish = 6,
}

/// <summary>The typed result of resolving a plot range.</summary>
/// <param name="Start">The resolved inclusive start when successful.</param>
/// <param name="Finish">The resolved inclusive finish when successful.</param>
/// <param name="Refusal">The typed refusal when unsuccessful.</param>
public sealed record PlotRangeOutcome(DateOnly? Start, DateOnly? Finish, PlotRangeRefusal? Refusal)
{
    /// <summary>Gets whether a range was resolved.</summary>
    public bool Succeeded => Start is not null && Finish is not null;
}
