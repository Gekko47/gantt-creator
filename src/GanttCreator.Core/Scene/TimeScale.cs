namespace GanttCreator.Core.Scene;

/// <summary>A deterministic, culture-free mapping from calendar dates to point X coordinates.</summary>
public sealed class TimeScale
{
    private TimeScale(
        DateOnly plotStart,
        DateOnly plotFinish,
        double plotLeftPt,
        double plotRightPt,
        int totalPlotDays,
        double dayWidth)
    {
        PlotStart = plotStart;
        PlotFinish = plotFinish;
        PlotLeftPt = plotLeftPt;
        PlotRightPt = plotRightPt;
        TotalPlotDays = totalPlotDays;
        DayWidth = dayWidth;
    }

    /// <summary>Gets the inclusive plot start date.</summary>
    public DateOnly PlotStart { get; }

    /// <summary>Gets the inclusive plot finish date.</summary>
    public DateOnly PlotFinish { get; }

    /// <summary>Gets the plot left boundary in points.</summary>
    public double PlotLeftPt { get; }

    /// <summary>Gets the plot right boundary in points.</summary>
    public double PlotRightPt { get; }

    /// <summary>Gets the inclusive number of calendar days in the plot.</summary>
    public int TotalPlotDays { get; }

    /// <summary>Gets the point width of one calendar day.</summary>
    public double DayWidth { get; }

    /// <summary>Attempts to create an immutable time scale.</summary>
    /// <param name="plotStart">The inclusive plot start date.</param>
    /// <param name="plotFinish">The inclusive plot finish date.</param>
    /// <param name="plotLeftPt">The plot left boundary in points.</param>
    /// <param name="plotRightPt">The plot right boundary in points.</param>
    /// <returns>A typed creation outcome.</returns>
    public static TimeScaleCreationOutcome TryCreate(
        DateOnly plotStart,
        DateOnly plotFinish,
        double plotLeftPt,
        double plotRightPt)
    {
        if (plotStart == default || plotFinish == default)
        {
            return new TimeScaleCreationOutcome(null, TimeScaleCreationRefusal.DefaultDates);
        }

        if (plotStart > plotFinish)
        {
            return new TimeScaleCreationOutcome(null, TimeScaleCreationRefusal.StartAfterFinish);
        }

        if (!double.IsFinite(plotLeftPt) || !double.IsFinite(plotRightPt) || plotRightPt <= plotLeftPt)
        {
            return new TimeScaleCreationOutcome(null, TimeScaleCreationRefusal.NonFiniteOrDegenerateWidth);
        }

        var totalPlotDays = plotFinish.DayNumber - plotStart.DayNumber + 1;
        var dayWidth = (plotRightPt - plotLeftPt) / totalPlotDays;
        return !double.IsFinite(dayWidth) || dayWidth <= 0
            ? new TimeScaleCreationOutcome(null, TimeScaleCreationRefusal.NonFiniteOrDegenerateWidth)
            : new TimeScaleCreationOutcome(
                new TimeScale(plotStart, plotFinish, plotLeftPt, plotRightPt, totalPlotDays, dayWidth),
                null);
    }

    /// <summary>Attempts to map a date to its start-of-day point X coordinate.</summary>
    /// <param name="date">The date to map.</param>
    /// <param name="x">The mapped point X coordinate when successful.</param>
    /// <returns><see langword="true"/> when the date is inside the inclusive plot range.</returns>
    public bool TryDateToX(DateOnly date, out double x)
    {
        var inRange = date >= PlotStart && date <= PlotFinish;
        x = inRange
            ? PlotLeftPt + ((date.DayNumber - PlotStart.DayNumber) * DayWidth)
            : 0;
        return inRange;
    }

    /// <summary>Maps a date to its start-of-day point X coordinate.</summary>
    /// <param name="date">The date to map.</param>
    /// <returns>The point X coordinate.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the date is outside the inclusive plot range.</exception>
    public double DateToX(DateOnly date) =>
        TryDateToX(date, out var x)
            ? x
            : throw new ArgumentOutOfRangeException(nameof(date), date, "Date is outside the time-scale range.");

    /// <summary>Calculates inclusive activity duration in calendar days.</summary>
    /// <param name="start">The inclusive activity start date.</param>
    /// <param name="finish">The inclusive activity finish date.</param>
    /// <returns>The inclusive duration in days.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when either date is outside the scale or finish is before start.</exception>
    public int DurationDays(DateOnly start, DateOnly finish) =>
        TryDurationDays(start, finish, out var durationDays)
            ? durationDays
            : throw new ArgumentOutOfRangeException(nameof(finish), finish, "Activity dates must be ordered and inside the time-scale range.");

    /// <summary>Attempts to calculate inclusive activity duration in calendar days.</summary>
    /// <param name="start">The inclusive activity start date.</param>
    /// <param name="finish">The inclusive activity finish date.</param>
    /// <param name="durationDays">The inclusive duration when successful.</param>
    /// <returns><see langword="true"/> when the activity dates are valid.</returns>
    public bool TryDurationDays(DateOnly start, DateOnly finish, out int durationDays)
    {
        var valid = start >= PlotStart && start <= PlotFinish &&
            finish >= PlotStart && finish <= PlotFinish && finish >= start;
        durationDays = valid ? finish.DayNumber - start.DayNumber + 1 : 0;
        return valid;
    }

    /// <summary>Calculates the duration-derived right edge of an activity.</summary>
    /// <param name="start">The inclusive activity start date.</param>
    /// <param name="finish">The inclusive activity finish date.</param>
    /// <returns>The activity right edge in points.</returns>
    public double ActivityRightX(DateOnly start, DateOnly finish) =>
        TryDurationDays(start, finish, out var durationDays)
            ? DateToX(start) + (durationDays * DayWidth)
            : throw new ArgumentOutOfRangeException(nameof(finish), finish, "Activity dates must be ordered and inside the time-scale range.");

    /// <summary>Maps a point-event date to its exact start-of-day point X coordinate.</summary>
    /// <param name="date">The point-event date.</param>
    /// <returns>The point X coordinate.</returns>
    public double PointEventX(DateOnly date) => DateToX(date);
}
