using System.Globalization;

namespace GanttCreator.Core.Scene;

/// <summary>The reason a band sequence could not be created.</summary>
public enum BandSequenceRefusal
{
    /// <summary>The time scale was null.</summary>
    NullTimeScale = 0,

    /// <summary>The scale and label format were incompatible.</summary>
    IncompatibleSettings = 1,

    /// <summary>The plot geometry or minimum label width was invalid.</summary>
    InvalidGeometry = 2,
}

/// <summary>The typed result of creating clipped year and period sequences.</summary>
/// <param name="Sequence">The successful sequence, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record BandSequenceCreationOutcome(BandSequence? Sequence, BandSequenceRefusal? Refusal)
{
    /// <summary>Gets whether sequence creation succeeded.</summary>
    public bool Succeeded => Sequence is not null;
}

/// <summary>Immutable clipped year and period sequences sharing exact TimeScale X edges.</summary>
/// <param name="Years">The year intervals.</param>
/// <param name="Periods">The scale-period intervals.</param>
public sealed record BandSequence(IReadOnlyList<BandInterval> Years, IReadOnlyList<BandInterval> Periods)
{
    /// <summary>Attempts to create year and scale-period sequences.</summary>
    /// <param name="timeScale">The validated date-to-X scale.</param>
    /// <param name="scale">The selected calendar scale.</param>
    /// <param name="format">The selected period label format.</param>
    /// <param name="plotBounds">The measured plot bounds.</param>
    /// <param name="minimumLabelWidthPt">The minimum visible label width.</param>
    /// <returns>A typed sequence outcome.</returns>
    public static BandSequenceCreationOutcome TryCreate(
        TimeScale? timeScale,
        GanttTimeScale scale,
        GanttPeriodLabelFormat format,
        RectD plotBounds,
        double minimumLabelWidthPt
    )
    {
        return timeScale is not { } validScale ? Refused(BandSequenceRefusal.NullTimeScale)
            : !GanttChartSettings.IsCompatible(scale, format) ? Refused(BandSequenceRefusal.IncompatibleSettings)
            : !IsValidPlotBounds(plotBounds, minimumLabelWidthPt) ? Refused(BandSequenceRefusal.InvalidGeometry)
            : new BandSequenceCreationOutcome(
                new BandSequence(
                    CreateYears(validScale, minimumLabelWidthPt),
                    CreatePeriods(validScale, scale, format, minimumLabelWidthPt)
                ),
                null
            );
    }

    private static bool IsValidPlotBounds(RectD plotBounds, double minimumLabelWidthPt) =>
        plotBounds.Width > 0 && plotBounds.Height > 0 && double.IsFinite(minimumLabelWidthPt) && minimumLabelWidthPt >= 0;

    private static List<BandInterval> CreateYears(TimeScale scale, double minimumWidth)
    {
        List<BandInterval> intervals = [];
        DateOnly periodStart = new(scale.PlotStart.Year, 1, 1);
        while (periodStart <= scale.PlotFinish)
        {
            DateOnly periodFinish = periodStart.AddYears(1).AddDays(-1);
            AddInterval(
                intervals,
                scale,
                periodStart,
                periodFinish,
                periodStart.ToString("yyyy", CultureInfo.InvariantCulture),
                minimumWidth
            );
            periodStart = periodStart.AddYears(1);
        }

        return intervals;
    }

    private static List<BandInterval> CreatePeriods(
        TimeScale scale,
        GanttTimeScale periodScale,
        GanttPeriodLabelFormat format,
        double minimumWidth
    )
    {
        List<BandInterval> intervals = [];
        DateOnly periodStart = periodScale switch
        {
            GanttTimeScale.Month => new(scale.PlotStart.Year, scale.PlotStart.Month, 1),
            GanttTimeScale.Quarter => new(scale.PlotStart.Year, ((scale.PlotStart.Month - 1) / 3 * 3) + 1, 1),
            GanttTimeScale.Year => new(scale.PlotStart.Year, 1, 1),
            _ => throw new ArgumentOutOfRangeException(nameof(periodScale)),
        };

        while (periodStart <= scale.PlotFinish)
        {
            DateOnly next = periodScale switch
            {
                GanttTimeScale.Month => periodStart.AddMonths(1),
                GanttTimeScale.Quarter => periodStart.AddMonths(3),
                GanttTimeScale.Year => periodStart.AddYears(1),
                _ => throw new ArgumentOutOfRangeException(nameof(periodScale)),
            };
            DateOnly periodFinish = next.AddDays(-1);
            AddInterval(intervals, scale, periodStart, periodFinish, FormatPeriod(periodStart, format), minimumWidth);
            periodStart = next;
        }

        return intervals;
    }

    private static void AddInterval(
        List<BandInterval> intervals,
        TimeScale scale,
        DateOnly periodStart,
        DateOnly periodFinish,
        string label,
        double minimumWidth
    )
    {
        DateOnly visibleStart = periodStart < scale.PlotStart ? scale.PlotStart : periodStart;
        DateOnly visibleFinish = periodFinish > scale.PlotFinish ? scale.PlotFinish : periodFinish;
        double left = visibleStart == scale.PlotStart ? scale.PlotLeftPt : scale.DateToX(visibleStart);
        double right = visibleFinish == scale.PlotFinish ? scale.PlotRightPt : scale.DateToX(periodFinish.AddDays(1));
        intervals.Add(new BandInterval(visibleStart, visibleFinish, left, right, label, right - left >= minimumWidth));
    }

    private static string FormatPeriod(DateOnly start, GanttPeriodLabelFormat format) =>
        format switch
        {
            GanttPeriodLabelFormat.MM => start.ToString("MM", CultureInfo.InvariantCulture),
            GanttPeriodLabelFormat.MMM => start.ToString("MMM", CultureInfo.InvariantCulture),
            GanttPeriodLabelFormat.Quarter => $"Q{((start.Month - 1) / 3) + 1}",
            GanttPeriodLabelFormat.Year => start.ToString("yyyy", CultureInfo.InvariantCulture),
            _ => throw new ArgumentOutOfRangeException(nameof(format)),
        };

    private static BandSequenceCreationOutcome Refused(BandSequenceRefusal refusal) => new(null, refusal);
}
