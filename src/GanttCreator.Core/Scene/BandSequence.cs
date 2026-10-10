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
    /// <param name="periodMeasurer">
    /// The width seam used to fit each period label to its interval, or
    /// <see langword="null"/> to fall back to the fixed floor for every period.
    /// </param>
    /// <returns>A typed sequence outcome.</returns>
    /// <remarks>
    /// <para>
    /// <b>Years keep the fixed floor; periods are fit-gated.</b> A year interval
    /// always spans a whole plot, so its four-digit label is shown whenever the
    /// interval clears <paramref name="minimumLabelWidthPt"/>. A period interval
    /// can be far narrower than its label, so when <paramref name="periodMeasurer"/>
    /// is supplied the period label is shown only when the interval is at least as
    /// wide as the label measures. This is uniform by construction: the format is a
    /// single scene-wide choice, so every shown period label carries the same form
    /// (all <c>MMM</c> or all <c>MM</c>) and only the show/hide decision varies per
    /// interval. <c>MM</c> measures shorter than <c>MMM</c>, so it survives on
    /// narrower intervals -- the intent of exposing the two forms.
    /// </para>
    /// <para>
    /// The measurement uses the same single seam the chart title already uses, so
    /// the period header inherits that seam's residual against its own render size
    /// rather than introducing a second font-size authority. With no measurer, or a
    /// measurement that fails, a period falls back to the floor, so fit-gating
    /// activates only when a measurement actually succeeds and a missing seam can
    /// never silently open the gate.
    /// </para>
    /// </remarks>
    public static BandSequenceCreationOutcome TryCreate(
        TimeScale? timeScale,
        GanttTimeScale scale,
        GanttPeriodLabelFormat format,
        RectD plotBounds,
        double minimumLabelWidthPt,
        ITextWidthMeasurer? periodMeasurer = null
    )
    {
        return timeScale is not { } validScale ? Refused(BandSequenceRefusal.NullTimeScale)
            : !GanttChartSettings.IsCompatible(scale, format) ? Refused(BandSequenceRefusal.IncompatibleSettings)
            : !IsValidPlotBounds(plotBounds, minimumLabelWidthPt) ? Refused(BandSequenceRefusal.InvalidGeometry)
            : new BandSequenceCreationOutcome(
                new BandSequence(
                    CreateYears(validScale, minimumLabelWidthPt),
                    CreatePeriods(validScale, scale, format, minimumLabelWidthPt, periodMeasurer)
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
        double minimumWidth,
        ITextWidthMeasurer? periodMeasurer
    )
    {
        List<BandInterval> intervals = [];
        DateOnly periodStart = periodScale switch
        {
            GanttTimeScale.Month => new(scale.PlotStart.Year, scale.PlotStart.Month, 1),
            GanttTimeScale.Quarter => new(scale.PlotStart.Year, ((scale.PlotStart.Month - 1) / 3 * 3) + 1, 1),
            GanttTimeScale.Week => StartOfIsoWeek(scale.PlotStart),
            _ => throw new ArgumentOutOfRangeException(nameof(periodScale)),
        };

        while (periodStart <= scale.PlotFinish)
        {
            DateOnly next = periodScale switch
            {
                GanttTimeScale.Month => periodStart.AddMonths(1),
                GanttTimeScale.Quarter => periodStart.AddMonths(3),
                GanttTimeScale.Week => periodStart.AddDays(7),
                _ => throw new ArgumentOutOfRangeException(nameof(periodScale)),
            };
            DateOnly periodFinish = next.AddDays(-1);
            var label = FormatPeriod(periodStart, format);
            AddInterval(intervals, scale, periodStart, periodFinish, label, RequiredPeriodWidth(label, minimumWidth, periodMeasurer));
            periodStart = next;
        }

        return intervals;
    }

    /// <summary>
    /// The minimum interval width a period label needs to be shown: its measured
    /// width when a measurer is supplied and succeeds, otherwise the fixed floor.
    /// </summary>
    /// <param name="label">The formatted period label.</param>
    /// <param name="floorWidthPt">The fixed suppression floor.</param>
    /// <param name="measurer">The optional width seam.</param>
    /// <returns>The required interval width in points.</returns>
    /// <remarks>
    /// A non-finite or negative measurement is treated as a failure and falls back
    /// to the floor, matching <see cref="ITextWidthMeasurer"/>'s own contract that a
    /// successful measure is finite and non-negative. This keeps a misbehaving seam
    /// from either opening the gate (a zero/negative width shows every label) or
    /// closing it (an infinite width suppresses every label).
    /// </remarks>
    private static double RequiredPeriodWidth(string label, double floorWidthPt, ITextWidthMeasurer? measurer) =>
        measurer is not null && measurer.TryMeasure(label, out var widthPt) && double.IsFinite(widthPt) && widthPt >= 0
            ? widthPt
            : floorWidthPt;

    private static void AddInterval(
        List<BandInterval> intervals,
        TimeScale scale,
        DateOnly periodStart,
        DateOnly periodFinish,
        string label,
        double requiredWidth
    )
    {
        DateOnly visibleStart = periodStart < scale.PlotStart ? scale.PlotStart : periodStart;
        DateOnly visibleFinish = periodFinish > scale.PlotFinish ? scale.PlotFinish : periodFinish;
        var left = visibleStart == scale.PlotStart ? scale.PlotLeftPt : scale.DateToX(visibleStart);
        var right = visibleFinish == scale.PlotFinish ? scale.PlotRightPt : scale.DateToX(periodFinish.AddDays(1));
        intervals.Add(new BandInterval(visibleStart, visibleFinish, left, right, label, right - left >= requiredWidth));
    }

    private static string FormatPeriod(DateOnly start, GanttPeriodLabelFormat format) =>
        format switch
        {
            GanttPeriodLabelFormat.MM => start.ToString("MM", CultureInfo.InvariantCulture),
            GanttPeriodLabelFormat.MMM => start.ToString("MMM", CultureInfo.InvariantCulture),
            GanttPeriodLabelFormat.Quarter => $"Q{((start.Month - 1) / 3) + 1}",
            GanttPeriodLabelFormat.Week => $"W{ISOWeek.GetWeekOfYear(start):D2}",
            _ => throw new ArgumentOutOfRangeException(nameof(format)),
        };

    /// <summary>
    /// The Monday that begins the ISO week containing <paramref name="date"/>.
    /// </summary>
    /// <param name="date">Any date inside the week.</param>
    /// <returns>The Monday on or before <paramref name="date"/>.</returns>
    /// <remarks>
    /// ISO weeks begin on Monday, so the first band is clipped back to that Monday
    /// and every following band steps a whole number of days from it. The
    /// <see cref="ISOWeek"/> numbering is culture-invariant, matching the rest of
    /// the sequence.
    /// </remarks>
    private static DateOnly StartOfIsoWeek(DateOnly date)
    {
        var offset = ((int)date.DayOfWeek - (int)DayOfWeek.Monday + 7) % 7;
        return date.AddDays(-offset);
    }

    private static BandSequenceCreationOutcome Refused(BandSequenceRefusal refusal) => new(null, refusal);
}
