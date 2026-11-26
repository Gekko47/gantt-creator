using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// The plot-range and plot-layout settings authority: the closed reads of one
/// workbook settings map, the stored-date normalisation, the explicit-range
/// inversion checks, and the derivation of the two effective display dates.
/// </summary>
/// <remarks>
/// Every member is a pure read or a pure transition over <see cref="RibbonState"/>.
/// The setters and the persistence order stay on <see cref="RibbonStateService"/>,
/// which owns the state field, the writer and the refresh scheduling; this type
/// answers the questions those paths ask.
/// </remarks>
internal static class PlotRangeProjection
{
    /// <summary>
    /// The one canonical period format for a scale: MMM for months, Qn for
    /// quarters, Wnn for weeks. Used when a stored format is absent, names an
    /// unknown value, or is incompatible with the scale it is stored against.
    /// </summary>
    /// <param name="scale">The scale the format is displayed against.</param>
    /// <returns>The scale's canonical format.</returns>
    internal static GanttPeriodLabelFormat CanonicalPeriodLabelFormat(GanttTimeScale scale) => scale switch
    {
        GanttTimeScale.Month => GanttPeriodLabelFormat.MMM,
        GanttTimeScale.Quarter => GanttPeriodLabelFormat.Quarter,
        GanttTimeScale.Week => GanttPeriodLabelFormat.Week,
        _ => GanttPeriodLabelFormat.MMM,
    };

    /// <summary>
    /// Reads the stored period label format, falling back to the scale's
    /// canonical form when the key is absent, names an unknown format, or is
    /// incompatible with the scale it is stored against (R5.12).
    /// </summary>
    /// <param name="settings">The settings map just loaded from the workbook.</param>
    /// <param name="scale">The parsed time scale.</param>
    /// <returns>The stored format when compatible, otherwise the canonical one.</returns>
    internal static GanttPeriodLabelFormat ReadPeriodLabelFormat(
        IReadOnlyDictionary<string, string> settings,
        GanttTimeScale scale) =>
        settings.TryGetValue("PeriodLabelFormat", out string? text)
        && GanttChartSettings.TryParsePeriodLabelFormat(text, out GanttPeriodLabelFormat parsed)
        && GanttChartSettings.IsCompatible(scale, parsed)
            ? parsed
            : CanonicalPeriodLabelFormat(scale);

    /// <summary>
    /// Reads the stored plot time scale, falling back to the catalogue default
    /// when the key is absent or names an unknown scale (R5.2).
    /// </summary>
    /// <param name="settings">The settings map just loaded from the workbook.</param>
    /// <returns>The stored scale, or the month default.</returns>
    internal static GanttTimeScale ReadTimeScale(IReadOnlyDictionary<string, string> settings) =>
        settings.TryGetValue("TimeScale", out string? text)
            && GanttChartSettings.TryParseTimeScale(text, out GanttTimeScale scale)
            ? scale
            : GanttTimeScale.Month;

    /// <summary>
    /// Reads the stored plot margin preset, falling back to the Windows
    /// "Normal" default when the key is absent or names an unknown preset (R5.2).
    /// </summary>
    /// <param name="settings">The settings map just loaded from the workbook.</param>
    /// <returns>The stored preset, or the default.</returns>
    internal static GanttPlotMargin ReadMargin(IReadOnlyDictionary<string, string> settings) =>
        settings.TryGetValue("Margin", out string? text)
            && GanttPlotMargins.TryParse(text, out GanttPlotMargin margin)
            ? margin
            : GanttPlotMargins.Default;

    /// <summary>
    /// Reads the stored custom margin, in centimetres (R5.2). An absent or
    /// unparsable value falls back to the catalogue default.
    /// </summary>
    /// <param name="settings">The settings map just loaded from the workbook.</param>
    /// <returns>The stored custom margin in centimetres, or the default.</returns>
    internal static double ReadMarginCm(IReadOnlyDictionary<string, string> settings) =>
        settings.TryGetValue("MarginCm", out string? raw)
            && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            && double.IsFinite(parsed)
            ? parsed
            : GanttPlotMargins.DefaultCustomCm;

    /// <summary>
    /// Reads the stored output size preset, returning null only when the key is absent or
    /// names an unknown preset (R5.2, R8.9).
    /// </summary>
    /// <param name="settings">The settings map just loaded from the workbook.</param>
    /// <returns>The stored preset, or null.</returns>
    /// <remarks>
    /// A Custom key always resolves. The width is clamped into the band rather than
    /// validated, because a refusal here would leave the ribbon showing no preset while the
    /// workbook still says Custom, and the refresh would refuse with no way to correct it
    /// from the Size control.
    /// </remarks>
    internal static SizePreset? ReadPreset(IReadOnlyDictionary<string, string> settings) =>
        settings.TryGetValue("SizePreset", out string? text)
            ? string.Equals(text, SizePresets.CustomKey, StringComparison.Ordinal)
                ? ReadCustomPreset(settings)
                : SizePresets.ByKey(text)
            : null;

    private static SizePreset ReadCustomPreset(IReadOnlyDictionary<string, string> settings) =>
        SizePresets.Custom(CustomSizeCm.FromClampedWidth(ReadSizeWidthCm(settings)));

    /// <summary>Reads the stored custom width in centimetres (R8.9).</summary>
    /// <param name="settings">The settings map just loaded from the workbook.</param>
    /// <returns>The stored width, or the A4-landscape default.</returns>
    internal static double ReadSizeWidthCm(IReadOnlyDictionary<string, string> settings) =>
        ReadFinite(settings, "SizeWidthCm", CustomSizeCm.DefaultWidthCm);

    /// <summary>Reads the stored live-height mirror in centimetres (R8.9).</summary>
    /// <param name="settings">The settings map just loaded from the workbook.</param>
    /// <returns>The stored mirror, or zero when absent. Zero means "not measured yet".</returns>
    /// <remarks>
    /// This is the mirror the ribbon writes back after measuring the plot, never an input.
    /// It exists so a reopened workbook can show the last known height before the first
    /// measurement of the session.
    /// </remarks>
    internal static double ReadSizeHeightCm(IReadOnlyDictionary<string, string> settings) =>
        ReadFinite(settings, "SizeHeightCm", 0);

    private static double ReadFinite(
        IReadOnlyDictionary<string, string> settings,
        string key,
        double fallback) =>
        settings.TryGetValue(key, out string? raw)
            && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            && double.IsFinite(parsed)
            ? parsed
            : fallback;

    /// <summary>
    /// Normalises one stored explicit plot date for display and state: a
    /// parseable value becomes canonical dd-MMM-yy; blank stays blank; an
    /// unparsable value is kept verbatim so the failure stays visible.
    /// </summary>
    /// <param name="stored">The stored date text.</param>
    /// <returns>The display-ready date text.</returns>
    internal static string NormaliseStoredPlotDate(string stored)
    {
        return string.IsNullOrWhiteSpace(stored)
            ? string.Empty
            : PlotRangeResolver.TryParsePlotDate(stored, out DateOnly parsed)
                ? GanttDateFormatting.FormatDdMMMyy(parsed)
                : stored;
    }

    /// <summary>
    /// Reads the automatic-mode range padding, mirroring the Refresh pipeline's
    /// default: a stored value outside the valid range falls back to 3.
    /// </summary>
    /// <param name="settings">The settings map just loaded from the workbook.</param>
    /// <returns>The non-negative padding in days per side.</returns>
    internal static int ReadRangePaddingDays(IReadOnlyDictionary<string, string> settings)
    {
        return settings.TryGetValue("RangePaddingDays", out string? text)
            && int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
            && parsed >= 0
            ? parsed
            : 3;
    }

    /// <summary>
    /// Accepts a candidate explicit plot-start date, or refuses it because it
    /// would invert the range.
    /// </summary>
    /// <param name="state">The snapshot the finish end is read from.</param>
    /// <param name="normalisedStart">The candidate start date, formatted for storage.</param>
    /// <param name="warning">The user-facing message when the candidate is refused.</param>
    /// <returns>Whether the candidate may be committed.</returns>
    internal static bool TryAcceptExplicitStart(
        RibbonState state,
        string normalisedStart,
        out string? warning)
    {
        warning = null;

        if (!PlotRangeResolver.TryParsePlotDate(state.EffectivePlotFinishDate, out DateOnly finish))
        {
            return true;
        }

        if (!PlotRangeResolver.TryParsePlotDate(normalisedStart, out DateOnly start))
        {
            return true;
        }

        if (start <= finish)
        {
            return true;
        }

        warning =
            $"The plot start {start:dd-MMM-yy} is after the plot finish {finish:dd-MMM-yy}. "
            + "The chart cannot be drawn over that range, so the start date has been kept.";
        return false;
    }

    /// <summary>
    /// Accepts a candidate explicit plot-finish date, or refuses it because it
    /// would invert the range. The symmetric counterpart of
    /// <see cref="TryAcceptExplicitStart"/>.
    /// </summary>
    /// <param name="state">The snapshot the start end is read from.</param>
    /// <param name="normalisedFinish">The candidate finish date, formatted for storage.</param>
    /// <param name="warning">The user-facing message when the candidate is refused.</param>
    /// <returns>Whether the candidate may be committed.</returns>
    internal static bool TryAcceptExplicitFinish(
        RibbonState state,
        string normalisedFinish,
        out string? warning)
    {
        warning = null;

        if (!PlotRangeResolver.TryParsePlotDate(state.EffectivePlotStartDate, out DateOnly start))
        {
            return true;
        }

        if (!PlotRangeResolver.TryParsePlotDate(normalisedFinish, out DateOnly finish))
        {
            return true;
        }

        if (finish >= start)
        {
            return true;
        }

        warning =
            $"The plot finish {finish:dd-MMM-yy} is before the plot start {start:dd-MMM-yy}. "
            + "The chart cannot be drawn over that range, so the finish date has been kept.";
        return false;
    }

    /// <summary>
    /// Derives the effective display dates for both edit boxes (fix plan ruling 1):
    /// the explicit date in explicit mode, the TRUE chart plot bound in automatic
    /// mode, or empty when neither is known. A failed derivation keeps the last
    /// known display value rather than blanking the box.
    /// </summary>
    /// <param name="state">The snapshot being projected.</param>
    /// <param name="settings">The effective settings map just loaded.</param>
    /// <param name="readTable">Reads the live table; called inside the guard.</param>
    /// <param name="today">The caller's current day.</param>
    /// <returns>The two display dates; the current ones when anything fails.</returns>
    internal static (string Start, string Finish) ResolveEffectiveDates(
        RibbonState state,
        IReadOnlyDictionary<string, string> settings,
        Func<GanttTableReadOutcome?> readTable,
        DateOnly today)
    {
        ArgumentNullException.ThrowIfNull(settings);
        ArgumentNullException.ThrowIfNull(readTable);

        // CA1031: runs inside Ribbon callbacks and Excel events; any failure
        // keeps the last known display values.
#pragma warning disable CA1031
        try
        {
            string explicitStart = state.PlotStartAuto ? string.Empty : state.PlotStartDate;
            string explicitFinish = state.PlotFinishAuto ? string.Empty : state.PlotFinishDate;

            string effectiveStart = explicitStart;
            string effectiveFinish = explicitFinish;

            if (state.PlotStartAuto || state.PlotFinishAuto)
            {
                GanttTableReadOutcome? table = readTable();
                if (table is { Succeeded: true } && table.Rows is not null)
                {
                    GanttValidationOutcome validation = GanttRowValidator.Validate(table.Rows);
                    int padding = ReadRangePaddingDays(settings);
                    PlotRangeOutcome range = PlotRangeResolver.TryResolve(
                        validation.Events,
                        state.PlotStartAuto ? nameof(PlotRangeMode.DataRange) : nameof(PlotRangeMode.Explicit),
                        state.PlotFinishAuto ? nameof(PlotRangeMode.DataRange) : nameof(PlotRangeMode.Explicit),
                        state.PlotStartDate,
                        state.PlotFinishDate,
                        padding,
                        today: today);

                    if (range.Succeeded && range.Start is { } resolvedStart && range.Finish is { } resolvedFinish)
                    {
                        // The TRUE chart bounds: per-end padding-then-snap through the
                        // single PlotMonthBounds authority Refresh renders (owner ruling
                        // 2026-10-02). DataRange ends are month-snapped; explicit ends
                        // pass through verbatim.
                        PlotRangeMode startMode = state.PlotStartAuto ? PlotRangeMode.DataRange : PlotRangeMode.Explicit;
                        PlotRangeMode finishMode = state.PlotFinishAuto ? PlotRangeMode.DataRange : PlotRangeMode.Explicit;

                        DateOnly start = resolvedStart;
                        DateOnly finish = resolvedFinish;

                        if (startMode == PlotRangeMode.DataRange)
                        {
                            start = PlotMonthBounds.SnapToMonthStart(start);
                        }

                        if (finishMode == PlotRangeMode.DataRange)
                        {
                            finish = PlotMonthBounds.SnapToMonthEnd(finish);
                        }

                        // Degenerate guard: if snap moved finish on or before start,
                        // keep the last display values.
                        if (finish.DayNumber <= start.DayNumber)
                        {
                            return (state.EffectivePlotStartDate, state.EffectivePlotFinishDate);
                        }

                        if (state.PlotStartAuto)
                        {
                            effectiveStart = GanttDateFormatting.FormatDdMMMyy(start);
                        }

                        if (state.PlotFinishAuto)
                        {
                            effectiveFinish = GanttDateFormatting.FormatDdMMMyy(finish);
                        }
                    }
                }
            }

            return (effectiveStart, effectiveFinish);
        }
        catch
        {
            // Degrade gracefully: the last known display values stand.
            return (state.EffectivePlotStartDate, state.EffectivePlotFinishDate);
        }
#pragma warning restore CA1031
    }
}
