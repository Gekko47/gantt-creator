using System.Globalization;

namespace GanttCreator.Core;

/// <summary>
/// Resolves the raw inclusive plot-range extent from per-end modes (R5.1 D1/D2).
/// </summary>
/// <remarks>
/// <para>
/// Each end resolves from its own mode: <see cref="PlotRangeMode.DataRange"/>
/// derives from the validated events, <see cref="PlotRangeMode.Explicit"/>
/// parses the user-supplied date text. Modes never blend — an explicit end is
/// never padded, widened, or narrowed here, and an unknown mode or a bad
/// explicit date is a typed refusal rather than a silent fall back to the data.
/// </para>
/// <para>
/// The returned extent is RAW: month-snapping and the degenerate widening stay
/// with the single downstream consumer, so both modes flow through the same
/// pipeline and cannot disagree about what "the range" means.
/// </para>
/// </remarks>
public static class PlotRangeResolver
{
    /// <summary>
    /// Attempts to resolve the raw inclusive plot-range extent.
    /// </summary>
    /// <param name="events">The validated events to derive from.</param>
    /// <param name="startModeText">The stored <c>PlotStartMode</c> text.</param>
    /// <param name="finishModeText">The stored <c>PlotFinishMode</c> text.</param>
    /// <param name="startDateText">The stored <c>PlotStartDate</c> text.</param>
    /// <param name="finishDateText">The stored <c>PlotFinishDate</c> text.</param>
    /// <param name="paddingDays">The automatic-mode padding, in days per side.</param>
    /// <returns>The raw extent, or the typed refusal.</returns>
    public static PlotRangeOutcome TryResolve(
        IReadOnlyList<GanttEvent>? events,
        string? startModeText,
        string? finishModeText,
        string? startDateText,
        string? finishDateText,
        int paddingDays)
    {
        if (events is null)
        {
            return new PlotRangeOutcome(null, null, PlotRangeRefusal.NullEvents);
        }

        if (!PlotRangeModes.TryParse(startModeText, out PlotRangeMode startMode)
            || !PlotRangeModes.TryParse(finishModeText, out PlotRangeMode finishMode))
        {
            return new PlotRangeOutcome(null, null, PlotRangeRefusal.UnknownMode);
        }

        DateOnly? start = null;
        if (startMode == PlotRangeMode.Explicit)
        {
            start = ParseExplicitDate(startDateText, out PlotRangeRefusal? startRefusal);
            if (start is null)
            {
                return new PlotRangeOutcome(null, null, startRefusal!.Value);
            }
        }

        DateOnly? finish = null;
        if (finishMode == PlotRangeMode.Explicit)
        {
            finish = ParseExplicitDate(finishDateText, out PlotRangeRefusal? finishRefusal);
            if (finish is null)
            {
                return new PlotRangeOutcome(null, null, finishRefusal!.Value);
            }
        }

        if (startMode == PlotRangeMode.DataRange || finishMode == PlotRangeMode.DataRange)
        {
            if (!TryDeriveDataExtent(events, paddingDays, out DateOnly dataStart, out DateOnly dataFinish))
            {
                return new PlotRangeOutcome(null, null, PlotRangeRefusal.NoPlotRange);
            }

            start ??= dataStart;
            finish ??= dataFinish;
        }

        return start!.Value > finish!.Value
            ? new PlotRangeOutcome(null, null, PlotRangeRefusal.StartAfterFinish)
            : new PlotRangeOutcome(start, finish, null);
    }

    /// <summary>
    /// Parses one user-entered explicit date, accepting every reasonable date
    /// format: day-first formats first (dd/MM/yyyy, d/M/yyyy, etc.), then
    /// month-first formats (MM/dd/yyyy, M/d/yyyy, etc.), and finally ISO
    /// (yyyy-MM-dd). Blank text fails rather than defaulting. No
    /// CurrentCulture fallback is used so unmatched dates fail deterministically.
    /// Office-free so both the ribbon setters and the resolve path share one
    /// authority, and the stored form is always normalised to <c>dd/MM/yyyy</c>
    /// by the caller.
    /// </summary>
    /// <param name="text">The user-entered date text.</param>
    /// <param name="parsed">The parsed date when successful.</param>
    /// <returns><see langword="true" /> when the text is an actual date.</returns>
    public static bool TryParsePlotDate(string? text, out DateOnly parsed)
    {
        parsed = default;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var trimmed = text.Trim();

        // Day-first formats (tried first)
        string[] dayFirstFormats =
        [
            "dd/MM/yyyy",
            "d/M/yyyy",
            "d/M/yy",
            "dd-MMM-yyyy",
            "dd/MMM/yyyy",
            "d MMM yyyy",
            "dd MMM yyyy",
        ];

        // Month-first formats (tried only if no day-first format matches)
        string[] monthFirstFormats =
        [
            "MM/dd/yyyy",
            "M/d/yyyy",
            "M/d/yy",
        ];

        // ISO format (tried last)
        string[] isoFormats =
        [
            "yyyy-MM-dd",
        ];

        // Try day-first formats first
        foreach (var format in dayFirstFormats)
        {
            if (DateOnly.TryParseExact(trimmed, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed)
                && parsed != default)
            {
                return true;
            }
        }

        // Only try month-first formats if no day-first format matched
        foreach (var format in monthFirstFormats)
        {
            if (DateOnly.TryParseExact(trimmed, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed)
                && parsed != default)
            {
                return true;
            }
        }

        // Try ISO format
        foreach (var format in isoFormats)
        {
            if (DateOnly.TryParseExact(trimmed, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed)
                && parsed != default)
            {
                return true;
            }
        }

        // No CurrentCulture fallback - unmatched dates fail deterministically
        return false;
    }

    /// <summary>
    /// Parses one explicit date: blank is missing, unparsable is unparsable,
    /// and the default value is the <c>TimeScale</c> default-dates refusal.
    /// </summary>
    private static DateOnly? ParseExplicitDate(string? text, out PlotRangeRefusal? refusal)
    {
        refusal = null;
        if (string.IsNullOrWhiteSpace(text))
        {
            refusal = PlotRangeRefusal.MissingExplicitDate;
            return null;
        }

        if (!TryParsePlotDate(text, out DateOnly parsed))
        {
            refusal = PlotRangeRefusal.UnparsableExplicitDate;
            return null;
        }

        if (parsed == default)
        {
            refusal = PlotRangeRefusal.DefaultDates;
            return null;
        }

        return parsed;
    }

    /// <summary>
    /// Derives the padded data extent: min <c>Start</c> over every event that
    /// carries one, max over the starts then over <c>Finish</c> of span-dated
    /// types only.
    /// </summary>
    /// <remarks>
    /// Every entity contributes its <c>Start</c>, and ONLY a span-dated one
    /// contributes its <c>Finish</c> — a milestone, delineator, Splitter or
    /// Spacer reads <c>Start</c> as its single date per the entity guide's
    /// date-mode classification. The later of the two maxima wins, so a
    /// single-date event always lands inside the range.
    /// </remarks>
    private static bool TryDeriveDataExtent(
        IReadOnlyList<GanttEvent> events,
        int paddingDays,
        out DateOnly dataStart,
        out DateOnly dataFinish)
    {
        dataStart = default;
        dataFinish = default;

        var starts = new List<DateOnly>();
        var finishes = new List<DateOnly>();
        foreach (GanttEvent @event in events)
        {
            if (@event.Start is { } start)
            {
                starts.Add(start);
            }

            if (EntityTypeCatalog.GetDefinition(@event.Type)?.DateMode == EntityDateMode.StartFinish
                && @event.Finish is { } finish)
            {
                finishes.Add(finish);
            }
        }

        if (starts.Count == 0)
        {
            return false;
        }

        var padding = Math.Max(0, paddingDays);
        DateOnly earliest = starts.Min();
        DateOnly latest = starts.Max();
        if (finishes.Count > 0 && finishes.Max() > latest)
        {
            latest = finishes.Max();
        }

        dataStart = earliest.AddDays(-padding);
        dataFinish = latest.AddDays(padding);
        return true;
    }

    /// <summary>
    /// Validates plot-range settings without resolving actual dates.
    /// Used by the Ribbon state service to commit settings before
    /// they reach the scene-request factory.
    /// </summary>
    /// <param name="settings">The settings to validate.</param>
    /// <exception cref="ArgumentException">
    /// Thrown when the mode is unknown, or when an explicit mode
    /// is missing one or both dates, or when the dates are unparsable
    /// in any reasonable format (stored as dd/MM/yyyy) or the start
    /// is after the finish.
    /// </exception>
    public static void ValidateSettings(PlotRangeSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);

        if (!PlotRangeModes.TryParse(settings.Mode, out _))
        {
            throw new ArgumentException(
                $"Invalid plot-range mode: {settings.Mode}",
                nameof(settings));
        }

        if (settings.Mode == nameof(PlotRangeMode.Explicit))
        {
            if (string.IsNullOrWhiteSpace(settings.StartDate))
            {
                throw new ArgumentException("Explicit mode requires a start date.", nameof(settings));
            }

            if (string.IsNullOrWhiteSpace(settings.FinishDate))
            {
                throw new ArgumentException("Explicit mode requires a finish date.", nameof(settings));
            }

            if (!TryParsePlotDate(settings.StartDate, out DateOnly startParsed))
            {
                throw new ArgumentException($"Unparsable start date: {settings.StartDate}", nameof(settings));
            }

            if (!TryParsePlotDate(settings.FinishDate, out DateOnly finishParsed))
            {
                throw new ArgumentException($"Unparsable finish date: {settings.FinishDate}", nameof(settings));
            }

            if (startParsed > finishParsed)
            {
                throw new ArgumentException(
                    $"Start date {settings.StartDate} is after finish date {settings.FinishDate}.",
                    nameof(settings));
            }
        }
    }
}
