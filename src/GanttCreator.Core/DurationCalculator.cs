namespace GanttCreator.Core;

/// <summary>Why a duration could not be calculated for a row.</summary>
public enum DurationRefusal
{
    /// <summary>The row's <c>Start</c> cell is blank, so there is no span to measure.</summary>
    StartMissing = 0,

    /// <summary>A span row's <c>Finish</c> cell is blank.</summary>
    FinishMissing = 1,

    /// <summary>The row's <c>Finish</c> is before its <c>Start</c>.</summary>
    FinishBeforeStart = 2,
}

/// <summary>One row's calculated <c>Duration</c> and the text to write.</summary>
/// <param name="RowNumber">The worksheet row number.</param>
/// <param name="Text">
/// The display text: the inclusive day count for a span, the approved marker for a
/// point entity, or blank to clear a stale value.
/// </param>
public sealed record DurationWrite(int RowNumber, string Text);

/// <summary>One row that produced no duration, and why.</summary>
/// <param name="RowNumber">The worksheet row number.</param>
/// <param name="Reason">The typed reason the duration is unavailable.</param>
public sealed record DurationRefusalEntry(int RowNumber, DurationRefusal Reason);

/// <summary>
/// The typed outcome of planning the <c>Duration</c> write for a set of rows.
/// </summary>
/// <param name="Writes">
/// One entry per row whose <c>Duration</c> changes. A row already carrying the
/// calculated value is absent, so a correctly written sheet produces an empty write
/// set and the adapter performs no work at all.
/// </param>
/// <param name="Refusals">
/// The rows that produced no value, with the reason. This is reported rather than
/// swallowed: D4 forbids inventing a <c>0</c>, a <c>-1</c>, or a carry-over from the
/// previous Refresh, so a row that cannot be measured must stay visible.
/// </param>
public sealed record DurationWritePlan(
    IReadOnlyList<DurationWrite> Writes,
    IReadOnlyList<DurationRefusalEntry> Refusals)
{
    /// <summary>How many cells the plan would write.</summary>
    public int WriteCount => Writes.Count;
}

/// <summary>
/// Calculates the <c>Duration</c> column from date-only semantics (R4.7F).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this is pure and lives in Core.</b> It is a function over two dates and a
/// row type, so it is fully testable without a workbook and the Office layer only
/// performs the write. That is also why ADR-0006's rejection of the 1904 date
/// system is irrelevant here: the calculation is on <see cref="DateOnly"/>, which
/// has no serial-number arithmetic to get wrong.
/// </para>
/// <para>
/// <b>Why the span is inclusive.</b> The entity guide's existing rule is
/// <c>DurationDays = Finish.DayNumber - Start.DayNumber + 1</c>, matching
/// <c>TimeScale.TryDurationDays</c> which the scene already uses for bar width. This
/// row implements that rule rather than introducing a second one, and <c>Finish</c>
/// is never incremented.
/// </para>
/// <para>
/// <b>Why point and structural rows are not spans.</b> A milestone and a delineator
/// read <c>Start</c> as their single date and never use <c>Finish</c> for geometry.
/// Computing <c>0</c> for them would imply a zero-length span, which is a different
/// claim. A <c>Splitter</c> and a <c>Spacer</c> read no dates at all, so their cell is
/// cleared rather than filled.
/// </para>
/// <para>
/// <b>Why a bad date produces no value rather than a sentinel.</b> A <c>0</c> or a
/// <c>-1</c> would both read as a real, wrong duration in the user's table. D4
/// requires the cell to be left empty and the row reported instead, so a
/// miscalculated duration cannot be mistaken for a zero-length one.
/// </para>
/// </remarks>
public static class DurationCalculator
{
    /// <summary>
    /// The marker shown for a point entity, which has a date but no span.
    /// </summary>
    public const string NonDurationMarker = "-";

    /// <summary>
    /// The inclusive calendar-day duration of a span, or a refusal.
    /// </summary>
    /// <param name="start">The inclusive start date.</param>
    /// <param name="finish">The inclusive finish date.</param>
    /// <param name="days">The inclusive day count on success.</param>
    /// <returns><see langword="true"/> when the finish is not before the start.</returns>
    public static bool TryInclusiveDays(DateOnly start, DateOnly finish, out int days)
    {
        // `DateOnly` is already date-only, so D1's normalisation is structural: a
        // caller holding a DateTime must convert, and there is no time component
        // here for this function to strip.
        days = finish.DayNumber - start.DayNumber + 1;
        return finish >= start;
    }

    /// <summary>
    /// Calculates the <c>Duration</c> text for one row, or a refusal.
    /// </summary>
    /// <param name="rowNumber">The worksheet row number, for the refusal entry.</param>
    /// <param name="type">The row's resolved entity type.</param>
    /// <param name="start">The row's start date.</param>
    /// <param name="finish">The row's finish date.</param>
    /// <param name="text">The display text, or <see langword="null"/> when refused.</param>
    /// <param name="refusal">The typed refusal, or <see langword="null"/> on success.</param>
    /// <returns><see langword="true"/> when a value was produced.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when <paramref name="type"/> is not in the catalogue. An unknown type is
    /// refused by the validator before this point, and guessing a date mode here would
    /// invent a duration for a row whose type was never understood.
    /// </exception>
    public static bool TryCalculate(
        int rowNumber,
        GanttEntityType type,
        DateOnly? start,
        DateOnly? finish,
        out string? text,
        out DurationRefusalEntry? refusal)
    {
        // The date MODE decides what is read, not the kind: a delineator is a
        // structural drawing primitive but still reads Start as its single date.
        EntityDateMode mode = EntityTypeCatalog.GetDefinition(type)?.DateMode
            ?? throw new ArgumentOutOfRangeException(
                nameof(type),
                type,
                "The entity type is not in the catalogue, so its date mode is unknown.");

        switch (mode)
        {
            case EntityDateMode.None:
                // A Splitter or Spacer has no dates at all. Blank, not a marker: the
                // row occupies space but states nothing about time.
                text = string.Empty;
                refusal = null;
                return true;

            case EntityDateMode.StartOnly:
                if (start is null)
                {
                    text = null;
                    refusal = new DurationRefusalEntry(rowNumber, DurationRefusal.StartMissing);
                    return false;
                }

                // A point event has no span, so it shows the marker rather than a
                // count. Its Finish is irrelevant and deliberately unread.
                text = NonDurationMarker;
                refusal = null;
                return true;

            case EntityDateMode.StartFinish:
                if (start is null)
                {
                    text = null;
                    refusal = new DurationRefusalEntry(rowNumber, DurationRefusal.StartMissing);
                    return false;
                }

                if (finish is null)
                {
                    text = null;
                    refusal = new DurationRefusalEntry(rowNumber, DurationRefusal.FinishMissing);
                    return false;
                }

                if (!TryInclusiveDays(start.Value, finish.Value, out var days))
                {
                    text = null;
                    refusal = new DurationRefusalEntry(rowNumber, DurationRefusal.FinishBeforeStart);
                    return false;
                }

                text = days.ToString(System.Globalization.CultureInfo.InvariantCulture);
                refusal = null;
                return true;

            default:
                throw new ArgumentOutOfRangeException(
                    nameof(type),
                    type,
                    "An undefined EntityDateMode must be refused, not defaulted.");
        }
    }

    /// <summary>
    /// Plans the <c>Duration</c> write for a batch of validated events.
    /// </summary>
    /// <param name="events">The validated events, in any order.</param>
    /// <param name="currentDurationByRow">
    /// The <c>Duration</c> text each row currently shows, keyed by row number, so an
    /// unchanged row is excluded from the write set. A row absent from the map is
    /// treated as currently blank.
    /// </param>
    /// <returns>
    /// The write plan and the refusals, both ordered by row number so the plan is
    /// deterministic.
    /// </returns>
    /// <exception cref="ArgumentNullException">
    /// Thrown when <paramref name="events"/> or <paramref name="currentDurationByRow"/>
    /// is null. The current-value map is required rather than optional so a caller
    /// cannot omit it and silently rewrite every row.
    /// </exception>
    public static DurationWritePlan Plan(
        IReadOnlyList<GanttEvent> events,
        IReadOnlyDictionary<int, string?> currentDurationByRow)
    {
        ArgumentNullException.ThrowIfNull(events);
        ArgumentNullException.ThrowIfNull(currentDurationByRow);

        List<DurationWrite> writes = [];
        List<DurationRefusalEntry> refusals = [];

        foreach (GanttEvent @event in events.OrderBy(static e => e.RowNumber))
        {
            if (!TryCalculate(
                    @event.RowNumber,
                    @event.Type,
                    @event.Start,
                    @event.Finish,
                    out var text,
                    out DurationRefusalEntry? refusal))
            {
                refusals.Add(refusal!);

                // A refused row is still cleared when it carries a stale value from a
                // previous Refresh. Leaving yesterday's number on a row whose dates are
                // now invalid is precisely the misleading value D4 forbids, so the
                // plan blanks the cell instead of abandoning it.
                if (currentDurationByRow.TryGetValue(@event.RowNumber, out var stale)
                    && !string.IsNullOrEmpty(stale))
                {
                    writes.Add(new DurationWrite(@event.RowNumber, string.Empty));
                }

                continue;
            }

            _ = currentDurationByRow.TryGetValue(@event.RowNumber, out var current);
            if (!string.Equals(current, text, StringComparison.Ordinal))
            {
                writes.Add(new DurationWrite(@event.RowNumber, text!));
            }
        }

        return new DurationWritePlan(writes, refusals);
    }
}

