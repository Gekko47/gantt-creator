namespace GanttCreator.Core;

/// <summary>
/// The single authority for the chart's month-snapped plot bounds (R5.1).
/// </summary>
/// <remarks>
/// <para>
/// Both the Refresh render pipeline (<c>ExcelSceneBuildRequestFactory</c>)
/// and the Ribbon's AUTO display derivation (<c>RibbonStateService</c>)
/// must show the same bounds: the padded data extent snapped outward to
/// whole months, with a degenerate (single-day) chart widened to two days.
/// Two copies of this arithmetic previously lived in those files behind a
/// "mirrors" comment; this type is the one implementation both call.
/// </para>
/// </remarks>
public static class PlotMonthBounds
{
    /// <summary>
    /// Returns the first day of the month containing <paramref name="date"/>:
    /// the chart's start bound.
    /// </summary>
    /// <param name="date">The resolved padded data date.</param>
    /// <returns>The first day of that month.</returns>
    /// <remarks>
    /// Constructed from the date's own year and month parts: no month
    /// arithmetic, so no December or <see cref="DateOnly.MinValue"/>
    /// boundary to reason about.
    /// </remarks>
    public static DateOnly SnapToMonthStart(DateOnly date) => new(date.Year, date.Month, 1);

    /// <summary>
    /// Returns the last day of the month containing <paramref name="date"/>:
    /// the chart's finish bound.
    /// </summary>
    /// <param name="date">The resolved padded data date.</param>
    /// <returns>The last day of that month.</returns>
    /// <remarks>
    /// Day zero of the following month is the last day of this one, which
    /// avoids both a hard-coded 28/30/31 table and the December overflow
    /// that <c>AddMonths(1)</c> would need a range check for.
    /// </remarks>
    public static DateOnly SnapToMonthEnd(DateOnly date) =>
        SnapToMonthStart(date).AddMonths(1).AddDays(-1);

    /// <summary>
    /// Snaps a resolved inclusive extent outward to whole months, widening a
    /// degenerate chart (finish on or before start) to two days.
    /// </summary>
    /// <param name="start">The resolved inclusive start.</param>
    /// <param name="finish">The resolved inclusive finish.</param>
    /// <returns>The month-snapped chart bounds Refresh renders.</returns>
    public static (DateOnly Start, DateOnly Finish) SnapExtent(DateOnly start, DateOnly finish)
    {
        DateOnly boundStart = SnapToMonthStart(start);
        DateOnly boundFinish = SnapToMonthEnd(finish);
        if (boundFinish.DayNumber <= boundStart.DayNumber)
        {
            boundFinish = boundStart.AddDays(1);
        }

        return (boundStart, boundFinish);
    }
}
