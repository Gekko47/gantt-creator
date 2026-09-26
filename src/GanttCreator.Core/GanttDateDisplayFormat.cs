using System.Globalization;

namespace GanttCreator.Core;

/// <summary>
/// The closed event-date display formats stored in workbook settings.
/// </summary>
/// <remarks>
/// <para>
/// ADR-0016 D3 fixes this set as closed. The first release has exactly one
/// approved member, <see cref="DdMMyyyy"/>, selected by the
/// <c>DateDisplayFormat</c> setting; the stored text is the member name, so
/// the value survives a workbook round-trip unchanged.
/// </para>
/// <para>
/// ADR-0016 D5: adding a later approved member changes the permitted value
/// set, not the settings key set, so it does not advance the workbook schema
/// version.
/// </para>
/// </remarks>
public enum GanttDateDisplayFormat
{
    /// <summary>Two-digit day, two-digit month, four-digit year, such as <c>05/01/2026</c>.</summary>
    DdMMyyyy = 0,
}

/// <summary>The culture-invariant date formatting contract (ADR-0016 D4).</summary>
/// <remarks>
/// <para>
/// Every rendered event date passes through here, so no scene builder and no
/// renderer can re-derive the pattern or apply the host's regional settings.
/// An <c>en-GB</c> host and an <c>en-US</c> host must produce byte-identical
/// scene text for the same date, which is what makes the R3.12 golden scene
/// snapshot stable.
/// </para>
/// </remarks>
public static class GanttDateFormatting
{
    /// <summary>The approved invariant pattern for <see cref="GanttDateDisplayFormat.DdMMyyyy"/>.</summary>
    public const string DdMMyyyyPattern = "dd/MM/yyyy";

    /// <summary>Formats a date using the invariant pattern for a display format.</summary>
    /// <param name="date">The date to format.</param>
    /// <param name="format">The approved display format.</param>
    /// <returns>The invariant formatted date text.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when <paramref name="format"/> is not a defined
    /// <see cref="GanttDateDisplayFormat"/> value.
    /// </exception>
    public static string Format(DateOnly date, GanttDateDisplayFormat format) =>
        format switch
        {
            GanttDateDisplayFormat.DdMMyyyy => FormatDdMMyyyy(date),
            _ => throw new ArgumentOutOfRangeException(
                nameof(format),
                format,
                "The date display format is not a defined value."),
        };

    /// <summary>Formats a date as the approved invariant <c>dd/mm/yyyy</c> text.</summary>
    /// <param name="date">The date to format.</param>
    /// <returns>The invariant formatted date text.</returns>
    public static string FormatDdMMyyyy(DateOnly date) =>
        date.ToString(DdMMyyyyPattern, CultureInfo.InvariantCulture);
}
