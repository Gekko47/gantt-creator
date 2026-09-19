using System.Globalization;

namespace GanttCreator.Core;

/// <summary>
/// Pure, Office-free converters from Excel <c>Range.Value2</c> cell payloads
/// to the neutral values carried by <see cref="GanttRowDto"/>. Every method
/// reads the cell value — never locale display text — and never throws for
/// routine bad input: unreadable or out-of-range values yield
/// <see langword="null"/>. Immutable (static only).
/// </summary>
public static class ExcelCellConverter
{
    /// <summary>
    /// The smallest valid OLE Automation date serial (1899-12-31).
    /// </summary>
    public const double MinSerial = 1.0;

    /// <summary>
    /// The largest valid OLE Automation date serial (9999-12-31).
    /// </summary>
    public const double MaxSerial = 2958465.0;

    /// <summary>
    /// Normalizes workbook text: trims surrounding whitespace and maps empty
    /// or whitespace-only input to <see langword="null"/>. Strings stay
    /// Ordinal with no case-folding (mirrors <see cref="EntityTypeCatalog"/>
    /// and <see cref="GanttRowId"/>). Excel error cells surface as
    /// <see cref="int"/> in <c>Value2</c> and map to <see langword="null"/>
    /// with the row preserved. Never throws.
    /// </summary>
    /// <param name="value">The raw <c>Value2</c> cell payload.</param>
    /// <returns>The trimmed text, or <see langword="null"/> when blank, null, or unreadable.</returns>
    public static string? ToText(object? value)
    {
        return value switch
        {
            null => null,
            int => null,
            string text => TrimOrNull(text),
            bool boolean => boolean ? "TRUE" : "FALSE",
            double number when double.IsFinite(number) => number.ToString(CultureInfo.InvariantCulture),
            DateTime dateTime => dateTime.ToString("o", CultureInfo.InvariantCulture),
            _ => null,
        };
    }

    /// <summary>
    /// Converts a <c>Value2</c> date payload to <see cref="DateOnly"/>.
    /// Accepts <see cref="double"/> serials (1900 date system) and
    /// <see cref="DateTime"/>. Strings are never date-parsed (anti-locale
    /// rule); <see cref="bool"/> and Excel error <see cref="int"/> map to
    /// <see langword="null"/>. Serial 60 (the inherited 1900 leap-bug phantom
    /// 1900-02-29) converts via <see cref="DateTime.FromOADate"/> semantics
    /// and is pinned by test; it is not "fixed" here. Never throws.
    /// </summary>
    /// <param name="value">The raw <c>Value2</c> cell payload.</param>
    /// <param name="date">The converted date, or <see langword="null"/> when blank or unreadable.</param>
    /// <returns><see langword="true"/> when a date was produced.</returns>
    public static bool TryConvertDate(object? value, out DateOnly? date)
    {
        date = null;
        switch (value)
        {
            case null:
                return false;
            case double serial:
                if (!double.IsFinite(serial) || serial < MinSerial || serial > MaxSerial)
                {
                    return false;
                }

                date = DateOnly.FromDateTime(DateTime.FromOADate(serial));
                return true;
            case DateTime dateTime:
                date = DateOnly.FromDateTime(dateTime);
                return true;
            default:
                return false;
        }
    }

    /// <summary>
    /// Converts a <c>Value2</c> stack-index payload to <see cref="int"/>.
    /// Accepts whole <see cref="double"/>s and invariant <see cref="string"/>
    /// integers. Fractional, non-finite, negative, and unreadable values map
    /// to <see langword="null"/>. Never throws.
    /// </summary>
    /// <param name="value">The raw <c>Value2</c> cell payload.</param>
    /// <param name="index">The converted index, or <see langword="null"/> when blank or unreadable.</param>
    /// <returns><see langword="true"/> when an index was produced.</returns>
    public static bool TryConvertStackIndex(object? value, out int? index)
    {
        index = null;
        switch (value)
        {
            case null:
                return false;
            case double number:
                if (!double.IsFinite(number) || number < 0 || number > int.MaxValue || number != Math.Truncate(number))
                {
                    return false;
                }

                index = (int)number;
                return true;
            case string text:
                if (TrimOrNull(text) is not { } trimmed
                    || !int.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)
                    || parsed < 0)
                {
                    return false;
                }

                index = parsed;
                return true;
            default:
                return false;
        }
    }

    /// <summary>
    /// Converts a <c>Value2</c> visibility payload to <see cref="bool"/>.
    /// Accepts <see cref="bool"/>, <c>1/0</c> doubles, and the invariant
    /// strings <c>TRUE/FALSE/1/0</c>. Blank and unreadable values map to
    /// <see langword="null"/>. Never throws.
    /// </summary>
    /// <param name="value">The raw <c>Value2</c> cell payload.</param>
    /// <param name="visible">The converted value, or <see langword="null"/> when blank or unreadable.</param>
    /// <returns><see langword="true"/> when a value was produced.</returns>
    public static bool TryConvertVisible(object? value, out bool? visible)
    {
        visible = null;
        switch (value)
        {
            case null:
                return false;
            case bool boolean:
                visible = boolean;
                return true;
            case double number:
                if (number == 1.0)
                {
                    visible = true;
                    return true;
                }

                if (number == 0.0)
                {
                    visible = false;
                    return true;
                }

                return false;
            case string text:
                if (TrimOrNull(text) is not { } trimmed)
                {
                    return false;
                }

                if (string.Equals(trimmed, "TRUE", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(trimmed, "1", StringComparison.Ordinal))
                {
                    visible = true;
                    return true;
                }

                if (string.Equals(trimmed, "FALSE", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(trimmed, "0", StringComparison.Ordinal))
                {
                    visible = false;
                    return true;
                }

                return false;
            default:
                return false;
        }
    }

    /// <summary>
    /// Trims text and maps empty results to <see langword="null"/>.
    /// </summary>
    /// <param name="text">The text to trim.</param>
    /// <returns>The trimmed text, or <see langword="null"/> when empty.</returns>
    private static string? TrimOrNull(string text)
    {
        var trimmed = text.Trim();
        return trimmed.Length == 0 ? null : trimmed;
    }

}
