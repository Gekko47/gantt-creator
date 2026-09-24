namespace GanttCreator.Core;

/// <summary>
/// Converts Excel cell date payloads for a supported workbook date system.
/// </summary>
public interface IExcelDateSystemConverter
{
    /// <summary>Determines whether a date system is supported.</summary>
    /// <param name="dateSystem">The workbook date system.</param>
    /// <returns><see langword="true"/> when conversion is supported.</returns>
    bool IsSupported(ExcelDateSystemKind dateSystem);

    /// <summary>Converts an Excel <c>Value2</c> date payload.</summary>
    /// <param name="value">The raw cell payload.</param>
    /// <param name="dateSystem">The workbook date system.</param>
    /// <param name="convertedDate">The converted date, or <see langword="null"/>.</param>
    /// <returns><see langword="true"/> when a date was produced.</returns>
    bool TryConvertDate(object? value, ExcelDateSystemKind dateSystem, out DateOnly? convertedDate);
}
