namespace GanttCreator.Core;

/// <summary>
/// Office-free Excel date-system conversion for supported workbook epochs.
/// </summary>
public sealed class ExcelDateSystemConverter : IExcelDateSystemConverter
{
    /// <summary>The shared stateless converter.</summary>
    public static ExcelDateSystemConverter Instance { get; } = new();

    /// <inheritdoc />
    public bool IsSupported(ExcelDateSystemKind dateSystem) => dateSystem == ExcelDateSystemKind.Windows1900;

    /// <inheritdoc />
    public bool TryConvertDate(object? value, ExcelDateSystemKind dateSystem, out DateOnly? convertedDate)
    {
        convertedDate = null;
        if (!IsSupported(dateSystem))
        {
            return false;
        }

        switch (value)
        {
            case double serial:
                if (!double.IsFinite(serial) || serial < ExcelCellConverter.MinSerial || serial > ExcelCellConverter.MaxSerial)
                {
                    return false;
                }

                convertedDate = DateOnly.FromDateTime(DateTime.FromOADate(serial));
                return true;
            case DateTime dateTime:
                convertedDate = DateOnly.FromDateTime(dateTime);
                return true;
            default:
                return false;
        }
    }
}
