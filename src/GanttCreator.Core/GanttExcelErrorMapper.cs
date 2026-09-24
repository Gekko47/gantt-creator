namespace GanttCreator.Core;

/// <summary>Maps Office <c>Value2</c> error integers to neutral error codes.</summary>
public static class GanttExcelErrorMapper
{
    private const int _errorBase = -2146826288;
    private const int _firstKnownCode = 2000;
    private const int _lastKnownCode = 2050;

    /// <summary>Maps a raw Excel error integer without referencing Office types.</summary>
    /// <param name="value2">The raw integer payload.</param>
    /// <param name="code">The neutral error code.</param>
    /// <returns><see langword="true"/> when the payload is an Excel error integer.</returns>
    public static bool TryMap(int value2, out GanttExcelErrorCode code)
    {
        var rawCode = unchecked(value2 - _errorBase) + 2000;
        if (rawCode is >= _firstKnownCode and <= _lastKnownCode && Enum.IsDefined(typeof(GanttExcelErrorCode), rawCode))
        {
            code = (GanttExcelErrorCode)rawCode;
            return true;
        }

        code = GanttExcelErrorCode.Unknown;
        return true;
    }
}
