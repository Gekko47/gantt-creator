namespace GanttCreator.Core;

/// <summary>Factory helpers for typed neutral cells.</summary>
public static class GanttCells
{
    /// <summary>Creates an empty cell.</summary>
    /// <typeparam name="T">The cell value type.</typeparam>
    /// <returns>An empty cell.</returns>
    public static GanttCell<T> Empty<T>() => new(GanttCellState.Empty, default);

    /// <summary>Creates a supported value cell.</summary>
    /// <typeparam name="T">The cell value type.</typeparam>
    /// <param name="value">The typed value.</param>
    /// <returns>A value cell.</returns>
    public static GanttCell<T> Value<T>(T value) => new(GanttCellState.Value, value);

    /// <summary>Creates an Excel error cell.</summary>
    /// <typeparam name="T">The cell value type.</typeparam>
    /// <param name="errorCode">The neutral error code.</param>
    /// <returns>An Excel error cell.</returns>
    public static GanttCell<T> ExcelError<T>(GanttExcelErrorCode errorCode) =>
        new(GanttCellState.ExcelError, default, ErrorCode: errorCode);

    /// <summary>Creates an unsupported-payload cell.</summary>
    /// <typeparam name="T">The cell value type.</typeparam>
    /// <returns>An unsupported cell.</returns>
    public static GanttCell<T> Unsupported<T>() => new(GanttCellState.Unsupported, default);
}
