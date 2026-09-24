namespace GanttCreator.Core;

/// <summary>The raw state carried by one neutral worksheet cell.</summary>
public enum GanttCellState
{
    /// <summary>The cell is blank.</summary>
    Empty = 0,

    /// <summary>The cell contains a supported typed value.</summary>
    Value = 1,

    /// <summary>The cell contains an Excel error value.</summary>
    ExcelError = 2,

    /// <summary>The cell contains a payload unsupported by the column type.</summary>
    Unsupported = 3,
}
