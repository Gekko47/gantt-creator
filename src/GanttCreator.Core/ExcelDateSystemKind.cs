namespace GanttCreator.Core;

/// <summary>
/// The Excel workbook date system used to interpret numeric date serials.
/// </summary>
public enum ExcelDateSystemKind
{
    /// <summary>The Windows 1900 date system.</summary>
    Windows1900 = 0,

    /// <summary>The Macintosh 1904 date system.</summary>
    Macintosh1904 = 1,
}
