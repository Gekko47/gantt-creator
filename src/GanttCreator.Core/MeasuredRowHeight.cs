namespace GanttCreator.Core;

/// <summary>What kind of row a measured height belongs to.</summary>
public enum MeasuredRowKind
{
    /// <summary>An ordinary managed row: an activity, milestone, or child.</summary>
    Managed = 0,

    /// <summary>A <c>Splitter</c> section-header row.</summary>
    Splitter = 1,

    /// <summary>A <c>Spacer</c> blank-separator row.</summary>
    Spacer = 2,
}

/// <summary>
/// One worksheet row's measured height, as read from the live sheet.
/// </summary>
/// <param name="RowNumber">The one-based worksheet row number.</param>
/// <param name="CurrentHeightPt">The row's measured height in points.</param>
/// <param name="Kind">Which height policy the row follows.</param>
public sealed record MeasuredRowHeight(
    int RowNumber,
    double CurrentHeightPt,
    MeasuredRowKind Kind = MeasuredRowKind.Managed);
