namespace GanttCreator.Core.Scene;

/// <summary>
/// The margin between the chart's content bounds and the chart frame, one value
/// per side (ADR-0031 D1).
/// </summary>
/// <param name="LeftPt">The margin to the content's left edge, in points.</param>
/// <param name="TopPt">The margin above the content's top edge, in points.</param>
/// <param name="RightPt">The margin to the content's right edge, in points.</param>
/// <param name="BottomPt">The margin below the content's bottom edge, in points.</param>
/// <remarks>
/// <para>
/// <b>Why this is per-side rather than one scalar.</b> It was a single
/// <c>ChartOuterPaddingPt</c> applied to all four sides, which cannot express a live
/// chart whose top and bottom padding are whole worksheet rows while its left
/// padding is nothing. A live chart sits beside a data table, so a left margin
/// pushes the plot away from the table it belongs to; the top and bottom margins
/// want to be row-aligned so the chart's edges meet the sheet's grid. One scalar
/// forces all three to be the same number.
/// </para>
/// <para>
/// <b>Why the record does not validate.</b> A negative or non-finite side is
/// refused where the frame is built, not here, so there is exactly one place that
/// decides what a bad margin means and one refusal to test. Constructing the record
/// is pure data.
/// </para>
/// <para>
/// <b>Export keeps a uniform margin.</b> A PNG or a slide has no table beside it
/// and no row grid to align to, so the symmetric margin is still correct there.
/// <see cref="Uniform"/> is how that case is expressed without a second code path.
/// </para>
/// </remarks>
public readonly record struct ChartPaddingPt(double LeftPt, double TopPt, double RightPt, double BottomPt)
{
    /// <summary>The no-padding value, used where a frame hugs its content.</summary>
    public static ChartPaddingPt None => new(0, 0, 0, 0);

    /// <summary>Creates the same margin on all four sides.</summary>
    /// <param name="points">The margin in points.</param>
    /// <returns>The uniform padding.</returns>
    public static ChartPaddingPt Uniform(double points) => new(points, points, points, points);

    /// <summary>Gets the largest single margin on any side.</summary>
    /// <remarks>
    /// The uniform case a caller must not mistake for this one: a frame built with
    /// a uniform margin needs a uniform amount, and a frame built from rows needs
    /// the rows' own heights. This is the maximum of the four, which is the
    /// conservative value when a caller has only one number.
    /// </remarks>
    public double MaxPt => Math.Max(Math.Max(LeftPt, TopPt), Math.Max(RightPt, BottomPt));
}
