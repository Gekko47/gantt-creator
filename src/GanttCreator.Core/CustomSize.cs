namespace GanttCreator.Core;

/// <summary>Why a custom centimetre width was refused.</summary>
public enum CustomSizeRefusal
{
    /// <summary>The width is not a finite positive number.</summary>
    InvalidWidth = 0,

    /// <summary>The width is outside the approved centimetre band.</summary>
    WidthOutOfBand = 1,
}

/// <summary>
/// The user-entered custom **width** in centimetres (R8.9, ADR-0072).
/// </summary>
/// <remarks>
/// <para>
/// The custom width is consumed as the whole composition width, exactly like a preset
/// width: the data table keeps its measured column widths and the plot takes whatever
/// they and the frame chrome leave. The value is kept in centimetres here because that
/// is what the user typed; points are derived once so the ribbon display and the
/// geometry cannot round differently.
/// </para>
/// <para>
/// <b>There is no custom height.</b> The live height is the measured rows (ADR-0030
/// D1/D2), and a typed height could only be honoured by resizing the user's rows,
/// scaling the composition (which changes the type), or deriving a width from a fixed
/// ratio that cannot hold while the width is also the live budget. The height is a
/// measurement, mirrored into the <c>SizeHeightCm</c> row by the AddIn.
/// </para>
/// <para>
/// The band is the approved 1.0-200.0 cm. A width in the band can still be refused
/// later by the plot-geometry shortfall when the measured table leaves too little room;
/// that minimum depends on measurements the ribbon owns, so it is clamped there rather
/// than here.
/// </para>
/// </remarks>
public sealed record CustomSizeCm
{
    /// <summary>The default custom width, in centimetres (A4 landscape).</summary>
    public const double DefaultWidthCm = 29.7;

    /// <summary>The smallest permitted custom width, in centimetres.</summary>
    public const double MinimumCm = 1.0;

    /// <summary>The largest permitted custom width, in centimetres.</summary>
    public const double MaximumCm = 200.0;

    private CustomSizeCm(double widthCm)
    {
        WidthCm = widthCm;
    }

    /// <summary>The width the user typed, in centimetres.</summary>
    public double WidthCm { get; }

    /// <summary>The width in points, derived from the centimetres once.</summary>
    public double WidthPt => WidthCm * 10.0 * SizePresets.PointsPerMillimetre;

    /// <summary>Attempts to create a custom width inside the approved band.</summary>
    /// <param name="widthCm">The width the user typed, in centimetres.</param>
    /// <param name="size">The validated width on success; otherwise null.</param>
    /// <param name="refusal">The typed refusal on failure; otherwise null.</param>
    /// <returns><see langword="true"/> when the width is valid.</returns>
    public static bool TryCreate(
        double widthCm,
        out CustomSizeCm? size,
        out CustomSizeRefusal? refusal)
    {
        size = null;
        refusal = null;

        if (!double.IsFinite(widthCm) || widthCm <= 0)
        {
            refusal = CustomSizeRefusal.InvalidWidth;
            return false;
        }

        if (widthCm is < MinimumCm or > MaximumCm)
        {
            refusal = CustomSizeRefusal.WidthOutOfBand;
            return false;
        }

        size = new CustomSizeCm(widthCm);
        return true;
    }

    /// <summary>
    /// Clamps a width into the approved band, keeping a non-finite value at the default.
    /// </summary>
    /// <param name="widthCm">The width to clamp.</param>
    /// <returns>A width inside the band.</returns>
    /// <remarks>
    /// The ribbon and the request factory both read a stored centimetre value, and a
    /// hand-edited workbook can hold anything. Clamping rather than refusing at the read
    /// is what stops a bad stored value collapsing the mode: the user keeps a working
    /// Custom width and can correct it in the box.
    /// </remarks>
    public static double ClampWidth(double widthCm) =>
        double.IsFinite(widthCm)
            ? Math.Clamp(widthCm, MinimumCm, MaximumCm)
            : DefaultWidthCm;

    /// <summary>Builds a width from a stored value, clamping it into the band.</summary>
    /// <param name="widthCm">The stored width, in centimetres.</param>
    /// <returns>A width inside the approved band.</returns>
    /// <remarks>
    /// This is the read path: a stored value that is merely out of band is corrected rather
    /// than refused, because a refusal here would collapse the mode and leave the user with
    /// a ribbon that disagrees with the workbook. A non-finite or unparsable value never
    /// reaches this method — the caller refuses it as unreadable first.
    /// </remarks>
    public static CustomSizeCm FromClampedWidth(double widthCm) => new(ClampWidth(widthCm));
}
