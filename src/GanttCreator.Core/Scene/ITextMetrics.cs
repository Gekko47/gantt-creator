namespace GanttCreator.Core.Scene;

/// <summary>One resolved text measurement in points.</summary>
/// <param name="WidthPt">The measured text width in points.</param>
/// <param name="HeightPt">The measured single-line text height in points.</param>
public sealed record TextMeasurement(double WidthPt, double HeightPt);

/// <summary>
/// The single injected deterministic text-metrics seam used to measure labels
/// during scene construction.
/// </summary>
/// <remarks>
/// <para>
/// <c>docs/02-ARCHITECTURE.md</c> ("Coordinate and rounding policy") requires
/// exactly one such service: the scene measures, and every renderer consumes the
/// resolved label bounds and may not re-select a side. The implementation, font,
/// and version are pinned for tests.
/// </para>
/// <para>
/// R3.6 introduced this port with both a width and a height because label
/// placement needs a label box, not just a run width. R3.5's narrower
/// <c>ITextWidthMeasurer</c> measured width only, because header labels were
/// centred inside a band whose height was already known; the frame/band builder
/// composes onto this port so Core exposes one measuring seam.
/// </para>
/// </remarks>
public interface ITextMetrics
{
    /// <summary>Attempts to measure a single line of text.</summary>
    /// <param name="text">The exact text to measure.</param>
    /// <param name="measurement">The resolved measurement when successful.</param>
    /// <returns><see langword="true"/> when a finite, non-negative measurement was produced.</returns>
    bool TryMeasure(string text, out TextMeasurement? measurement);
}

/// <summary>
/// A deterministic, table-driven text-metrics implementation used by Core tests
/// and by any host that has not yet supplied a real font.
/// </summary>
/// <remarks>
/// Widths come from a fixed per-character advance table so a measurement never
/// depends on an installed font, a culture, or a machine. Tests asserting exact
/// label bounds construct this fake explicitly and pin the table.
/// </remarks>
public sealed class FakeTextMetrics : ITextMetrics
{
    private const double _defaultAdvancePt = 4.0;
    private const double _defaultLineHeightPt = 10.0;

    private readonly Func<char, double> _advancePt;
    private readonly double _lineHeightPt;

    /// <summary>Initialises the fake with the default fixed advance and line height.</summary>
    public FakeTextMetrics()
        : this(_ => _defaultAdvancePt, _defaultLineHeightPt) { }

    /// <summary>Initialises the fake with an explicit advance table and line height.</summary>
    /// <param name="advancePt">The per-character advance lookup in points.</param>
    /// <param name="lineHeightPt">The single-line text height in points.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="advancePt"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="lineHeightPt"/> is not finite and positive.</exception>
    public FakeTextMetrics(Func<char, double> advancePt, double lineHeightPt)
    {
        ArgumentNullException.ThrowIfNull(advancePt);
        if (!double.IsFinite(lineHeightPt) || lineHeightPt <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(lineHeightPt), lineHeightPt, "Line height must be finite and positive.");
        }

        _advancePt = advancePt;
        _lineHeightPt = lineHeightPt;
    }

    /// <inheritdoc />
    public bool TryMeasure(string text, out TextMeasurement? measurement)
    {
        if (text is null)
        {
            measurement = null;
            return false;
        }

        var widthPt = 0.0;
        foreach (var character in text)
        {
            var advancePt = _advancePt(character);
            if (!double.IsFinite(advancePt) || advancePt < 0)
            {
                measurement = null;
                return false;
            }

            widthPt += advancePt;
        }

        measurement = new TextMeasurement(widthPt, _lineHeightPt);
        return true;
    }
}

/// <summary>
/// The shared text helpers Core uses whenever it must fit or truncate a label to
/// a measured width.
/// </summary>
public static class LabelText
{
    /// <summary>
    /// The truncation marker. A single character is used so the marker cannot
    /// silently widen a label past the space it was fitted into, and it matches
    /// the chart-title truncation already landed in R3.5.
    /// </summary>
    public const char Ellipsis = '…';

    /// <summary>
    /// Returns the longest prefix of <paramref name="text"/> that, once suffixed
    /// with the ellipsis, measures no wider than
    /// <paramref name="availableWidthPt"/>.
    /// </summary>
    /// <param name="text">The text to truncate.</param>
    /// <param name="availableWidthPt">The width available in points.</param>
    /// <param name="metrics">The injected deterministic text-metrics seam.</param>
    /// <returns>
    /// The truncated text, or the bare ellipsis when not even an empty prefix
    /// plus the marker fits.
    /// </returns>
    /// <remarks>
    /// <para>
    /// No font shrinking and no wrapping: entity guide §22 "Overflow" forbids
    /// both, and this is the only truncation mechanism Core uses.
    /// </para>
    /// <para>
    /// The walk is by UTF-16 code unit, so it can land between a surrogate pair.
    /// Any orphaned high surrogate is dropped before the marker is appended, so
    /// the result is always well-formed text. This mirrors the guard the
    /// validation-note truncation already carries.
    /// </para>
    /// </remarks>
    public static string Ellipsize(string text, double availableWidthPt, ITextMetrics metrics)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(metrics);
        if (!double.IsFinite(availableWidthPt))
        {
            throw new ArgumentOutOfRangeException(nameof(availableWidthPt), availableWidthPt, "Available width must be finite.");
        }

        for (var length = text.Length - 1; length >= 0; length--)
        {
            var candidate = FitPrefix(text, length);
            if (
                metrics.TryMeasure(candidate, out TextMeasurement? measurement)
                && measurement is not null
                && measurement.WidthPt <= availableWidthPt
            )
            {
                return candidate;
            }
        }

        return Ellipsis.ToString();
    }

    /// <summary>
    /// Returns whether a text runs wider than the space available, so the caller
    /// truncates and emits exactly one warning.
    /// </summary>
    /// <param name="text">The text to measure.</param>
    /// <param name="availableWidthPt">The width available in points.</param>
    /// <param name="metrics">The injected deterministic text-metrics seam.</param>
    /// <returns><see langword="true"/> when the measured width exceeds the available width.</returns>
    public static bool Overflows(string text, double availableWidthPt, ITextMetrics metrics)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(metrics);
        return double.IsFinite(availableWidthPt) && availableWidthPt >= 0 && (!metrics.TryMeasure(text, out TextMeasurement? measurement) || measurement is null || measurement.WidthPt > availableWidthPt);
    }

    private static string FitPrefix(string text, int length)
    {
        ReadOnlySpan<char> prefix = text.AsSpan(0, length);
        if (length > 0 && char.IsHighSurrogate(text[length - 1]))
        {
            prefix = prefix[..^1];
        }

        return string.Concat(prefix, Ellipsis.ToString());
    }
}
