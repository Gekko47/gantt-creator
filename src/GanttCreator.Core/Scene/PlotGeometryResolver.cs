namespace GanttCreator.Core.Scene;

/// <summary>Why a plot geometry could not be resolved.</summary>
public enum PlotGeometryRefusal
{
    /// <summary>The preset was absent.</summary>
    NoPreset = 0,

    /// <summary>A preset dimension was non-finite or non-positive.</summary>
    InvalidPresetDimensions = 1,

    /// <summary>The measured text-panel width was non-finite or negative.</summary>
    InvalidTextPanelWidth = 2,

    /// <summary>A margin or border was non-finite or negative.</summary>
    InvalidChrome = 3,

    /// <summary>
    /// The preset leaves no room for a plot wide enough to read. The refusal
    /// carries the shortfall so the message can name it.
    /// </summary>
    InsufficientPlotWidth = 4,

    /// <summary>The resolved plot origin or height was non-finite or degenerate.</summary>
    InvalidOrigin = 5,
}

/// <summary>One resolved plot composition: the budget split, not a measurement.</summary>
/// <param name="Preset">The preset the plot was resolved from.</param>
/// <param name="TextPanelWidthPt">The measured text-panel width, unchanged by resolution.</param>
/// <param name="PlotBounds">The resolved plot rectangle in points.</param>
/// <param name="ShortfallPt">
/// The width the preset could not give the plot, when the resolution was refused.
/// </param>
public sealed record PlotGeometry(
    SizePreset Preset,
    double TextPanelWidthPt,
    RectD PlotBounds,
    double ShortfallPt)
{
    /// <summary>Gets whether a valid plot was resolved.</summary>
    public bool Succeeded => ShortfallPt <= 0;
}

/// <summary>The typed result of resolving plot geometry.</summary>
/// <param name="Geometry">
/// The resolved geometry, or <see langword="null"/> for a refusal that carries no
/// usable figure. <see cref="PlotGeometryRefusal.InsufficientPlotWidth"/> is the
/// one refusal that DOES carry a geometry, because the shortfall it reports is the
/// actionable part of the message.
/// </param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record PlotGeometryOutcome(PlotGeometry? Geometry, PlotGeometryRefusal? Refusal)
{
    /// <summary>
    /// Gets whether a usable plot was resolved. This is deliberately
    /// <see cref="PlotGeometry"/>.Succeeded rather than a null check: the
    /// insufficient-width refusal carries a populated geometry so the caller can
    /// report the shortfall, and a bare null test would report that refusal as a
    /// success.
    /// </summary>
    public bool Succeeded => Geometry is not null && Geometry.Succeeded;
}

/// <summary>
/// The <b>single authority</b> that derives <see cref="RectD"/> plot bounds from a
/// size preset, the measured text-panel width, and the approved chrome
/// (R4.7H D1).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why one authority.</b> D1 exists to prevent a second derivation appearing in
/// a command, a renderer, or a Ribbon callback. Three implementations of the same
/// subtraction would disagree by a margin somewhere, and the disagreement would
/// read as a layout bug rather than as duplicated logic.
/// </para>
/// <para>
/// <b>The formula is fixed (D2):</b>
/// <c>PlotWidth = PresetWidth - TextPanelWidth - Chrome</c>. Everything the preset
/// does not consume becomes plot width.
/// </para>
/// <para>
/// <b>Why the text panel is an input and never an output (D3).</b> A preset that
/// resized a user's text column to make itself fit would silently rewrite the one
/// part of the chart they read. That is a hard product rule, not a preference, so
/// the measured width is carried through untouched.
/// </para>
/// <para>
/// <b>Why insufficient width refuses (D4).</b> When the subtraction leaves nothing,
/// the resolver reports <see cref="PlotGeometryRefusal.InsufficientPlotWidth"/>
/// with the shortfall. It never shrinks the panel, the plot, or the font, because
/// each of those is a silent change to something the user owns.
/// </para>
/// </remarks>
public static class PlotGeometryResolver
{
    /// <summary>
    /// The narrowest plot that can show a readable day column. Below this a bar
    /// cannot be distinguished from its neighbour, so a narrower remainder is a
    /// refusal rather than a rendered chart.
    /// </summary>
    public const double MinimumPlotWidthPt = 24.0;

    /// <summary>
    /// Resolves the plot bounds from a preset and the measured text-panel width.
    /// </summary>
    /// <param name="preset">The selected preset, or <see langword="null"/>.</param>
    /// <param name="textPanelWidthPt">The measured text-panel width in points.</param>
    /// <param name="leftChromePt">Chrome consumed left of the plot (frame/border allowance).</param>
    /// <param name="rightChromePt">Chrome consumed right of the plot.</param>
    /// <param name="topPt">The plot's top coordinate.</param>
    /// <param name="heightPt">The plot height, carried from the preset's budget.</param>
    /// <param name="boundVerticallyToPage">
    /// Whether the preset's page height bounds the plot vertically. <see langword="true"/>
    /// for an EXPORT composition, which is drawn onto a page. <see langword="false"/>
    /// for a LIVE worksheet chart (ADR-0030 D7): there the vertical extent is however
    /// many rows the user has, so page-bounding it would refuse to render a 40-row
    /// schedule merely because A4 landscape is 29 rows tall — the same defect as
    /// lane auto-growth, wearing a page budget instead of a row height.
    /// </param>
    /// <returns>A typed result carrying the resolved bounds or the shortfall.</returns>
    public static PlotGeometryOutcome TryResolve(
        SizePreset? preset,
        double textPanelWidthPt,
        double leftChromePt,
        double rightChromePt,
        double topPt,
        double heightPt,
        bool boundVerticallyToPage = true)
    {
        if (preset is null)
        {
            return Refused(PlotGeometryRefusal.NoPreset);
        }

        if (!double.IsFinite(preset.WidthPt)
            || preset.WidthPt <= 0
            || !double.IsFinite(preset.HeightPt)
            || preset.HeightPt <= 0)
        {
            return Refused(PlotGeometryRefusal.InvalidPresetDimensions);
        }

        // Negative is refused, not clamped: a negative measured width means the
        // measurement is wrong, and clamping it would launder the defect into a
        // plausible-looking plot.
        if (!double.IsFinite(textPanelWidthPt) || textPanelWidthPt < 0)
        {
            return Refused(PlotGeometryRefusal.InvalidTextPanelWidth);
        }

        if (!double.IsFinite(leftChromePt)
            || leftChromePt < 0
            || !double.IsFinite(rightChromePt)
            || rightChromePt < 0)
        {
            return Refused(PlotGeometryRefusal.InvalidChrome);
        }

        if (!double.IsFinite(topPt) || !double.IsFinite(heightPt) || heightPt <= 0)
        {
            return Refused(PlotGeometryRefusal.InvalidOrigin);
        }

        // The vertical bounds must stay INSIDE the preset's own height, exactly as
        // the horizontal check below refuses a plot wider than the page. Without it a
        // caller could place the plot at y = -50, or ask for a height that runs past
        // the bottom of the page, and the resolver would return a rectangle that
        // describes nothing printable -- with no typed refusal to report. Refused
        // rather than clamped: clamping would silently draw a chart somewhere the
        // caller did not ask for, which is the substitution this resolver exists to
        // prevent (D3/D4).
        if (boundVerticallyToPage && (topPt < 0 || topPt + heightPt > preset.HeightPt))
        {
            return Refused(PlotGeometryRefusal.InvalidOrigin);
        }

        // A live chart is not page-bounded, but its origin and height must still be
        // real numbers: a negative top or a non-positive height describes no geometry,
        // and clamping either would place the chart somewhere nobody asked for.
        if (!double.IsFinite(topPt)
            || (!boundVerticallyToPage && topPt < 0)
            || !double.IsFinite(heightPt)
            || heightPt <= 0)
        {
            return Refused(PlotGeometryRefusal.InvalidOrigin);
        }

        var plotWidth = preset.WidthPt - textPanelWidthPt - leftChromePt - rightChromePt;
        if (plotWidth < MinimumPlotWidthPt)
        {
            // Report the shortfall rather than rendering a too-narrow plot. The
            // number makes the message actionable: it tells the user how much a
            // column must be narrowed, instead of merely that it failed.
            var shortfall = MinimumPlotWidthPt - plotWidth;
            return new PlotGeometryOutcome(
                new PlotGeometry(preset, textPanelWidthPt, default, shortfall),
                PlotGeometryRefusal.InsufficientPlotWidth);
        }

        var originX = textPanelWidthPt + leftChromePt;
        return !double.IsFinite(originX) || originX + plotWidth > preset.WidthPt
            ? Refused(PlotGeometryRefusal.InvalidOrigin)
            : new PlotGeometryOutcome(
                new PlotGeometry(preset, textPanelWidthPt, new RectD(originX, topPt, plotWidth, heightPt), 0),
                null);
    }

    private static PlotGeometryOutcome Refused(PlotGeometryRefusal refusal) => new(null, refusal);
}
