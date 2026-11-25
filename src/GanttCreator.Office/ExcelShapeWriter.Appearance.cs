using GanttCreator.Core;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

public partial class ExcelShapeWriter
{
    /// <summary>
    /// Writes the resolved fill and stroke onto a shape.
    /// </summary>
    /// <param name="shape">The shape to style.</param>
    /// <param name="request">The request whose style members are written.</param>
    /// <remarks>
    /// <para>
    /// R4.6. The host's format members are written from
    /// <see cref="OfficeStyleMapper"/>'s output verbatim: no colour is chosen here,
    /// no token is defaulted, and no theme colour is touched. The mapper owns the
    /// arithmetic; this method owns the COM.
    /// </para>
    /// <para>
    /// <see cref="MsoPatternType"/> is the facility the Step-0 probe established for
    /// a hatch, and it is a <em>fixed pitch</em> pattern set: the probe found no
    /// member controlling the <c>HatchPitchPt</c> spacing or the <c>HatchLinePt</c>
    /// stroke width. Those two tokens are therefore <strong>not honoured</strong> by
    /// the native pattern, which is exactly the case entity-guide section 18
    /// anticipates ("may use a native pattern only if the compatibility test proves
    /// equivalent bounds and adequate appearance"). The deviation is recorded for
    /// R8.3's equivalence policy rather than silently approximated, and a
    /// <see cref="GanttHatchPattern"/> the host cannot express is
    /// <strong>refused</strong> rather than drawn as the nearest pattern.
    /// </para>
    /// <para>
    /// <c>internal virtual</c> is a test seam, per the
    /// <see cref="ApplyText"/> pattern: the write reaches COM through a chained
    /// <c>Shape.Fill</c>/<c>Shape.Line</c> walk that a contract test cannot observe,
    /// so the seam lets a test record that the style was applied rather than
    /// inferring it.
    /// </para>
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when a hatch pattern has no host member. Never caught, because an
    /// unexpressible pattern is a contract fault the caller must see, not a host
    /// failure to swallow into a refusal.
    /// </exception>
    internal virtual void ApplyStyle(Excel.Shape shape, OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(shape);
        ArgumentNullException.ThrowIfNull(request);

        // Read-back verification (ADR-0054 amendment 2026-11-20). The live trace measured
        // 969 ms of style writes for 148 shapes through ShapeRange, i.e. the host applies
        // the block per member; a converged shape therefore pays four-to-ten reads and no
        // writes. The comparison covers every member ApplyStyleCore writes, so a mismatch
        // still takes the full block and the two paths cannot disagree.
        OfficeShapeStyle style = OfficeStyleMapper.Map(request);
        Excel.FillFormat? fill = shape.Fill;
        Excel.LineFormat? line = shape.Line;
        if (StyleMatches(fill, line, style))
        {
            return;
        }

        ApplyStyleCore(fill, line, style);
    }

    /// <summary>
    /// Applies the same style to a group of shapes through one <c>ShapeRange</c>.
    /// </summary>
    /// <param name="range">The group's range.</param>
    /// <param name="request">The request whose resolved style the group shares.</param>
    /// <remarks>
    /// Microsoft documents that setting a secondary format object's properties on a
    /// <c>ShapeRange</c> applies them to every shape in the range. The live probe measured
    /// 8.5–9.7 ms for 20 rectangles against 65.8–77.5 ms shape by shape, with zero
    /// read-back mismatches. <see cref="ApplyStyleCore"/> is the same code the single-shape
    /// path uses, so the two cannot drift.
    /// </remarks>
    internal virtual void ApplyStyleToRange(Excel.ShapeRange range, OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(range);
        ArgumentNullException.ThrowIfNull(request);

        ApplyStyleCore(range.Fill, range.Line, OfficeStyleMapper.Map(request));
    }

    /// <summary>
    /// Whether the live fill and line already carry every member the mapped style writes.
    /// </summary>
    /// <param name="fill">The shape's fill, or null when the host exposes none.</param>
    /// <param name="line">The shape's line, or null when the host exposes none.</param>
    /// <param name="style">The style the request resolved to.</param>
    /// <returns><see langword="true"/> when no write is needed.</returns>
    /// <remarks>
    /// Only members the mapper actually resolves are compared, so a <c>null</c> token still
    /// means "the scene resolved no value" rather than "match anything" (R4.7 D5).
    /// </remarks>
    private static bool StyleMatches(Excel.FillFormat? fill, Excel.LineFormat? line, OfficeShapeStyle style)
    {
        if (fill is not null)
        {
            bool fillVisible = fill.Visible == MsoTriState.msoTrue;
            if (fillVisible != style.FillVisible)
            {
                return false;
            }

            if (style.FillVisible)
            {
                if (style.FillRgb is { } foreground
                    && fill.ForeColor is { } foreColour
                    && foreColour.RGB != foreground)
                {
                    return false;
                }

                if (style.FillBackgroundRgb is { } background
                    && fill.BackColor is { } backColour
                    && backColour.RGB != background)
                {
                    return false;
                }

                if (!ApproximatelyEqual(fill.Transparency, style.FillTransparency, 0.001f))
                {
                    return false;
                }

                if (style.HatchPattern is { } hatch)
                {
                    if (fill.Type != MsoFillType.msoFillPatterned || fill.Pattern != MapHatchPattern(hatch))
                    {
                        return false;
                    }
                }
                else if (fill.Type != MsoFillType.msoFillSolid)
                {
                    return false;
                }
            }
        }

        if (line is null)
        {
            return true;
        }

        bool lineVisible = line.Visible == MsoTriState.msoTrue;
        if (lineVisible != style.LineVisible)
        {
            return false;
        }

        if (!style.LineVisible)
        {
            return true;
        }

        bool strokeMatches = style.LineRgb is not { } stroke
            || line.ForeColor is not { } strokeColour
            || strokeColour.RGB == stroke;
        bool weightMatches = style.LineWeightPt is not { } weight
            || ApproximatelyEqual(line.Weight, weight, 0.01f);

        return strokeMatches
            && weightMatches
            && line.DashStyle == MapDashStyle(style.StrokeDash)
            && ApproximatelyEqual(line.Transparency, style.LineTransparency, 0.001f);
    }

    /// <summary>Compares two host float values within an explicit tolerance.</summary>
    private static bool ApproximatelyEqual(float actual, float expected, float tolerance) =>
        Math.Abs(actual - expected) <= tolerance;

    /// <summary>Writes the resolved style onto one fill/line pair.</summary>
    /// <param name="fill">The shape's or range's fill, or null when the host exposes none.</param>
    /// <param name="line">The shape's or range's line, or null when the host exposes none.</param>
    /// <param name="style">The request's mapped style.</param>
    private static void ApplyStyleCore(Excel.FillFormat? fill, Excel.LineFormat? line, OfficeShapeStyle style)
    {
        if (fill is not null)
        {
            if (!style.FillVisible)
            {
                // Explicitly off rather than left alone. A shape the host created
                // with its default fill would otherwise keep that fill, and
                // "the scene resolved no fill" has to be able to mean "no fill".
                fill.Visible = MsoTriState.msoFalse;
            }
            else
            {
                Excel.ColorFormat? backColour = fill.BackColor;
                if (backColour is not null && style.FillBackgroundRgb is { } background)
                {
                    backColour.RGB = background;
                }

                Excel.ColorFormat? foreColour = fill.ForeColor;
                if (foreColour is not null && style.FillRgb is { } foreground)
                {
                    foreColour.RGB = foreground;
                }

                fill.Transparency = style.FillTransparency;
                fill.Visible = MsoTriState.msoTrue;

                // Solid() before Patterned(): the host keeps whichever fill type
                // it last had, so a shape created solid and later given a hatch
                // needs the type set explicitly or the colour write above would
                // land on a solid fill and the hatch would never appear.
                if (style.HatchPattern is { } hatch)
                {
                    fill.Patterned(MapHatchPattern(hatch));
                }
                else
                {
                    fill.Solid();
                }
            }
        }

        if (line is null)
        {
            return;
        }

        if (!style.LineVisible)
        {
            line.Visible = MsoTriState.msoFalse;
            return;
        }

        Excel.ColorFormat? strokeColour = line.ForeColor;
        if (strokeColour is not null && style.LineRgb is { } stroke)
        {
            strokeColour.RGB = stroke;
        }

        if (style.LineWeightPt is { } weight)
        {
            line.Weight = weight;
        }

        // The dash is a resolved style token, so the renderer writes it
        // verbatim rather than choosing one: a line the scene resolved as
        // dashed must reach the host dashed. Solid is written explicitly
        // because the host keeps whichever dash it last had.
        line.DashStyle = MapDashStyle(style.StrokeDash);

        line.Transparency = style.LineTransparency;
        line.Visible = MsoTriState.msoTrue;
    }

    /// <summary>
    /// Maps a resolved stroke dash pattern onto the host's line-dash constant.
    /// </summary>
    /// <param name="dash">The resolved dash pattern.</param>
    /// <returns>The host's dash constant.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when the dash has no host member. Never caught, because a
    /// contract fault is not a host failure: the caller discards the shape
    /// and reports the refusal (the same contract as
    /// <see cref="MapHatchPattern"/>).
    /// </exception>
    internal static MsoLineDashStyle MapDashStyle(GanttStrokeDash dash) => dash switch
    {
        GanttStrokeDash.Solid => MsoLineDashStyle.msoLineSolid,
        GanttStrokeDash.SquareDot => MsoLineDashStyle.msoLineSquareDot,
        GanttStrokeDash.RoundDot => MsoLineDashStyle.msoLineRoundDot,
        GanttStrokeDash.Dash => MsoLineDashStyle.msoLineDash,
        GanttStrokeDash.DashDot => MsoLineDashStyle.msoLineDashDot,
        GanttStrokeDash.LongDash => MsoLineDashStyle.msoLineLongDash,
        _ => throw new ArgumentOutOfRangeException(nameof(dash), dash, "No host dash member."),
    };

    /// <summary>
    /// Applies the style, discarding the shape if the host refuses it.
    /// </summary>
    /// <param name="shape">The created shape.</param>
    /// <param name="request">The request whose style is applied.</param>
    /// <returns>
    /// <see langword="true"/> when the style landed; <see langword="false"/>
    /// when the host refused it and the shape has been discarded.
    /// </returns>
    /// <remarks>
    /// Best-effort on the same reasoning as
    /// <see cref="ApplyOwnershipOrDiscard"/>: a shape that could not be styled
    /// must not be left on the sheet, because the caller is about to be told the
    /// create failed and would never learn about the orphan.
    /// <para>
    /// The outcome is returned rather than swallowed, because
    /// <see cref="Create"/> must not continue past it. Continuing would stamp
    /// ownership onto a shape this method has just deleted, and that write is
    /// either an escaping error or a false success for a shape that no longer
    /// exists. The caller turns <see langword="false"/> into
    /// <see cref="ShapeWriteRefusal.HostRejected"/> and skips the ownership
    /// stamp entirely.
    /// </para>
    /// </remarks>
    private bool ApplyStyleOrDiscard(Excel.Shape shape, OfficeShapeRequest request) =>
        Appearance().ApplyStyleOrDiscard(shape, request);

    /// <summary>
    /// Writes a new text box's content, discarding the shape if the host refuses.
    /// </summary>
    /// <param name="shape">The created text box.</param>
    /// <param name="request">The request whose text members are written.</param>
    /// <returns>
    /// <see langword="true"/> when the content landed; <see langword="false"/>
    /// when the host refused it and the shape has been discarded.
    /// </returns>
    /// <remarks>
    /// <para>
    /// The create-path twin of <see cref="ApplyStyleOrDiscard"/>, and it exists
    /// because a text box is the one kind whose content is written inside
    /// <see cref="AddShape"/>, before the shape is stamped and before
    /// <see cref="Create"/> has any handle on it. A throwing
    /// <see cref="ApplyText"/> there had two bad outcomes: the exception
    /// escaped the create as an unhandled COM error from inside a render
    /// command, and the text box stayed on the sheet with no ownership tag, so
    /// nothing could ever update or remove it.
    /// </para>
    /// <para>
    /// Returning <see langword="false"/> makes <see cref="AddShape"/> return
    /// <see langword="null"/>, which <see cref="Create"/> already reports as
    /// <see cref="ShapeWriteRefusal.HostRejected"/> — the same typed refusal
    /// every other create failure uses, so the caller needs no new case.
    /// </para>
    /// </remarks>
    private bool ApplyTextOrDiscard(Excel.Shape shape, OfficeShapeRequest request) =>
        Appearance().ApplyTextOrDiscard(shape, request);

    /// <summary>
    /// Writes only a duplicated text box's string, discarding it if the host refuses.
    /// </summary>
    /// <param name="shape">The duplicate, whose formatting came from the template.</param>
    /// <param name="request">The request carrying the string to write.</param>
    /// <returns>
    /// <see langword="true"/> when the string landed; <see langword="false"/> when the
    /// host refused it and the duplicate was discarded.
    /// </returns>
    /// <remarks>
    /// A template duplicate already carries every formatting member the full
    /// <see cref="ApplyText"/> writes: the template key includes the font family, size,
    /// weight, alignment, text colour, and the left inset that expresses a grouped
    /// child's indentation. Re-writing them cost 25.7 ms per text box on the measured
    /// 20-activity copy and was 71% of the whole create block; the string is the only
    /// member that differs. The first shape of each key still takes the full path, so
    /// its formatting is never inherited from an empty host default.
    /// </remarks>
    private bool ApplyTextContentOrDiscard(Excel.Shape shape, OfficeShapeRequest request) =>
        Appearance().ApplyTextContentOrDiscard(shape, request);

    /// <summary>
    /// Applies a text box's content on the update path, containing a host failure.
    /// </summary>
    /// <param name="shape">The owned shape whose content is rewritten.</param>
    /// <param name="request">The request whose text members are written.</param>
    /// <returns>
    /// <see langword="true"/> when the content landed; <see langword="false"/>
    /// when the host refused the write.
    /// </returns>
    /// <remarks>
    /// The update-path twin of <see cref="ApplyTextOrDiscard"/>, and it differs in
    /// exactly one respect: there is nothing to discard. The shape was found, not
    /// created, so a failed content write leaves an existing owned shape in place
    /// and the caller is told the update failed.
    /// <para>
    /// <see cref="ArgumentOutOfRangeException"/> is deliberately not caught. It is
    /// what <see cref="MapAlignment"/> raises for a scene alignment with no host
    /// mapping - a contract fault the caller must see, and the same rule
    /// <see cref="ApplyTextOrDiscard"/> follows.
    /// </para>
    /// </remarks>
    private bool TryApplyContent(Excel.Shape shape, OfficeShapeRequest request) =>
        Appearance().TryApplyContent(shape, request);

    /// <summary>
    /// Writes only the string of a text box whose formatting the grouped pass already applied.
    /// </summary>
    /// <param name="shape">The owned text box.</param>
    /// <param name="request">The request carrying the string.</param>
    /// <returns><see langword="true"/> when the string landed; otherwise a host refusal.</returns>
    private bool TryApplyTextContent(Excel.Shape shape, OfficeShapeRequest request) =>
        Appearance().TryApplyTextContent(shape, request);

    /// <summary>
    /// Applies the style on the update path, containing a host failure.
    /// </summary>
    /// <param name="shape">The owned shape whose style is rewritten.</param>
    /// <param name="request">The request whose style members are written.</param>
    /// <returns>
    /// <see langword="true"/> when the style landed; <see langword="false"/> when
    /// the host refused the write.
    /// </returns>
    /// <remarks>
    /// The update-path twin of <see cref="ApplyStyleOrDiscard"/>, and it differs in
    /// exactly one respect: there is nothing to discard, for the same reason
    /// <see cref="TryApplyContent"/> does not discard. <see cref="ApplyGeometry"/>
    /// has already run by this point and is not undone, so a refused style leaves
    /// the shape moved and un-restyled; that is stated on
    /// <see cref="IShapeWritePort.Update(OfficeShapeRequest)"/> rather than left for
    /// a reader to infer.
    /// <para>
    /// <see cref="ArgumentOutOfRangeException"/> is deliberately not caught, for
    /// the reason <see cref="ApplyStyle"/> documents: it is
    /// <see cref="MapHatchPattern"/>'s contract fault, not a host failure.
    /// </para>
    /// </remarks>
    private bool TryApplyStyle(Excel.Shape shape, OfficeShapeRequest request) =>
        Appearance().TryApplyStyle(shape, request);

    /// <summary>
    /// Maps a resolved hatch pattern onto the host's pattern constant.
    /// </summary>
    /// <param name="pattern">The resolved pattern.</param>
    /// <returns>The host pattern constant.</returns>
    /// <remarks>
    /// <para>
    /// The mapping is closed because <see cref="GanttHatchPattern"/> is. The
    /// default arm throws rather than falling back to a nearby pattern, for the
    /// same reason <see cref="MapAlignment"/> does: a pattern the scene asked for
    /// and the host cannot draw must fail loudly, so the guide and this method
    /// cannot drift apart unnoticed. A silent nearest-match would render a
    /// procurement bar in a pattern nobody approved.
    /// </para>
    /// <para>
    /// Direction is read from the entity guide's own wording, not guessed:
    /// <c>ForwardDiagonal</c> is "top-left to bottom-right", which on a
    /// screen-coordinate host is the <em>downward</em> diagonal.
    /// </para>
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the pattern is not a defined member.</exception>
    internal static MsoPatternType MapHatchPattern(GanttHatchPattern pattern) =>
        pattern switch
        {
            GanttHatchPattern.ForwardDiagonal => MsoPatternType.msoPatternLightDownwardDiagonal,
            GanttHatchPattern.BackwardDiagonal => MsoPatternType.msoPatternLightUpwardDiagonal,
            GanttHatchPattern.Cross => MsoPatternType.msoPatternDiagonalCross,

            // None is listed rather than left to the default arm so the intent is
            // explicit: a solid fill is not a hatch, and OfficeStyleMapper turns
            // None into a null pattern rather than calling this. Reaching here
            // means a caller mapped a "no hatch" as though it were one.
            GanttHatchPattern.None => throw new ArgumentOutOfRangeException(
                nameof(pattern),
                pattern,
                "None means a solid fill and must not be mapped to a host pattern."),

            _ => throw new ArgumentOutOfRangeException(
                nameof(pattern),
                pattern,
                "The scene produced a hatch pattern with no host mapping. Extend MapHatchPattern and the guide together."),
        };

}
