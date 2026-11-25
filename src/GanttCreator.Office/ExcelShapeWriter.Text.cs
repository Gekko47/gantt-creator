using GanttCreator.Core;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

public partial class ExcelShapeWriter
{
    /// <summary>
    /// Writes the text content, typography, and alignment onto a text shape.
    /// </summary>
    /// <param name="shape">The created text box.</param>
    /// <param name="request">The request whose text members are written.</param>
    /// <remarks>
    /// <para>
    /// R4.4: the renderer <strong>consumes</strong> the scene's resolved text.
    /// The string is written verbatim, including any ellipsis R3.6's overflow
    /// policy already applied, and the resolved <c>TextBounds</c> positions the
    /// shape. Nothing here measures, truncates, re-wraps, or re-selects a label
    /// side; the scene owns all of that.
    /// </para>
    /// <para>
    /// The <c>TextFrame2</c> path is used rather than the legacy
    /// <c>TextFrame.Characters</c> path because the latter's font members are
    /// typed <c>Object</c>, and this repository treats an untyped interop member
    /// as a live hazard - R4.1's <c>Shape.Tag</c> and R4.2's
    /// <c>Application.Calculation</c> were both assumptions about interop member
    /// shapes that did not survive contact. R4.4 D3's original
    /// <c>xlAlignLeft</c>-style constants do not exist in the installed
    /// assembly; the real values are <see cref="MsoParagraphAlignment"/> members.
    /// </para>
    /// <para>
    /// <c>internal virtual</c> is a test seam, per the
    /// <c>ExcelWorkbookInitialiser</c> pattern: the content write reaches COM
    /// through a chained <c>Shape.TextFrame2</c> walk that a contract test
    /// cannot observe, so the seam lets a test record that the update path
    /// applied the content rather than inferring it from the geometry.
    /// </para>
    /// <para>
    /// Word wrap is switched off deliberately. A text box that auto-fits would
    /// resize itself away from the bounds the scene resolved, which is the same
    /// class of silent re-layout the no-remeasure rule forbids.
    /// </para>
    /// </remarks>
    internal virtual void ApplyText(Excel.Shape shape, OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(shape);
        ArgumentNullException.ThrowIfNull(request);

        Excel.TextFrame2? frame = shape.TextFrame2;
        if (frame is null)
        {
            return;
        }

        TextRange2? textRange = frame.TextRange;
        if (textRange is null)
        {
            return;
        }

        ApplyTextFormatting(frame, textRange, request);

        // Verbatim: the scene already applied the overflow policy.
        textRange.Text = request.Text ?? string.Empty;
    }

    /// <summary>
    /// Writes a text box's formatting without writing its string.
    /// </summary>
    /// <param name="shape">The text box whose formatting is re-applied.</param>
    /// <param name="request">The request whose typography is applied.</param>
    internal virtual void ApplyTextFormatting(Excel.Shape shape, OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(shape);
        ArgumentNullException.ThrowIfNull(request);

        Excel.TextFrame2? frame = shape.TextFrame2;
        if (frame is null)
        {
            return;
        }

        TextRange2? textRange = frame.TextRange;
        if (textRange is null)
        {
            return;
        }

        ApplyTextFormatting(frame, textRange, request);
    }

    /// <summary>
    /// Writes a group's text formatting once, without writing any string.
    /// </summary>
    /// <param name="range">The group's range.</param>
    /// <param name="request">The request whose typography the group shares.</param>
    /// <remarks>
    /// The host requires a single shape for <c>TextRange.Text</c>, so content stays
    /// per-shape; every formatting member can be applied to the range. The live probe
    /// measured 35.9–37.9 ms for 15 text boxes against 164.7–211.5 ms shape by shape,
    /// with zero read-back mismatches.
    /// </remarks>
    internal virtual void ApplyTextFormattingToRange(Excel.ShapeRange range, OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(range);
        ArgumentNullException.ThrowIfNull(request);

        Excel.TextFrame2? frame = range.TextFrame2;
        if (frame is null)
        {
            return;
        }

        TextRange2? textRange = frame.TextRange;
        if (textRange is null)
        {
            return;
        }

        ApplyTextFormatting(frame, textRange, request);
    }

    /// <summary>
    /// Whether the live text frame and font already carry every member the typography
    /// write sets.
    /// </summary>
    /// <param name="shape">The text box to compare.</param>
    /// <param name="request">The request whose typography would be written.</param>
    /// <returns><see langword="true"/> when no typography write is needed.</returns>
    internal virtual bool TextFormattingMatches(Excel.Shape shape, OfficeShapeRequest request)
    {
        Excel.TextFrame2? frame = shape.TextFrame2;
        if (frame is null)
        {
            return false;
        }

        TextRange2? textRange = frame.TextRange;
        if (textRange is null)
        {
            return false;
        }

        if (frame.WordWrap != MsoTriState.msoFalse
            || frame.AutoSize != MsoAutoSize.msoAutoSizeNone
            || frame.VerticalAnchor != MsoVerticalAnchor.msoAnchorMiddle)
        {
            return false;
        }

        if (!ApproximatelyEqual(frame.MarginLeft, (float)request.LeftInsetPt, 0.01f)
            || !ApproximatelyEqual(frame.MarginRight, 0f, 0.01f)
            || !ApproximatelyEqual(frame.MarginTop, 0f, 0.01f)
            || !ApproximatelyEqual(frame.MarginBottom, 0f, 0.01f))
        {
            return false;
        }

        if (request.Alignment is { } alignment
            && textRange.ParagraphFormat is { } paragraph
            && paragraph.Alignment != MapAlignment(alignment))
        {
            return false;
        }

        if (textRange.Font is not { } font)
        {
            return true;
        }

        if (request.FontFamily is { } family && !string.Equals(font.Name, family, StringComparison.Ordinal))
        {
            return false;
        }

        if (request.FontSizePt is { } size && !ApproximatelyEqual(font.Size, (float)size, 0.01f))
        {
            return false;
        }

        if (request.Bold is { } bold)
        {
            bool liveBold = font.Bold == MsoTriState.msoTrue;
            if (liveBold != bold)
            {
                return false;
            }
        }

        if (request.TextColour is { } textColour && font.Fill is { } fontFill)
        {
            int rgb = OfficeStyleMapper.ToOfficeRgb(textColour);
            if (fontFill.ForeColor is { } fore && fore.RGB != rgb)
            {
                return false;
            }

            if (fontFill.Visible != MsoTriState.msoTrue)
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>Whether a text box's live string already equals the request's.</summary>
    /// <param name="shape">The text box to compare.</param>
    /// <param name="expected">The scene's string, or null for an empty one.</param>
    /// <returns><see langword="true"/> when the string needs no write.</returns>
    internal virtual bool TextMatches(Excel.Shape shape, string? expected)
    {
        Excel.TextFrame2? frame = shape.TextFrame2;
        TextRange2? textRange = frame?.TextRange;
        return textRange is not null
            && string.Equals(textRange.Text ?? string.Empty, expected ?? string.Empty, StringComparison.Ordinal);
    }

    /// <summary>Writes the shared text formatting onto one frame/text-range pair.</summary>
    /// <param name="frame">The shape's or range's text frame.</param>
    /// <param name="textRange">The frame's text range.</param>
    /// <param name="request">The request whose typography is applied.</param>
    private static void ApplyTextFormatting(
        Excel.TextFrame2 frame,
        TextRange2 textRange,
        OfficeShapeRequest request)
    {
        frame.WordWrap = MsoTriState.msoFalse;
        frame.AutoSize = MsoAutoSize.msoAutoSizeNone;

        // Vertical centring, which the entity guide's section 22 requires at EVERY
        // label position ("Right: ... vertically centred", "Left: ... vertically
        // centred"). This was never written, so Excel fell back to its own default,
        // msoAnchorTop, and every label sat at the top of the row-tall box the scene
        // resolved for it - visible in a live F5 session on 2026-10-04 as labels
        // hugging the top edge of their bar rather than centred in it.
        //
        // Set explicitly rather than relied upon, for the same reason the margins
        // below are zeroed: the scene already resolved the box and the vertical
        // placement, so a host default is a second authority that can disagree with
        // it. Verified against the installed interop assembly (office.dll 16.0) that
        // TextFrame2.VerticalAnchor is a settable Single-valued property of type
        // MsoVerticalAnchor, and that msoAnchorMiddle is the member named Middle
        // (value 3).
        frame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;

        // The internal margins are zeroed because the scene already resolved the
        // box, and a host default inset would silently move the text away from the
        // position the scene chose - the same no-remeasure violation as auto-sizing.
        // Excel's textbox default is roughly 0.1in left/right and 0.05in top/bottom,
        // which is 3.6pt/7.2pt: enough to make a Right-anchored label start visibly
        // right of the bar it labels, and to make a Left-anchored one end short of
        // it. Verified against the installed interop assembly that all four members
        // are settable Single properties on TextFrame2.
        //
        // The LEFT margin is the scene's resolved inset, not zero. A grouped child's
        // data-panel Description cell carries a positive inset so the copied cell
        // reproduces the worksheet's IndentLevel; every other primitive resolves zero
        // and keeps the established behaviour exactly (SceneText.LeftInsetPt defaults
        // to zero).
        frame.MarginLeft = (float)request.LeftInsetPt;
        frame.MarginRight = 0f;
        frame.MarginTop = 0f;
        frame.MarginBottom = 0f;

        if (request.Alignment is { } alignment && textRange.ParagraphFormat is { } paragraphFormat)
        {
            // ParagraphFormat2 is held in a local rather than written through
            // textRange.ParagraphFormat: a chained COM property call creates a
            // second proxy the adapter never names, and the ownership rule is that
            // every proxy is held in a local. The proxy is Excel-owned and shared,
            // so it is not force-released - see the class remarks.
            paragraphFormat.Alignment = MapAlignment(alignment);
        }

        if (textRange.Font is { } font)
        {
            // Same rule for Font2, which is read once and written through for
            // family, size, and weight alike.
            if (request.FontFamily is { } family)
            {
                font.Name = family;
            }

            if (request.FontSizePt is { } size)
            {
                font.Size = (float)size;
            }

            if (request.Bold is { } bold)
            {
                font.Bold = bold ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            }
        }

        // The text colour goes through the FONT's fill, not the shape's: the shape
        // fill is already spoken for by ApplyStyle, and the two are independent
        // properties on the host. A shape can therefore be a red delay body with a
        // white label, which is exactly what entity guide section 17 requires.
        //
        // The chain is TextRange2.Font (Font2) -> Fill (FillFormat) -> ForeColor
        // (ColorFormat) -> RGB, verified by reflection against the installed
        // office.dll 16.0.0 rather than assumed: Font2 exposes UnderlineColor
        // directly but has no Font.Color, so the fill is the only colour surface
        // on a Font2. Every proxy is read once into a local and written through,
        // per the ownership rule, and none is force-released because they are
        // Excel-owned and shared.
        OfficeShapeStyle style = OfficeStyleMapper.Map(request);
        if (style.TextRgb is { } textRgb
            && textRange.Font is { } textFont
            && textFont.Fill is { } textFill)
        {
            // Solid() before the colour write, for the same reason ApplyStyle does
            // it for a shape fill: the host keeps whichever fill type it last had,
            // and a text fill inherited from the theme can silently ignore a
            // ForeColor write. Making it solid first is what makes the colour land.
            textFill.Solid();
            if (textFill.ForeColor is { } textForeColor)
            {
                textForeColor.RGB = textRgb;
            }

            textFill.Visible = MsoTriState.msoTrue;
        }
    }

    /// <summary>
    /// Writes only a duplicated text box's string. Test seam over the one content write.
    /// </summary>
    /// <param name="shape">The duplicate, whose formatting came from its template.</param>
    /// <param name="request">The request carrying the string to write.</param>
    /// <remarks>
    /// <para>
    /// The template key guarantees the font family, size, weight, alignment, text colour,
    /// and left inset are identical, so re-writing them would be a second authority over
    /// values the host already holds. That re-write cost 25.7 ms per text box on the
    /// measured 20-activity copy and was 71% of the create block.
    /// </para>
    /// <para>
    /// The left inset is load-bearing for a grouped child's indentation: it is part of the
    /// key, so a child never duplicates a non-indented template. A host that exposes no
    /// text frame is a refusal, not a silent no-op: the duplicate would otherwise keep the
    /// template's string.
    /// </para>
    /// </remarks>
    /// <exception cref="InvalidOperationException">Thrown when the host exposes no text frame or text range.</exception>
    internal virtual void ApplyTextContent(Excel.Shape shape, OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(shape);
        ArgumentNullException.ThrowIfNull(request);

        Excel.TextFrame2 frame = shape.TextFrame2
            ?? throw new InvalidOperationException("The host did not expose a text frame for the duplicated text box.");
        TextRange2 textRange = frame.TextRange
            ?? throw new InvalidOperationException("The host did not expose a text range for the duplicated text box.");

        textRange.Text = request.Text ?? string.Empty;
    }

    /// <summary>
    /// Maps a resolved scene alignment onto the host's paragraph-alignment
    /// constant.
    /// </summary>
    /// <param name="alignment">The alignment the scene resolved.</param>
    /// <returns>The host alignment constant.</returns>
    /// <remarks>
    /// The mapping is closed because <see cref="GanttTextAlignment"/> is. The
    /// default arm is a throw rather than a silent fallback: an alignment added
    /// to the scene enum must fail loudly here, so the guide and this method
    /// cannot drift apart unnoticed.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the alignment is not a defined member.</exception>
    private static MsoParagraphAlignment MapAlignment(GanttTextAlignment alignment) =>
        alignment switch
        {
            GanttTextAlignment.Left => MsoParagraphAlignment.msoAlignLeft,
            GanttTextAlignment.Centre => MsoParagraphAlignment.msoAlignCenter,
            GanttTextAlignment.Right => MsoParagraphAlignment.msoAlignRight,
            _ => throw new ArgumentOutOfRangeException(
                nameof(alignment),
                alignment,
                "The scene produced an alignment with no host mapping. Extend MapAlignment and the guide together."),
        };

    // Explicit implementations of the appearance host surface. They forward to the
    // internal virtual seams above, so a test subclass's override is still observed and
    // the collaborator never needs a reference to this class.
    void IShapeAppearanceHost.ApplyStyle(Excel.Shape shape, OfficeShapeRequest request) =>
        ApplyStyle(shape, request);

    void IShapeAppearanceHost.ApplyText(Excel.Shape shape, OfficeShapeRequest request) =>
        ApplyText(shape, request);

    void IShapeAppearanceHost.ApplyTextContent(Excel.Shape shape, OfficeShapeRequest request) =>
        ApplyTextContent(shape, request);

    bool IShapeAppearanceHost.TextFormattingMatches(Excel.Shape shape, OfficeShapeRequest request) =>
        TextFormattingMatches(shape, request);

    bool IShapeAppearanceHost.TextMatches(Excel.Shape shape, string? expected) =>
        TextMatches(shape, expected);

    void IShapeAppearanceHost.DiscardUnownedShape(Excel.Shape shape) =>
        DiscardUnownedShape(shape);
}
