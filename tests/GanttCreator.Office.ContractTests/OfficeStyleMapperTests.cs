using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Microsoft.Office.Core;
using Xunit.Abstractions;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R4.6's complete style-token-to-Office-property matrix: every colour token in
/// the entity guide's "Shared colour and typography tokens" table maps to one
/// exact host property value, and the un-mapped-token guard fires when one does
/// not.
/// </summary>
/// <remarks>
/// <para>
/// The expected values are transcribed by hand from the guide's tables and written
/// as literals, exactly as the R4.3 integration test writes its geometry literals.
/// A matrix built by reading the expected value back out of the code under test
/// would be self-consistent by construction: a wrong byte order or a swapped pair
/// would move both sides together and every row would still pass.
/// </para>
/// <para>
/// The host packs a colour least-significant-byte first
/// (<c>blue * 65536 + green * 256 + red</c>, MS-OI29500 6.1.2.7.1.11), which is
/// the opposite of the <c>#RRGGBB</c> order the tokens are authored in. The
/// literals below therefore look "wrong" against the token text by design, and
/// several rows are deliberately asymmetric so a byte swap cannot pass.
/// </para>
/// </remarks>
public class OfficeStyleMapperTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    /// <summary>
    /// Every colour token in the guide's table, with the host value it must
    /// produce. Transcribed from the guide, not read from the catalogue.
    /// </summary>
    public static TheoryData<string, string, int> GuideColourTokens =>
        new()
        {
            // Token, authored hex, expected packed host value.
            { "ActualFill", "#00B0F0", 0xF0B000 },
            { "ActualOutline", "#0070C0", 0xC07000 },
            { "PlannedFill", "#92D050", 0x50D092 },
            { "PlannedOutline", "#548235", 0x358254 },
            { "BaselineFill", "#00B050", 0x50B000 },
            { "BaselineOutline", "#006100", 0x006100 },
            { "CriticalStroke", "#FF0000", 0x0000FF },
            // R4.7C / ADR-0027 D2: CriticalFill shares CriticalStroke's value but
            // differs by role, so the catalogue never carries a stroke token used
            // as a fill. Its host value is therefore identical.
            { "CriticalFill", "#FF0000", 0x0000FF },
            { "CriticalOutline", "#C00000", 0x0000C0 },
            { "DelayFill", "#FF0000", 0x0000FF },
            { "DelayText", "#FFFFFF", 0xFFFFFF },
            { "DefaultText", "#000000", 0x000000 },
            { "ChartBackground", "#FFFFFF", 0xFFFFFF },
            { "DataPanelFill", "#FFFFFF", 0xFFFFFF },
            { "HeaderFill", "#FFFFFF", 0xFFFFFF },
            { "YearHeaderFill", "#D9D9D9", 0xD9D9D9 },
            { "SplitterFill", "#FFE699", 0x99E6FF },
            { "AlternateBandFill", "#F2F2F2", 0xF2F2F2 },
            { "MinorGridStroke", "#D9D9D9", 0xD9D9D9 },
            { "MajorGridStroke", "#000000", 0x000000 },
            { "DelineatorStroke", "#404040", 0x404040 },
            { "WarningFill", "#FFF2CC", 0xCCF2FF },
        };

    [Theory]
    [MemberData(nameof(GuideColourTokens))]
    public void Every_guide_colour_token_maps_to_its_exact_host_value(
        string tokenName,
        string authoredHex,
        int expectedRgb)
    {
        var colour = ColourHex.Parse(authoredHex);

        Assert.Equal(expectedRgb, OfficeStyleMapper.ToOfficeRgb(colour));

        // The token must reach the fill property through a real request, not
        // only through the arithmetic helper: this is the end-to-end token path.
        OfficeShapeStyle style = OfficeStyleMapper.Map(
            RectangleRequest(fill: colour));

        Assert.Equal(expectedRgb, style.FillRgb);
        _output.WriteLine($"{tokenName} {authoredHex} -> 0x{expectedRgb:X6}");
    }

    [Fact]
    public void The_host_byte_order_is_the_inverse_of_the_authored_hex_order()
    {
        // A row that documents the packing rule directly. #010203 has three
        // distinct bytes, so any reordering of the channels is visible.
        Assert.Equal(0x030201, OfficeStyleMapper.ToOfficeRgb(ColourHex.Parse("#010203")));
    }

    [Fact]
    public void A_three_digit_token_cannot_agree_with_its_reversed_packing()
    {
        // The negative control for the byte-order rows: proves the matrix would
        // actually fail on a swap rather than passing by construction.
        const int wrong = 0x010203;
        Assert.NotEqual(wrong, OfficeStyleMapper.ToOfficeRgb(ColourHex.Parse("#010203")));
    }

    [Fact]
    public void An_opaque_token_maps_to_zero_transparency()
    {
        OfficeShapeStyle style = OfficeStyleMapper.Map(
            RectangleRequest(fill: ColourHex.Parse("#92D050")));

        Assert.Equal(0f, style.FillTransparency);
    }

    [Fact]
    public void An_explicit_alpha_token_maps_to_the_inverted_host_transparency_scale()
    {
        // The host's Transparency is 0-opaque while the token's alpha is
        // 255-opaque, so the two scales are inverses. #92D05080 carries alpha
        // 0x80 = 128, giving 1 - 128/255.
        OfficeShapeStyle style = OfficeStyleMapper.Map(
            RectangleRequest(fill: ColourHex.Parse("#92D05080")));

        Assert.Equal(1f - (128f / 255f), style.FillTransparency, 5);
    }

    [Fact]
    public void A_six_digit_and_a_fully_opaque_eight_digit_token_agree_on_transparency()
    {
        // The entity guide states colours carry an explicit FF unless transparency
        // is named, so the two forms must not disagree about an opaque colour.
        var shortForm = OfficeStyleMapper.Map(RectangleRequest(fill: ColourHex.Parse("#92D050")));
        var longForm = OfficeStyleMapper.Map(RectangleRequest(fill: ColourHex.Parse("#92D050FF")));

        Assert.Equal(shortForm.FillTransparency, longForm.FillTransparency);
        Assert.Equal(shortForm.FillRgb, longForm.FillRgb);
    }

    /// <summary>
    /// Every style preset the guide's per-subtype sections name, with the fill and
    /// stroke the guide assigns it. The point is coverage of the <em>families</em>
    /// (planned/actual/baseline/critical/delay/procurement/milestone/delineator/
    /// splitter), not of individual tokens.
    /// </summary>
    public static TheoryData<string, string, string, string> GuideSubtypeStyles =>
        new()
        {
            // Style key, fill token, stroke token, both authored hex values.
            { "AsPlannedActivity", "PlannedFill", "PlannedOutline", "#92D050|#548235" },
            { "AsBuiltActivity", "ActualFill", "ActualOutline", "#00B0F0|#0070C0" },
            { "BaselineActivity", "BaselineFill", "BaselineOutline", "#00B050|#006100" },
            // ADR-0027 D2/D5: the critical overlay is a FILLED rectangle, so its
            // visual contract is the CriticalFill token. It previously read "(none) /
            // CriticalStroke", which described the rejected line representation: a
            // filled rect with no fill carried no visual meaning at all, and the
            // stroke row is what the entity used to be drawn as.
            { "CriticalInterval", "CriticalFill", "CriticalStroke", "#FF0000|#FF0000" },
            { "DelayEvent", "DelayFill", "CriticalOutline", "#FF0000|#C00000" },
            { "AsPlannedProcurement", "PlannedFill", "PlannedOutline", "#92D050|#548235" },
            { "AsBuiltProcurement", "ActualFill", "ActualOutline", "#00B0F0|#0070C0" },
            { "BaselineProcurement", "BaselineFill", "BaselineOutline", "#00B050|#006100" },
            { "AsPlannedMilestone", "PlannedFill", "PlannedOutline", "#92D050|#548235" },
            { "AsBuiltMilestone", "ActualFill", "ActualOutline", "#00B0F0|#0070C0" },
            { "BaselineMilestone", "BaselineFill", "BaselineOutline", "#00B050|#006100" },
            { "CriticalMilestone", "CriticalStroke", "CriticalOutline", "#FF0000|#C00000" },
            { "DefaultDelineator", "(none)", "DelineatorStroke", "|#404040" },
            { "Splitter", "SplitterFill", "(none)", "#FFE699|" },
            { "Spacer", "(none)", "(none)", "|" },
        };

    [Theory]
    [MemberData(nameof(GuideSubtypeStyles))]
    public void Every_guide_subtype_style_maps_to_its_exact_fill_and_stroke(
        string styleKey,
        string fillToken,
        string strokeToken,
        string expectedHexPair)
    {
        ArgumentNullException.ThrowIfNull(expectedHexPair);
        string[] expected = expectedHexPair.Split('|');
        ColourHex? fill = expected[0].Length > 0 ? ColourHex.Parse(expected[0]) : null;
        ColourHex? stroke = expected[1].Length > 0 ? ColourHex.Parse(expected[1]) : null;

        OfficeShapeStyle style = OfficeStyleMapper.Map(
            RectangleRequest(id: $"row-1:{styleKey}", fill: fill, stroke: stroke));

        Assert.Equal(fill is not null, style.FillVisible);
        Assert.Equal(stroke is not null, style.LineVisible);
        Assert.Equal(fill is null ? null : OfficeStyleMapper.ToOfficeRgb(fill), style.FillRgb);
        Assert.Equal(stroke is null ? null : OfficeStyleMapper.ToOfficeRgb(stroke), style.LineRgb);
        _output.WriteLine($"{styleKey}: fill {fillToken}, stroke {strokeToken}");
    }

    /// <summary>
    /// The guide's metric tokens that reach a style property, with the exact host
    /// value. Section 16 fixes the standard outline at <c>StandardOutlinePt</c>.
    /// </summary>
    /// <remarks>
    /// <c>CriticalLinePt</c> is deliberately absent: ADR-0027 D4 retired it from
    /// <c>GanttCatalogues.Metrics</c> when the overlay became a filled rectangle, since
    /// a line thickness is no longer an input to anything. Leaving it here would assert
    /// that a retired token still maps to a host property.
    /// </remarks>
    public static TheoryData<string, double> GuideStrokeWidths =>
        new()
        {
            { "StandardOutlinePt", 0.75 },
            { "GridLinePt", 0.5 },
            { "MajorBoundaryPt", 1.0 },
            { "DelineatorLinePt", 0.75 },
        };

    [Theory]
    [MemberData(nameof(GuideStrokeWidths))]
    public void Every_guide_stroke_width_token_maps_to_the_exact_host_weight(
        string tokenName,
        double tokenValue)
    {
        OfficeShapeStyle style = OfficeStyleMapper.Map(
            RectangleRequest(stroke: ColourHex.Parse("#FF0000"), width: tokenValue));

        Assert.Equal((float)tokenValue, style.LineWeightPt);
        _output.WriteLine($"{tokenName} {tokenValue}pt -> Weight {style.LineWeightPt}");
    }

    [Fact]
    public void An_absent_stroke_width_stays_absent_rather_than_becoming_a_default()
    {
        // Section 12 says "Fill and outline are explicit; no renderer defaults."
        // Substituting a width would be exactly that default.
        OfficeShapeStyle style = OfficeStyleMapper.Map(
            RectangleRequest(stroke: ColourHex.Parse("#548235"), width: null));

        Assert.Null(style.LineWeightPt);
    }

    [Fact]
    public void An_absent_colour_stays_absent_rather_than_becoming_white()
    {
        OfficeShapeStyle style = OfficeStyleMapper.Map(
            RectangleRequest(fill: null, stroke: null));

        Assert.False(style.FillVisible);
        Assert.False(style.LineVisible);
        Assert.Null(style.FillRgb);
        Assert.Null(style.LineRgb);
    }

    [Fact]
    public void A_line_is_stroke_only_even_when_the_scene_resolved_a_fill()
    {
        // Section 16's critical overlay and section 24's delineator are lines. A
        // fill is meaningless on a line, so forwarding one would assert an
        // appearance the entity does not have.
        OfficeShapeStyle style = OfficeStyleMapper.Map(
            new OfficeShapeRequest(
                "chart:critical:0",
                OfficeShapeKind.Line,
                new OfficeShapeGeometry(From: new PointD(0, 0), To: new PointD(10, 0)),
                ZLayer.CriticalOverlay,
                FillColour: ColourHex.Parse("#FF0000"),
                StrokeColour: ColourHex.Parse("#FF0000"),
                LineWidthPt: 2.25));

        Assert.False(style.FillVisible);
        Assert.Null(style.FillRgb);
        Assert.True(style.LineVisible);
        Assert.Equal(2.25f, style.LineWeightPt);
    }

    [Fact]
    public void A_rectangle_a_text_box_and_a_diamond_each_carry_a_fill()
    {
        // The three kinds with an interior. Pinned individually because
        // CarriesFill is a closed switch and a new kind must be a deliberate act.
        Assert.True(OfficeStyleMapper.CarriesFill(OfficeShapeKind.Rectangle));
        Assert.True(OfficeStyleMapper.CarriesFill(OfficeShapeKind.TextBox));
        Assert.True(OfficeStyleMapper.CarriesFill(OfficeShapeKind.Diamond));
        Assert.False(OfficeStyleMapper.CarriesFill(OfficeShapeKind.Line));
    }

    /// <summary>
    /// The un-mapped-token guard. Every pattern the scene can resolve must have a
    /// host member; a new <see cref="GanttHatchPattern"/> member without one fails
    /// here rather than rendering as the nearest approved pattern.
    /// </summary>
    [Fact]
    public void Every_scene_hatch_pattern_has_a_host_mapping()
    {
        foreach (GanttHatchPattern pattern in Enum.GetValues<GanttHatchPattern>())
        {
            if (pattern == GanttHatchPattern.None)
            {
                // None is a solid fill, not a pattern; it must never be mapped.
                Assert.Throws<ArgumentOutOfRangeException>(
                    () => ExcelShapeWriter.MapHatchPattern(pattern));
                continue;
            }

            // The call is the assertion: an unmapped member throws rather than
            // returning a default, so this loop fails the moment a pattern is
            // added to the scene enum without a host member.
            MsoPatternType mapped = ExcelShapeWriter.MapHatchPattern(pattern);
            _output.WriteLine($"{pattern} -> {mapped}");
            Assert.True(Enum.IsDefined(mapped));
        }
    }

    [Fact]
    public void Each_hatch_pattern_maps_to_its_own_distinct_host_member()
    {
        // Two patterns sharing a member would be a silent collapse: the guide
        // distinguishes forward, backward, and cross, so the host must too.
        MsoPatternType forward = ExcelShapeWriter.MapHatchPattern(GanttHatchPattern.ForwardDiagonal);
        MsoPatternType backward = ExcelShapeWriter.MapHatchPattern(GanttHatchPattern.BackwardDiagonal);
        MsoPatternType cross = ExcelShapeWriter.MapHatchPattern(GanttHatchPattern.Cross);

        Assert.NotEqual(forward, backward);
        Assert.NotEqual(forward, cross);
        Assert.NotEqual(backward, cross);
    }

    [Theory]
    [InlineData(GanttHatchPattern.ForwardDiagonal, MsoPatternType.msoPatternLightDownwardDiagonal)]
    [InlineData(GanttHatchPattern.BackwardDiagonal, MsoPatternType.msoPatternLightUpwardDiagonal)]
    [InlineData(GanttHatchPattern.Cross, MsoPatternType.msoPatternDiagonalCross)]
    public void A_hatch_pattern_maps_to_the_direction_the_guide_names(
        GanttHatchPattern pattern,
        MsoPatternType expected)
    {
        // ForwardDiagonal is "top-left to bottom-right" per the token table,
        // which on a screen-coordinate host is the DOWNWARD diagonal. This is the
        // row that would fail if the direction were guessed the other way round.
        Assert.Equal(expected, ExcelShapeWriter.MapHatchPattern(pattern));
    }

    /// <summary>
    /// The guard's positive test: a pattern with no host member must throw rather
    /// than silently drawing the nearest pattern.
    /// </summary>
    [Fact]
    public void A_hatch_pattern_with_no_host_mapping_throws_rather_than_approximating()
    {
        Assert.Throws<ArgumentOutOfRangeException>(
            () => ExcelShapeWriter.MapHatchPattern((GanttHatchPattern)99));
    }

    [Fact]
    public void A_solid_fill_carries_no_pattern_rather_than_none()
    {
        // None must be normalised to a null pattern so the writer takes the
        // Solid() branch. Leaving it as None would make MapHatchPattern throw on
        // the commonest case in the product.
        OfficeShapeStyle style = OfficeStyleMapper.Map(
            RectangleRequest(fill: ColourHex.Parse("#92D050"), hatch: GanttHatchPattern.None));

        Assert.Null(style.HatchPattern);
    }

    [Theory]
    [InlineData(GanttHatchPattern.ForwardDiagonal)]
    [InlineData(GanttHatchPattern.BackwardDiagonal)]
    [InlineData(GanttHatchPattern.Cross)]
    public void A_hatched_fill_reaches_the_style_as_the_pattern_the_scene_resolved(
        GanttHatchPattern pattern)
    {
        OfficeShapeStyle style = OfficeStyleMapper.Map(
            RectangleRequest(
                fill: ColourHex.Parse("#92D050"),
                stroke: ColourHex.Parse("#548235"),
                hatch: pattern));

        Assert.Equal(pattern, style.HatchPattern);
    }

    [Fact]
    public void A_hatch_keeps_the_scenes_fill_and_stroke_rather_than_substituting_its_own()
    {
        // Section 18: the hatch is drawn "using the selected planned/actual/baseline
        // colour family". The family colours are the scene's; the pattern only
        // changes how the fill is drawn.
        OfficeShapeStyle style = OfficeStyleMapper.Map(
            RectangleRequest(
                fill: ColourHex.Parse("#00B0F0"),
                stroke: ColourHex.Parse("#0070C0"),
                hatch: GanttHatchPattern.ForwardDiagonal));

        Assert.Equal(0xF0B000, style.FillRgb);
        Assert.Equal(0xC07000, style.LineRgb);
    }

    [Fact]
    public void A_hatched_fill_leaves_the_pattern_background_unset_rather_than_choosing_white()
    {
        // The host requires a background colour for a patterned fill and will
        // apply one. Naming it here as null keeps the fact that the scene did not
        // resolve one visible to R8.3's equivalence policy, instead of baking a
        // renderer-chosen colour into the matrix.
        OfficeShapeStyle style = OfficeStyleMapper.Map(
            RectangleRequest(fill: ColourHex.Parse("#92D050"), hatch: GanttHatchPattern.ForwardDiagonal));

        Assert.Null(style.FillBackgroundRgb);
    }

    [Fact]
    public void Mapping_refuses_a_null_request()
    {
        Assert.Throws<ArgumentNullException>(() => OfficeStyleMapper.Map(null!));
    }

    [Theory]
    [MemberData(nameof(OfficeStyleMapperTests.GuideColourTokens))]
    public void Every_guide_colour_token_is_parsable_as_a_scene_colour(
        string tokenName,
        string authoredHex,
        int expectedRgb)
    {
        // The matrix rows assume the token text is in the form the scene carries.
        // If the guide's form and ColourHex's accepted form diverged, every other
        // row in this file would be asserting against a value the scene can never
        // produce.
        Assert.False(string.IsNullOrWhiteSpace(tokenName));
        Assert.True(ColourHex.TryParse(authoredHex, out ColourHex? parsed));
        Assert.Equal(expectedRgb, OfficeStyleMapper.ToOfficeRgb(parsed!));
    }

    [Fact]
    public void The_matrix_covers_every_colour_token_the_catalogue_publishes()
    {
        // The completeness guard for D4: if a twenty-third colour token is added
        // to the guide's table, this fails until a row is transcribed for it.
        // Without it, "the full matrix" would silently become "the matrix as it
        // was when it was written".
        var covered = GuideColourTokens
            .Select(row => Assert.IsType<string>((object)row[0]!))
            .ToHashSet(StringComparer.Ordinal);

        var published = GanttCatalogues.Colours
            .Select(token => token.Name)
            .ToHashSet(StringComparer.Ordinal);

        Assert.Equal(
            published.OrderBy(name => name, StringComparer.Ordinal),
            covered.OrderBy(name => name, StringComparer.Ordinal));
    }

    [Fact]
    public void The_matrix_covers_every_built_in_style_preset()
    {
        // The subtype rows must stay in step with the code-owned preset list, or
        // a new entity type renders uncompared against the visual specification.
        var covered = GuideSubtypeStyles
            .Select(row => Assert.IsType<string>((object)row[0]!))
            .ToHashSet(StringComparer.Ordinal);

        var published = GanttCatalogues.StylePresets
            .Select(preset => preset.StyleKey)
            .ToHashSet(StringComparer.Ordinal);

        Assert.Equal(
            published.OrderBy(name => name, StringComparer.Ordinal),
            covered.OrderBy(name => name, StringComparer.Ordinal));
    }

    [Fact]
    public void A_delay_label_reaches_the_request_and_maps_to_the_exact_host_value()
    {
        // Entity guide section 17's inside case: DelayText (#FFFFFF) is the one
        // token where a byte-order error is invisible, so the asymmetric #123456
        // row below is the one that actually pins the packing.
        OfficeShapeStyle style = OfficeStyleMapper.Map(
            LabelRequest(textColour: ColourHex.Parse("#FFFFFF")));

        Assert.Equal(0xFFFFFF, style.TextRgb);
    }

    [Theory]
    [InlineData("#FFFFFF", 0xFFFFFF)]
    [InlineData("#000000", 0x000000)]
    [InlineData("#123456", 0x563412)]
    [InlineData("#FF0000", 0x0000FF)]
    [InlineData("#92D050", 0x50D092)]
    public void Every_text_colour_token_packs_with_the_same_byte_order_as_a_fill(
        string authoredHex,
        int expectedRgb)
    {
        // The same ToOfficeRgb serves fill, stroke, and text, so the three cannot
        // disagree about channel order. #123456 has three distinct bytes, so any
        // reordering is visible; the delay red row is the same value the fill
        // matrix already pins, which is what proves the two paths agree.
        OfficeShapeRequest request = LabelRequest(textColour: ColourHex.Parse(authoredHex));

        Assert.Equal(expectedRgb, OfficeStyleMapper.Map(request).TextRgb);
        Assert.Equal(expectedRgb, OfficeStyleMapper.Map(RectangleRequest(fill: ColourHex.Parse(authoredHex))).FillRgb);
    }

    [Fact]
    public void An_unresolved_text_colour_stays_null_rather_than_becoming_black()
    {
        // The scene resolved no text colour, so the adapter must leave the font
        // alone. Substituting DefaultText here would silently restyle every label
        // whose style happened not to name a colour, and it would do it in the
        // renderer - the substitution the scene-first rule forbids.
        OfficeShapeStyle style = OfficeStyleMapper.Map(LabelRequest(textColour: null));

        Assert.Null(style.TextRgb);
    }

    [Fact]
    public void A_text_colour_reaches_the_request_without_being_disturbed_by_the_shape_fill()
    {
        // A delay body is a red rectangle whose label is white. Both are on the
        // same request, so this row fails if either value overwrites the other.
        OfficeShapeStyle style = OfficeStyleMapper.Map(
            new OfficeShapeRequest(
                "row-1:label",
                OfficeShapeKind.TextBox,
                new OfficeShapeGeometry(Bounds: new RectD(10, 20, 100, 10)),
                ZLayer.Label,
                FillColour: null,
                TextColour: ColourHex.Parse("#FFFFFF")));

        Assert.Equal(0xFFFFFF, style.TextRgb);
        Assert.False(style.FillVisible);
    }

    /// <summary>
    /// A label box is never painted, whatever fill or stroke its style names.
    /// </summary>
    /// <remarks>
    /// The live defect: <c>ExcelSceneBuildRequestFactory</c> built the label style as
    /// <c>new SceneStyle("DefaultText", fillColour: #000000)</c> — the colour went to
    /// the FILL, not the text. Because a TextBox "carries a fill", every description
    /// and date label reached Excel as an opaque black rectangle. This asserts the
    /// rule that prevents a recurrence: the entity decides, not the token, so even a
    /// style that explicitly names both a fill and a stroke paints nothing.
    /// </remarks>
    [Fact]
    public void A_label_is_never_filled_or_stroked_even_when_its_style_names_both()
    {
        OfficeShapeRequest request = new(
            "row-1:label",
            OfficeShapeKind.TextBox,
            new OfficeShapeGeometry(Bounds: new RectD(10, 20, 100, 18)),
            ZLayer.Label,
            Text: "abc",
            FillColour: ColourHex.Parse("#000000"),
            StrokeColour: ColourHex.Parse("#FF0000"),
            LineWidthPt: 1,
            TextColour: ColourHex.Parse("#000000"));

        OfficeShapeStyle style = OfficeStyleMapper.Map(request);

        Assert.False(style.FillVisible);
        Assert.Null(style.FillRgb);
        Assert.Null(style.FillBackgroundRgb);
        Assert.Equal(0f, style.FillTransparency);
        Assert.Null(style.HatchPattern);

        Assert.False(style.LineVisible);
        Assert.Null(style.LineRgb);
        Assert.Null(style.LineWeightPt);

        // The text colour is untouched: suppressing the fill must not suppress the
        // font, or the label would become invisible rather than merely unboxed.
        Assert.Equal(0x000000, style.TextRgb);
    }

    /// <summary>
    /// The same suppression must not leak onto the kinds that legitimately carry a
    /// fill and a stroke.
    /// </summary>
    [Fact]
    public void A_body_rectangle_still_reports_its_fill_and_stroke()
    {
        OfficeShapeStyle style = OfficeStyleMapper.Map(
            RectangleRequest(
                fill: ColourHex.Parse("#92D050"),
                stroke: ColourHex.Parse("#548235"),
                width: 0.75));

        Assert.True(style.FillVisible);
        Assert.Equal(0x50D092, style.FillRgb);
        Assert.True(style.LineVisible);
        Assert.Equal(0x358254, style.LineRgb);
    }

    private static OfficeShapeRequest LabelRequest(ColourHex? textColour) =>
        new(
            "row-1:label",
            OfficeShapeKind.TextBox,
            new OfficeShapeGeometry(Bounds: new RectD(10, 20, 100, 10)),
            ZLayer.Label,
            Text: "Delay",
            TextColour: textColour);

    private static OfficeShapeRequest RectangleRequest(
        string id = "row-1:bar",
        ColourHex? fill = null,
        ColourHex? stroke = null,
        double? width = null,
        GanttHatchPattern hatch = GanttHatchPattern.None) =>
        new(
            id,
            OfficeShapeKind.Rectangle,
            new OfficeShapeGeometry(Bounds: new RectD(10, 20, 100, 30)),
            ZLayer.ActivityBody,
            FillColour: fill,
            StrokeColour: stroke,
            LineWidthPt: width,
            HatchPattern: hatch);
}
