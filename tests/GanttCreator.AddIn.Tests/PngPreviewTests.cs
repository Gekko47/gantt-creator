using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Office;
using GanttCreator.Raster;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// R8.7 D1/D2/D3/D6: the pure PNG width-preview math. The preview is delegated
/// math — every dimension comes from <see cref="ExportSize"/> — so these facts
/// pin the delegation, the number-plus-unit composition (D3), the validation
/// matrix with one positive per violation, the culture-invariant formatting, and
/// the unit/preset item mapping.
/// </summary>
public sealed class PngPreviewTests
{
    /// <summary>An A4-landscape-shaped export aspect (841.89 × 479.5 pt).</summary>
    private static readonly ExportAspectPt LandscapeAspect = new(841.8897637795276, 479.5);

    /// <summary>
    /// The preview's pixel dimensions equal <see cref="ExportSize.ToPixels"/> for the
    /// same number, unit and aspect — the equality pin that forbids a second formula —
    /// and the formatted text carries exactly those numbers.
    /// </summary>
    [Theory]
    [InlineData("10", ExportUnit.Centimetres)]
    [InlineData("4", ExportUnit.Inches)]
    [InlineData("800", ExportUnit.Pixels)]
    [InlineData("29.7", ExportUnit.Centimetres)]
    [InlineData("  15 ", ExportUnit.Centimetres)]
    public void The_preview_dimensions_equal_the_export_size_output(string number, ExportUnit unit)
    {
        PngEvaluation evaluation = PngPreview.Evaluate(number, unit, LandscapeAspect);

        Assert.True(evaluation.Valid, evaluation.Message);
        Assert.True(PngPreview.IsValidInput(number, unit));

        // The composition under test: number + space + unit label.
        WidthRequest request = ExportSize.ParseWidth(
            string.Concat(number, " ", PngPreview.UnitLabel(unit)));
        PixelDimensions expected = ExportSize.ToPixels(
            request,
            LandscapeAspect.WidthPt,
            LandscapeAspect.HeightPt);

        Assert.Equal(request, evaluation.Request);
        Assert.Equal(expected, evaluation.Dimensions);

        string preview = PngPreview.FormatPreview(evaluation.Request, evaluation.Dimensions);
        Assert.Contains(
            $"{expected.PixelWidth} \u00d7 {expected.PixelHeight} px",
            preview,
            StringComparison.Ordinal);
        Assert.Contains($"{PngPreview.UnitLabel(unit)} wide", preview, StringComparison.Ordinal);
    }

    /// <summary>An empty width asks for one, rather than showing the parser's empty-value message.</summary>
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void An_empty_width_shows_the_entry_hint(string input)
    {
        PngEvaluation evaluation = PngPreview.Evaluate(input, ExportUnit.Centimetres, LandscapeAspect);

        Assert.False(evaluation.Valid);
        Assert.Equal(PngPreview.EmptyInputMessage, evaluation.Message);
        Assert.False(PngPreview.IsValidInput(input, ExportUnit.Centimetres));
    }

    /// <summary>
    /// D3: the number box is a number and the dropdown is the unit, so a unit typed
    /// into the number box is refused with the parser's value message — never
    /// silently accepted as a second unit authority.
    /// </summary>
    [Fact]
    public void A_unit_typed_into_the_number_box_is_refused()
    {
        PngEvaluation evaluation = PngPreview.Evaluate("10cm", ExportUnit.Centimetres, LandscapeAspect);

        Assert.False(evaluation.Valid);
        Assert.Contains("not a valid width value", evaluation.Message, StringComparison.Ordinal);
        Assert.False(PngPreview.IsValidInput("10cm", ExportUnit.Centimetres));
    }

    /// <summary>A value that is not a number is refused with the parser's message.</summary>
    [Fact]
    public void A_malformed_value_is_refused_with_the_parser_message()
    {
        PngEvaluation evaluation = PngPreview.Evaluate("10.5.5", ExportUnit.Centimetres, LandscapeAspect);

        Assert.False(evaluation.Valid);
        Assert.Contains("not a valid width value", evaluation.Message, StringComparison.Ordinal);
    }

    /// <summary>A zero or negative width is refused, and the message names the rule.</summary>
    [Theory]
    [InlineData("0", ExportUnit.Centimetres)]
    [InlineData("-2", ExportUnit.Inches)]
    public void A_zero_or_negative_width_is_refused(string number, ExportUnit unit)
    {
        PngEvaluation evaluation = PngPreview.Evaluate(number, unit, LandscapeAspect);

        Assert.False(evaluation.Valid);
        Assert.Contains("greater than zero", evaluation.Message, StringComparison.Ordinal);
    }

    /// <summary>A non-finite value is refused, and the message says so.</summary>
    [Fact]
    public void A_non_finite_width_is_refused()
    {
        PngEvaluation evaluation = PngPreview.Evaluate("1e999", ExportUnit.Centimetres, LandscapeAspect);

        Assert.False(evaluation.Valid);
        Assert.Contains("not a finite number", evaluation.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// A width beyond the per-axis budget names the limit, so the user can pick a
    /// smaller width rather than guess why the preview refused.
    /// </summary>
    [Fact]
    public void A_width_over_the_per_axis_limit_names_the_limit()
    {
        PngEvaluation evaluation = PngPreview.Evaluate("100000", ExportUnit.Centimetres, LandscapeAspect);

        Assert.False(evaluation.Valid);
        Assert.Contains("65535", evaluation.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// A per-axis-legal pair whose combined count exceeds the allocation budget names
    /// the combined limit — a tall aspect makes a 10,000 px width produce 20,000 px.
    /// </summary>
    [Fact]
    public void A_width_over_the_combined_budget_names_the_limit()
    {
        ExportAspectPt tall = new(100, 200);

        PngEvaluation evaluation = PngPreview.Evaluate("10000", ExportUnit.Pixels, tall);

        Assert.False(evaluation.Valid);
        Assert.Contains("maximum total", evaluation.Message, StringComparison.Ordinal);
    }

    /// <summary>A sub-pixel width rounds to zero and is refused rather than exported as nothing.</summary>
    [Fact]
    public void A_sub_pixel_width_is_refused()
    {
        PngEvaluation evaluation = PngPreview.Evaluate("0.001", ExportUnit.Pixels, LandscapeAspect);

        Assert.False(evaluation.Valid);
        Assert.Contains("at least 1 px", evaluation.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// The preview text is culture-invariant: a comma-decimal locale must not turn
    /// 29.7 into 297.
    /// </summary>
    [Fact]
    public void The_preview_text_is_culture_invariant()
    {
        CultureInfo previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("de-DE");

            PngEvaluation evaluation = PngPreview.Evaluate("29.7", ExportUnit.Centimetres, LandscapeAspect);
            Assert.True(evaluation.Valid, evaluation.Message);

            string preview = PngPreview.FormatPreview(evaluation.Request, evaluation.Dimensions);

            Assert.Contains("29.7 cm wide", preview, StringComparison.Ordinal);
            Assert.DoesNotContain("297 cm", preview, StringComparison.Ordinal);
        }
        finally
        {
            CultureInfo.CurrentCulture = previous;
        }
    }

    /// <summary>The unit item ids round-trip for every unit and reject an unknown id.</summary>
    [Fact]
    public void The_unit_item_ids_round_trip()
    {
        foreach (ExportUnit unit in Enum.GetValues<ExportUnit>())
        {
            Assert.True(PngPreview.TryParseUnitId(PngPreview.UnitItemId(unit), out ExportUnit parsed));
            Assert.Equal(unit, parsed);
        }

        Assert.False(PngPreview.TryParseUnitId("bogus", out ExportUnit fallback));
        Assert.Equal(PngPreview.DefaultUnit, fallback);
        Assert.False(PngPreview.TryParseUnitId(null, out _));

        Assert.Equal(0, PngPreview.UnitIndex(ExportUnit.Centimetres));
        Assert.Equal(1, PngPreview.UnitIndex(ExportUnit.Inches));
        Assert.Equal(2, PngPreview.UnitIndex(ExportUnit.Pixels));
    }

    /// <summary>
    /// Every preset item id round-trips to its preset, so the dropdown can never offer
    /// an item the setter cannot resolve.
    /// </summary>
    [Fact]
    public void The_preset_item_ids_round_trip_for_every_preset()
    {
        foreach (SizePreset preset in SizePresets.All)
        {
            Assert.Same(preset, PngPreview.FindPreset(PngPreview.PresetItemId(preset)));
        }

        Assert.Null(PngPreview.FindPreset("bogus"));
        Assert.Null(PngPreview.FindPreset(RibbonControlIds.PngSizeItemPrefix));
    }

    /// <summary>
    /// The preset width text is the preset's own width in centimetres — no new
    /// literal — and it round-trips to the preset's points within the display
    /// precision (<see cref="PngPreview.CmText"/> keeps three decimals, i.e. at
    /// most 0.0005 cm ≈ 0.014 pt ≈ 0.017 px at 300 DPI).
    /// </summary>
    [Fact]
    public void The_preset_width_text_is_the_presets_width_in_centimetres()
    {
        foreach (SizePreset preset in SizePresets.All)
        {
            WidthRequest request = ExportSize.ParseWidth(PngPreview.CmText(preset.WidthPt) + " cm");

            Assert.Equal(ExportUnit.Centimetres, request.Unit);

            double roundTrippedPt = request.Value / ExportSize.CmPerInch * SizePresets.PointsPerInch;
            Assert.True(
                Math.Abs(preset.WidthPt - roundTrippedPt) < 0.05,
                $"preset {preset.Key}: {roundTrippedPt} pt is not within 0.05 pt of {preset.WidthPt} pt");
        }
    }

    /// <summary>The default width is the A4-landscape preset's width, not a literal.</summary>
    [Fact]
    public void The_default_width_is_the_a4_landscape_preset()
    {
        Assert.Equal(PngPreview.CmText(SizePresets.A4Landscape.WidthPt), PngPreview.DefaultWidthText);

        WidthRequest request = ExportSize.ParseWidth(PngPreview.DefaultWidthText + " cm");
        Assert.Equal(ExportUnit.Centimetres, request.Unit);
        Assert.True(
            Math.Abs(
                SizePresets.A4Landscape.WidthPt
                - (request.Value / ExportSize.CmPerInch * SizePresets.PointsPerInch)) < 0.05);
    }

    /// <summary>
    /// The preset label names the preset and its width in centimetres, so the
    /// dropdown reads as catalogue data rather than a bare key.
    /// </summary>
    [Fact]
    public void The_preset_label_names_the_preset_and_its_width_in_centimetres()
    {
        foreach (SizePreset preset in SizePresets.All)
        {
            string label = PngPreview.PresetLabel(preset);

            Assert.StartsWith(preset.DisplayName, label, StringComparison.Ordinal);
            Assert.Contains(PngPreview.CmText(preset.WidthPt) + " cm", label, StringComparison.Ordinal);
        }
    }

    /// <summary>
    /// The preset dropdown selects only the exact text a selection wrote: a custom
    /// width or a non-centimetre unit reads as no selection rather than a phantom row.
    /// </summary>
    [Fact]
    public void The_preset_index_matches_only_the_text_a_selection_wrote()
    {
        int landscapeIndex = SizePresets.All
            .ToList()
            .FindIndex(preset => ReferenceEquals(preset, SizePresets.A4Landscape));

        Assert.Equal(
            landscapeIndex,
            PngPreview.PresetIndex(PngPreview.CmText(SizePresets.A4Landscape.WidthPt), ExportUnit.Centimetres));
        Assert.Equal(-1, PngPreview.PresetIndex("10", ExportUnit.Centimetres));
        Assert.Equal(-1, PngPreview.PresetIndex(PngPreview.CmText(SizePresets.A4Landscape.WidthPt), ExportUnit.Inches));
        Assert.Equal(-1, PngPreview.PresetIndex(null, ExportUnit.Centimetres));
    }
}
