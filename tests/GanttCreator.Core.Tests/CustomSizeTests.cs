namespace GanttCreator.Core.Tests;

/// <summary>
/// Tests for the custom centimetre **width** (R8.9, ADR-0072): the band, the exact
/// points conversion, the runtime Custom preset, and the clamp used at every read.
/// Every refusal arm has a positive test.
/// </summary>
/// <remarks>
/// The custom height was withdrawn on 2026-11-26: the live height is the measured rows,
/// so there is no height to validate, convert, or place. Anything here that mentions a
/// preset height exists only because <see cref="SizePreset"/> requires a positive extent.
/// </remarks>
public sealed class CustomSizeTests
{
    private static double PtPerCm => 10.0 * SizePresets.PointsPerMillimetre;

    [Fact]
    public void A_valid_width_converts_once()
    {
        Assert.True(CustomSizeCm.TryCreate(29.7, out CustomSizeCm? size, out CustomSizeRefusal? refusal));

        Assert.Null(refusal);
        Assert.Equal(29.7, size!.WidthCm);
        Assert.Equal(29.7 * PtPerCm, size.WidthPt, 9);
    }

    [Theory]
    [InlineData(0.0, CustomSizeRefusal.InvalidWidth)]
    [InlineData(-1.0, CustomSizeRefusal.InvalidWidth)]
    [InlineData(double.NaN, CustomSizeRefusal.InvalidWidth)]
    [InlineData(double.PositiveInfinity, CustomSizeRefusal.InvalidWidth)]
    [InlineData(0.5, CustomSizeRefusal.WidthOutOfBand)]
    [InlineData(200.1, CustomSizeRefusal.WidthOutOfBand)]
    public void Every_refusal_arm_is_positive(double widthCm, CustomSizeRefusal expected)
    {
        Assert.False(CustomSizeCm.TryCreate(widthCm, out CustomSizeCm? size, out CustomSizeRefusal? refusal));
        Assert.Null(size);
        Assert.Equal(expected, refusal);
    }

    [Fact]
    public void The_band_endpoints_are_accepted()
    {
        Assert.True(CustomSizeCm.TryCreate(CustomSizeCm.MinimumCm, out _, out _));
        Assert.True(CustomSizeCm.TryCreate(CustomSizeCm.MaximumCm, out _, out _));
    }

    [Theory]
    [InlineData(0.5)]
    [InlineData(200.1)]
    [InlineData(0.0)]
    [InlineData(double.NaN)]
    public void A_read_width_is_clamped_into_the_band_rather_than_refused(double stored)
    {
        // A hand-edited workbook can hold anything. Clamping at the read keeps the mode
        // alive; refusing here would collapse Custom and leave the ribbon disagreeing
        // with the stored SizePreset, which is the defect this clamp exists to prevent.
        CustomSizeCm size = CustomSizeCm.FromClampedWidth(stored);

        Assert.InRange(size.WidthCm, CustomSizeCm.MinimumCm, CustomSizeCm.MaximumCm);
    }

    [Fact]
    public void An_in_band_read_width_is_kept_exactly() =>
        Assert.Equal(42.0, CustomSizeCm.FromClampedWidth(42.0).WidthCm, 9);

    [Fact]
    public void The_custom_preset_carries_the_width_and_the_custom_key()
    {
        Assert.True(CustomSizeCm.TryCreate(42.0, out CustomSizeCm? size, out _));
        CustomSizeCm custom = size!;

        SizePreset preset = SizePresets.Custom(custom);

        Assert.Equal(SizePresets.CustomKey, preset.Key);
        Assert.Equal("Custom", preset.DisplayName);
        Assert.Equal(custom.WidthPt, preset.WidthPt, 9);

        // Not a page dimension: nothing consumes the preset height, and it is the width on
        // both axes so no stored centimetre can be mistaken for a second authority.
        Assert.Equal(custom.WidthPt, preset.HeightPt, 9);
        Assert.True(SizePresets.IsCustom(preset));
    }

    [Fact]
    public void The_fixed_catalogue_stays_fixed_and_does_not_resolve_custom_by_key()
    {
        Assert.Equal(4, SizePresets.All.Count);
        Assert.DoesNotContain(SizePresets.All, preset => SizePresets.IsCustom(preset));
        Assert.Null(SizePresets.ByKey(SizePresets.CustomKey));
        Assert.False(SizePresets.IsCustom(SizePresets.A4Portrait));
        Assert.False(SizePresets.IsCustom(null));
    }
}
