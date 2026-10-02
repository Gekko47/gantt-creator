namespace GanttCreator.Core.Tests;

/// <summary>
/// R4.7H: the presets carry <b>exact</b> dimensions and a declared orientation, and the
/// catalogue is internally consistent.
/// </summary>
public sealed class SizePresetTests
{
    /// <summary>The tolerance the exactness assertions use.</summary>
    private const double TolerancePt = 1e-9;

    [Fact]
    public void The_declared_catalogue_is_internally_consistent()
    {
        IReadOnlyList<string> failures = SizePresets.ValidateCatalogue();

        Assert.Empty(failures);
    }

    [Fact]
    public void A4_portrait_is_exactly_210_by_297_millimetres()
    {
        SizePreset a4 = SizePresets.A4Portrait;

        Assert.Equal(210.0 * SizePresets.PointsPerMillimetre, a4.WidthPt, TolerancePt);
        Assert.Equal(297.0 * SizePresets.PointsPerMillimetre, a4.HeightPt, TolerancePt);
        Assert.Equal(SizeOrientation.Portrait, a4.Orientation);
    }

    [Fact]
    public void A4_landscape_is_the_transpose_of_portrait()
    {
        // Transposing rather than re-declaring is the point: a landscape A4 that
        // was not the exact transpose would print on the wrong page orientation.
        Assert.Equal(SizePresets.A4Portrait.HeightPt, SizePresets.A4Landscape.WidthPt, TolerancePt);
        Assert.Equal(SizePresets.A4Portrait.WidthPt, SizePresets.A4Landscape.HeightPt, TolerancePt);
        Assert.Equal(SizeOrientation.Landscape, SizePresets.A4Landscape.Orientation);
    }

    [Fact]
    public void The_presentation_ratio_16_by_9_is_exact_and_not_approximate()
    {
        // The acceptance test requires exactness, not "close to". A slide that is
        // off by a fraction letterboxes on a projector.
        Assert.Equal(16.0 / 9.0, SizePresets.Presentation16x9.AspectRatio, TolerancePt);
        Assert.Equal(1280.0, SizePresets.Presentation16x9.WidthPt, TolerancePt);
        Assert.Equal(720.0, SizePresets.Presentation16x9.HeightPt, TolerancePt);
    }

    [Fact]
    public void The_presentation_ratio_4_by_3_is_exact_and_not_approximate()
    {
        Assert.Equal(4.0 / 3.0, SizePresets.Presentation4x3.AspectRatio, TolerancePt);
        Assert.Equal(1024.0, SizePresets.Presentation4x3.WidthPt, TolerancePt);
        Assert.Equal(768.0, SizePresets.Presentation4x3.HeightPt, TolerancePt);
    }

    [Fact]
    public void The_aspect_ratio_is_derived_from_the_dimensions_and_cannot_drift()
    {
        // Derived, not stored: a hand-maintained ratio field would eventually
        // disagree with its own dimensions and nobody would notice until print.
        SizePreset preset = SizePresets.Presentation16x9;

        Assert.Equal(preset.WidthPt / preset.HeightPt, preset.AspectRatio, TolerancePt);
    }

    [Fact]
    public void Every_preset_is_oriented_consistently_with_its_own_dimensions()
    {
        foreach (SizePreset preset in SizePresets.All)
        {
            var wider = preset.WidthPt > preset.HeightPt;
            var expectedLandscape = wider;
            Assert.True(
                expectedLandscape == (preset.Orientation == SizeOrientation.Landscape),
                $"{preset.Key} declares {preset.Orientation} but is {(wider ? "wider" : "taller")} than the other side.");
        }
    }

    [Fact]
    public void The_catalogue_holds_exactly_the_four_preset_keys()
    {
        // Three NAMED sizes produce four KEYS: A4 contributes two keys, one per
        // orientation (A4Portrait, A4Landscape), because a stored setting names a
        // specific page orientation rather than a paper size. The count is therefore 4
        // while the number of named sizes is 3, and both numbers are correct -- the test
        // name and its comment previously said "three", which contradicted the
        // assertion directly beneath them.
        Assert.Equal(4, SizePresets.All.Count);
        Assert.Equal(
            ["A4Landscape", "A4Portrait", "Presentation16x9", "Presentation4x3"],
            SizePresets.All.Select(preset => preset.Key).Order(StringComparer.Ordinal));
    }

    [Fact]
    public void Preset_keys_are_unique_so_a_stored_setting_names_exactly_one_preset()
    {
        Assert.Equal(
            SizePresets.All.Count,
            SizePresets.All.Select(preset => preset.Key).Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void A_stored_key_resolves_to_its_preset()
    {
        Assert.Same(SizePresets.Presentation16x9, SizePresets.ByKey("Presentation16x9"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("Slide16x9")]
    [InlineData("presentation16x9")]
    public void An_absent_or_unknown_key_resolves_to_nothing_rather_than_a_default(string? key)
    {
        // Null, not the default preset: silently substituting A4 for an unreadable
        // stored setting would render a chart the user never asked for. The caller
        // decides whether to fall back and can report that it did.
        Assert.Null(SizePresets.ByKey(key));
    }

    [Fact]
    public void The_key_lookup_is_case_sensitive_because_the_key_is_stored()
    {
        // The reverse case is already asserted above; this pins that the lookup does
        // not normalise, so two spellings can never silently share one preset.
        Assert.Null(SizePresets.ByKey("PRESENTATION16X9"));
    }

    [Fact]
    public void The_default_preset_is_a4_landscape()
    {
        Assert.Same(SizePresets.A4Landscape, SizePresets.Default);
    }
}
