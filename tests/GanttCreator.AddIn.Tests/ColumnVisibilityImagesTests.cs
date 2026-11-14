namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contracts for the column dropdown's two committed glyphs (R5.18h): both are embedded,
/// both decode, and they are not the same picture.
/// </summary>
/// <remarks>
/// The asset step is the part that breaks silently — a renamed file, a missed
/// <c>EmbeddedResource</c> glob or a corrupt PNG leaves the dropdown with no ticks and no
/// error anywhere — so it is asserted here rather than trusted to the .csproj.
/// </remarks>
public sealed class ColumnVisibilityImagesTests
{
    private static readonly byte[] PngSignature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

    /// <summary>Both committed sizes are embedded and are real PNGs.</summary>
    [Theory]
    [InlineData("column-visible-32.png")]
    [InlineData("column-hidden-32.png")]
    [InlineData("column-visible-16.png")]
    [InlineData("column-hidden-16.png")]
    public void Every_committed_glyph_is_embedded_and_is_a_png(string suffix)
    {
        byte[]? png = ColumnVisibilityImages.ReadPng(suffix);

        Assert.NotNull(png);
        Assert.True(png!.Length > PngSignature.Length);
        Assert.Equal(PngSignature, png[..PngSignature.Length]);
    }

    /// <summary>The tick and the dash are different pictures.</summary>
    [Fact]
    public void The_two_states_are_different_glyphs()
    {
        byte[] visible = ColumnVisibilityImages.ReadPng(ColumnVisibilityImages.VisibleResourceSuffix)!;
        byte[] hidden = ColumnVisibilityImages.ReadPng(ColumnVisibilityImages.HiddenResourceSuffix)!;

        Assert.NotEqual(visible.Length, hidden.Length);
    }

    /// <summary>An unknown suffix resolves to null rather than throwing.</summary>
    [Fact]
    public void An_unknown_suffix_resolves_to_null() =>
        Assert.Null(ColumnVisibilityImages.ReadPng("no-such-glyph.png"));

    /// <summary>
    /// Both states convert to the host's picture type. That conversion is the one step a
    /// headless test CAN reach — it needs no Excel, only the framework's OLE support.
    /// </summary>
    [Fact]
    public void Both_states_produce_a_picture()
    {
        Assert.NotNull(ColumnVisibilityImages.For(visible: true));
        Assert.NotNull(ColumnVisibilityImages.For(visible: false));
    }

    /// <summary>The dropdown's per-item image answers for every offered index only.</summary>
    [Fact]
    public void The_dropdown_answers_for_offered_indexes_only()
    {
        RibbonStateService service = new();

        Assert.NotNull(service.GetColumnVisibilityItemImage(0));
        Assert.NotNull(service.GetColumnVisibilityItemImage(4));
        Assert.Null(service.GetColumnVisibilityItemImage(5));
        Assert.Null(service.GetColumnVisibilityItemImage(-1));
    }
}
