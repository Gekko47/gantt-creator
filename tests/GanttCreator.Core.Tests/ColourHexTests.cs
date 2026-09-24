using System.Globalization;

namespace GanttCreator.Core.Tests;

public sealed class ColourHexTests
{
    [Fact]
    public void Parse_reads_six_digit_colour_with_implicit_alpha()
    {
        ColourHex colour = ColourHex.Parse("#0F0F0F");
        Assert.Equal(0xFF0F0F0Fu, colour.ARGB);
        Assert.Equal("#0F0F0F", colour.ToString());
    }

    [Fact]
    public void Parse_reads_explicit_alpha()
    {
        ColourHex colour = ColourHex.Parse("#0F0F0F80");
        Assert.Equal(0x800F0F0Fu, colour.ARGB);
        Assert.Equal("#0F0F0F80", colour.ToString());
    }

    [Fact]
    public void Six_and_eight_digit_forms_compare_by_argb()
    {
        ColourHex implicitAlpha = ColourHex.Parse("#0F0F0F");
        ColourHex explicitAlpha = ColourHex.Parse("#0F0F0FFF");
        Assert.Equal(implicitAlpha, explicitAlpha);
        Assert.True(implicitAlpha == explicitAlpha);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("0F0F0F")]
    [InlineData("#0F0F")]
    [InlineData("#0F0F0F0")]
    [InlineData("#0F0F0F0G")]
    [InlineData("#0F0G0F")]
    [InlineData("1e999")]
    [InlineData("#0f0f0f")]
    public void TryParse_rejects_malformed_input_without_throwing(string? text)
    {
        Assert.False(ColourHex.TryParse(text, out ColourHex? colour));
        Assert.Null(colour);
    }

    [Fact]
    public void TryParse_trims_surrounding_whitespace()
    {
        Assert.True(ColourHex.TryParse("  #0F0F0F\t", out ColourHex? colour));
        Assert.Equal("#0F0F0F", colour!.ToString());
    }

    [Fact]
    public void Parse_throws_for_null_and_malformed_input()
    {
        Assert.Throws<ArgumentNullException>(() => ColourHex.Parse(null));
        Assert.Throws<FormatException>(() => ColourHex.Parse("not-a-colour"));
    }

    [Fact]
    public void Colour_is_culture_invariant()
    {
        CultureInfo original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
            ColourHex colour = ColourHex.Parse("#0F0F0F");
            Assert.Equal("#0F0F0F", colour.ToString());
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }
}
