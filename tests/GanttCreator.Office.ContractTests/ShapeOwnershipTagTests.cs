using System.Globalization;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ShapeOwnershipTag"/>, the ADR-0019 carrier.
/// The tag is the R4.7 reconciliation partner and the R4.8 ownership filter, so
/// both its shape and its recogniser are pinned here.
/// </summary>
public class ShapeOwnershipTagTests
{
    [Fact]
    public void The_tag_is_the_prefix_plus_a_bounded_hex_hash()
    {
        string tag = ShapeOwnershipTag.ForPrimitiveId("row-1:bar");

        Assert.StartsWith(ShapeOwnershipTag.Prefix, tag, StringComparison.Ordinal);
        Assert.Equal(ShapeOwnershipTag.Prefix.Length + ShapeOwnershipTag.HashLength, tag.Length);
        Assert.Equal(tag, tag.ToUpperInvariant().Replace("GANTTCREATOR.OWNED.V1:", ShapeOwnershipTag.Prefix, StringComparison.Ordinal));
    }

    [Fact]
    public void The_same_identifier_always_produces_the_same_tag()
    {
        // Determinism is what makes the tag usable as a reconciliation key: the
        // same scene must reconcile against the same mark on every refresh.
        Assert.Equal(
            ShapeOwnershipTag.ForPrimitiveId("row-1:bar"),
            ShapeOwnershipTag.ForPrimitiveId("row-1:bar"));
    }

    [Fact]
    public void Different_identifiers_produce_different_tags()
    {
        Assert.NotEqual(
            ShapeOwnershipTag.ForPrimitiveId("row-1:bar"),
            ShapeOwnershipTag.ForPrimitiveId("row-1:bar2"));
    }

    [Fact]
    public void An_unbounded_shared_owner_identifier_still_yields_a_bounded_tag()
    {
        // ADR-0017: a shared owner's value is a pipe-joined row list, so its
        // PrimitiveId can be arbitrarily long. The tag must not grow with it, or
        // truncation would break the reconciliation key and the filter at once.
        string longId = string.Join('|', Enumerable.Range(1, 500).Select(i => "row-" + i.ToString(CultureInfo.InvariantCulture)));

        string tag = ShapeOwnershipTag.ForPrimitiveId(longId);

        Assert.True(longId.Length > 1000, "fixture should be genuinely long");
        Assert.Equal(ShapeOwnershipTag.Prefix.Length + ShapeOwnershipTag.HashLength, tag.Length);
    }

    [Fact]
    public void The_tag_does_not_depend_on_the_ambient_culture()
    {
        var original = CultureInfo.CurrentCulture;
        try
        {
            // A comma-decimal culture is the classic way a culture-sensitive
            // conversion would change a value; the tag is a string built from a
            // hex digest, so it must be byte-identical.
            CultureInfo.CurrentCulture = new CultureInfo("de-DE");
            string german = ShapeOwnershipTag.ForPrimitiveId("row-1:bar");
            CultureInfo.CurrentCulture = new CultureInfo("tr-TR");
            string turkish = ShapeOwnershipTag.ForPrimitiveId("row-1:bar");

            Assert.Equal(german, turkish);
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("a note the user typed")]
    [InlineData("GanttCreator")]
    public void Unowned_alternative_text_is_not_recognised_as_a_tag(string? alternativeText)
    {
        Assert.False(ShapeOwnershipTag.IsOwnedTag(alternativeText));
        Assert.False(ShapeOwnershipTag.HasOwnershipPrefix(alternativeText));
    }

    [Fact]
    public void A_prefix_with_a_damaged_hash_is_not_a_valid_tag_but_is_still_reported_as_malformed_owned()
    {
        // R9.4 needs the weaker test to tell a damaged owned shape from an
        // unowned one, so it can report the damage instead of ignoring it.
        string damaged = ShapeOwnershipTag.Prefix + "not-a-hash";

        Assert.False(ShapeOwnershipTag.IsOwnedTag(damaged));
        Assert.True(ShapeOwnershipTag.HasOwnershipPrefix(damaged));
    }

    [Fact]
    public void A_prefix_with_the_wrong_hash_length_is_not_a_valid_tag()
    {
        Assert.False(ShapeOwnershipTag.IsOwnedTag(ShapeOwnershipTag.Prefix));
        Assert.False(ShapeOwnershipTag.IsOwnedTag(ShapeOwnershipTag.Prefix + "ABC"));
    }

    [Fact]
    public void A_non_hex_hash_is_not_a_valid_tag()
    {
        Assert.False(ShapeOwnershipTag.IsOwnedTag(ShapeOwnershipTag.Prefix + new string('Z', ShapeOwnershipTag.HashLength)));
    }

    [Fact]
    public void A_well_formed_tag_is_recognised()
    {
        Assert.True(ShapeOwnershipTag.IsOwnedTag(ShapeOwnershipTag.ForPrimitiveId("row-1:bar")));
    }

    [Fact]
    public void A_blank_identifier_is_refused_rather_than_hashed()
    {
        // Positive test for the one argument guard: a blank identifier must fail
        // loudly, because a tag for "" would collide across every malformed call.
        Assert.Throws<ArgumentException>(() => ShapeOwnershipTag.ForPrimitiveId("   "));
        Assert.Throws<ArgumentNullException>(() => ShapeOwnershipTag.ForPrimitiveId(null!));
    }
}