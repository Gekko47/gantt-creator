using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace GanttCreator.Office;

/// <summary>
/// Builds and recognises the ownership tag the live renderer writes to every
/// shape it creates (ADR-0019 D1/D2).
/// </summary>
/// <remarks>
/// <para>
/// The tag is <c>GanttCreator.Owned.v1:{hash}</c>, written to the shape's
/// <c>AlternativeText</c> member. It is deliberately <em>not</em> the full
/// scene primitive identifier: under ADR-0017 a shared owner's
/// <c>SceneOwnerId.Value</c> is a <c>|</c>-joined list of contributing row
/// IDs, so the identifier is unbounded, and a truncated tag would break the
/// R4.7 reconciliation key and the R4.8 ownership filter at the same time.
/// The unbounded identifier travels in the shape's <c>Name</c> instead.
/// </para>
/// <para>
/// The <c>v1</c> segment lets R9.4 tell a tag written by this build apart
/// from user content or a foreign add-in without guessing. The prefix is the
/// filter; the hash is the identity.
/// </para>
/// <para>
/// The tag is <strong>not</strong> a workbook schema value: nothing in
/// <c>tblGanttSettings</c> or on <c>_GanttCreatorConfig</c> stores it, so
/// narrowing the hash width is not a schema migration. R3.16's
/// <c>EquivalenceFields</c> forbids a <c>GanttCreator.</c> string in Core
/// because the tag is composed by the renderer; this type therefore lives in
/// the Office project, never in <c>GanttCreator.Core</c>.
/// </para>
/// </remarks>
public static class ShapeOwnershipTag
{
    /// <summary>
    /// The exact prefix every owned shape's alternative text starts with. R4.8's
    /// ownership filter is a test for this prefix, never a substring search, so a
    /// user-authored alternative text that merely mentions Gantt Creator is never
    /// classified as owned.
    /// </summary>
    public const string Prefix = "GanttCreator.Owned.v1:";

    /// <summary>
    /// The number of hex characters of the digest included in the tag.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Sixteen characters is 64 bits, and the value is deliberate rather than
    /// incidental — but <strong>not for the reason a collision calculation
    /// would give.</strong> An earlier version of this comment justified the
    /// width with a birthday bound: "a 1,000-shape chart has a collision
    /// probability near 2.7 x 10^-14". That reasoning is wrong, and it was worth
    /// removing rather than restating, because it invites the reader to shorten
    /// the hash by a factor of four on a chart-size estimate.
    /// </para>
    /// <para>
    /// <strong>Collisions are inert here.</strong> The tag is never used as a
    /// lookup key. A shape is found by its <c>Name</c>, and ownership is then
    /// proved by <em>recomputing</em> the expected tag for that name and
    /// comparing — <c>CarriesOwnershipTagFor(shape, name)</c> in
    /// <see cref="ExcelShapeWriter"/>, and the same per-shape check in
    /// <c>ListOwned</c>. So if two owned shapes happened to hash alike, each is
    /// still verified against its own identifier and both remain correctly
    /// classified. No shape count makes the width insufficient, and no shape
    /// count would make a shorter width unsafe.
    /// </para>
    /// <para>
    /// <strong>What the width is actually for.</strong> The tag separates
    /// add-in-generated content from a user's own, so that a refresh never
    /// reorders, rewrites, or deletes a shape the user drew (R4.8). The
    /// <see cref="Prefix"/> does that filtering; the hash pins the tag to one
    /// specific identifier, so a shape carrying some <em>other</em> entity's
    /// valid tag is correctly treated as unowned. The width is sized generously
    /// because the whole tag is 38 characters and round-trips unaltered on the
    /// live host (verified by the R4.1 probe, which also confirmed Excel's name
    /// cap is unrelated — that limit is on <c>Name</c>, see KNOWN-LIMITATIONS
    /// L18). Generosity here is free, so there is no reason to economise.
    /// </para>
    /// <para>
    /// The <c>v1</c> segment, not this width, is the lever if the scheme ever
    /// changes: a future revision can widen or re-key the tag and R9.4 can then
    /// tell this build's tags from any other. Narrowing the hash is therefore a
    /// versioned change, not a tuning one, and should not be done by editing
    /// this constant alone.
    /// </para>
    /// </remarks>
    internal const int HashLength = 16;

    private const int _hashBytesToKeep = HashLength / 2;

    /// <summary>
    /// Builds the ownership tag for one scene primitive identifier.
    /// </summary>
    /// <param name="primitiveId">The stable role-derived scene primitive identifier.</param>
    /// <returns>The bounded <c>GanttCreator.Owned.v1:{hash}</c> tag.</returns>
    /// <exception cref="ArgumentException">
    /// <paramref name="primitiveId"/> is <see langword="null"/>, empty, or
    /// whitespace. <c>ScenePrimitive.CreateId</c> already refuses a blank role,
    /// so a blank identifier reaching here is a caller defect, not a data condition.
    /// </exception>
    public static string ForPrimitiveId(string primitiveId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(primitiveId);

        // Ordinal, never the ambient culture: the same identifier must produce
        // the same tag on every host, or the reconciliation key would be
        // machine-dependent.
        var utf8 = Encoding.UTF8.GetBytes(primitiveId);
        var digest = SHA256.HashData(utf8);

        return string.Create(
            CultureInfo.InvariantCulture,
            $"{Prefix}{Convert.ToHexString(digest.AsSpan(0, _hashBytesToKeep))}");
    }

    /// <summary>
    /// Determines whether an alternative-text value is a valid add-in ownership tag.
    /// </summary>
    /// <param name="alternativeText">
    /// The shape's alternative text, or <see langword="null"/> when the shape has none.
    /// </param>
    /// <returns>
    /// <see langword="true"/> when the value starts with <see cref="Prefix"/> and
    /// carries a well-formed hash after it; otherwise <see langword="false"/>.
    /// </returns>
    /// <remarks>
    /// A value that starts with the prefix but whose remainder is not exactly
    /// <see cref="HashLength"/> hex characters is reported as <em>not</em> owned.
    /// That is deliberate: a malformed tag must fall to the unowned path so R4.8
    /// never deletes it and R9.4 can report it, rather than being silently
    /// accepted as a valid mark.
    /// </remarks>
    public static bool IsOwnedTag(string? alternativeText)
    {
        if (alternativeText is null || !HasOwnershipPrefix(alternativeText))
        {
            return false;
        }

        ReadOnlySpan<char> hash = alternativeText.AsSpan(Prefix.Length);
        if (hash.Length != HashLength)
        {
            return false;
        }

        foreach (var character in hash)
        {
            var isHexDigit =
                character is (>= '0' and <= '9')
                or (>= 'A' and <= 'F')
                or (>= 'a' and <= 'f');
            if (!isHexDigit)
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Determines whether an alternative-text value carries the ownership prefix
    /// without validating the hash.
    /// </summary>
    /// <param name="alternativeText">The shape's alternative text, or <see langword="null"/>.</param>
    /// <returns><see langword="true"/> when the value starts with <see cref="Prefix"/>.</returns>
    /// <remarks>
    /// R9.4's repair classification needs this weaker test to tell a
    /// <em>malformed owned</em> shape (prefix present, hash damaged) apart from a
    /// wholly unowned one, so the repair row can report the damage instead of
    /// ignoring the shape entirely.
    /// </remarks>
    public static bool HasOwnershipPrefix(string? alternativeText) =>
        alternativeText is not null && alternativeText.StartsWith(Prefix, StringComparison.Ordinal);
}
