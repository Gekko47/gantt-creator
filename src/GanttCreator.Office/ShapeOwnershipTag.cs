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
    /// The number of hex characters of the digest included in the tag. Sixteen
    /// characters is 64 bits: wide enough that a 1,000-shape chart has a
    /// collision probability near 2.7 x 10^-14, and short enough that the whole
    /// tag stays well inside Excel's alternative-text capacity.
    /// </summary>
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
