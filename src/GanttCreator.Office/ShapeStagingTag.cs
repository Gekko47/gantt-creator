using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace GanttCreator.Office;

/// <summary>
/// Builds and recognises the staging ownership tag the editable-export composition
/// writes to every temporary shape (R6.2, work-item D4).
/// </summary>
/// <remarks>
/// <para>
/// The tag is <c>GanttCreator.Staging.v1:{hash}</c>, written to the shape's
/// <c>AlternativeText</c> member, exactly as <see cref="ShapeOwnershipTag"/> writes
/// the live chart's tag. The staged shape's <c>Name</c> is
/// <see cref="NamePrefix"/> followed by the scene primitive identifier.
/// </para>
/// <para>
/// <b>Why a distinct scheme at all.</b> Staged shapes land on the same worksheet as
/// the live chart. Reusing the live name and tag would make a staged shape
/// indistinguishable from a live one - Excel permits duplicate names, and the
/// reconciler would then see ambiguous duplicates and could delete a staged shape
/// during a coalesced refresh (work-item D4). The prefix pair makes each family's
/// ownership scan structurally blind to the other's shapes.
/// </para>
/// <para>
/// <b>The invariant is preserved: <c>tag = f(name)</c>.</b> The digest is computed
/// over the full staged name (prefix included), so proof is always a recomputation
/// from the live shape name - never a stored value trusted on its own. That is the
/// same rule <c>CarriesOwnershipTagFor</c> applies to live shapes.
/// </para>
/// <para>
/// The tag is <strong>not</strong> a workbook schema value: nothing in
/// <c>tblGanttSettings</c> or on <c>_GanttCreatorConfig</c> stores it, so no schema
/// version or catalogue hash moves with this type.
/// </para>
/// </remarks>
public static class ShapeStagingTag
{
    /// <summary>
    /// The exact prefix every staged shape's <c>Name</c> starts with, followed by the
    /// scene primitive identifier. Deliberately not a form
    /// <see cref="ShapeOwnershipTag.ForPrimitiveId"/> would ever produce.
    /// </summary>
    public const string NamePrefix = "GanttCreatorStaging:";

    /// <summary>
    /// The exact prefix every staged shape's alternative text starts with. It is a
    /// different prefix from <see cref="ShapeOwnershipTag.Prefix"/>, so neither
    /// family's ownership filter can classify the other's shapes as its own.
    /// </summary>
    public const string Prefix = "GanttCreator.Staging.v1:";

    /// <summary>
    /// The number of hex characters of the digest included in the tag. Shared with
    /// <see cref="ShapeOwnershipTag"/> rather than restated, so the two schemes cannot
    /// drift apart in width while both round-trip on the same host.
    /// </summary>
    internal const int HashLength = ShapeOwnershipTag.HashLength;

    private const int _hashBytesToKeep = HashLength / 2;

    /// <summary>
    /// Builds the staging tag for one staged shape name.
    /// </summary>
    /// <param name="stagedName">
    /// The full staged shape name: <see cref="NamePrefix"/> joined with the scene
    /// primitive identifier.
    /// </param>
    /// <returns>The bounded <c>GanttCreator.Staging.v1:{hash}</c> tag.</returns>
    /// <exception cref="ArgumentException">
    /// <paramref name="stagedName"/> is <see langword="null"/>, empty, or whitespace.
    /// </exception>
    public static string ForStagedName(string stagedName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(stagedName);

        // Ordinal, never the ambient culture: the same name must produce the same tag
        // on every host, or the ownership proof would be machine-dependent.
        byte[] utf8 = Encoding.UTF8.GetBytes(stagedName);
        byte[] digest = SHA256.HashData(utf8);

        return string.Create(
            CultureInfo.InvariantCulture,
            $"{Prefix}{Convert.ToHexString(digest.AsSpan(0, _hashBytesToKeep))}");
    }

    /// <summary>
    /// Determines whether an alternative-text value is a valid staging tag.
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
    /// <see cref="HashLength"/> hex characters is reported as <em>not</em> staging-owned,
    /// matching <see cref="ShapeOwnershipTag.IsOwnedTag"/>'s fail-to-unowned rule.
    /// </remarks>
    public static bool IsStagingTag(string? alternativeText)
    {
        if (alternativeText is null || !alternativeText.StartsWith(Prefix, StringComparison.Ordinal))
        {
            return false;
        }

        ReadOnlySpan<char> hash = alternativeText.AsSpan(Prefix.Length);
        if (hash.Length != HashLength)
        {
            return false;
        }

        foreach (char character in hash)
        {
            bool isHexDigit =
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
}
