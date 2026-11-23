namespace GanttCreator.Raster;

/// <summary>
/// One pinned font face identity: a face the scene's committed advance table
/// was generated from, and the exact face the raster renderer must draw with
/// (ADR-0070 D2).
/// </summary>
/// <param name="Family">The face's family name (<c>Aptos</c> or <c>Aptos Narrow</c>).</param>
/// <param name="PostScriptName">The PostScript name read from the face's <c>name</c> table.</param>
/// <param name="FileName">The file name inside the Office cloud-font family folder.</param>
/// <param name="NameTableVersion">The <c>name</c> table's version string (name ID 5).</param>
/// <param name="HeadRevision">The <c>head</c> table's font revision.</param>
/// <param name="UnitsPerEm">The face's design-unit denominator; 2048 for every pinned face.</param>
/// <param name="Sha256">The upper-case SHA-256 of the face file: the byte-identity pin.</param>
/// <param name="IsBold">Whether the face is the bold member of its family.</param>
public sealed record PinnedFontFace(
    string Family,
    string PostScriptName,
    string FileName,
    string NameTableVersion,
    double HeadRevision,
    int UnitsPerEm,
    string Sha256,
    bool IsBold)
{
    /// <summary>
    /// Compares this pin with an observed face identity and reports the first
    /// mismatch, or <see langword="null"/> when the observation is exactly the
    /// pinned face.
    /// </summary>
    /// <param name="observed">The identity observed from a loaded typeface.</param>
    /// <returns>The typed refusal, or <see langword="null"/> on an exact match.</returns>
    /// <remarks>
    /// This is the fail-closed identity check ADR-0070 D3 requires before the
    /// first text draw. Every member that identifies the face bytes is
    /// compared, and an observation whose <c>name</c>/<c>head</c> tables could
    /// not be read is refused as <see cref="PinnedFontRefusal.ParseFailure"/>
    /// rather than treated as a match: an unverifiable face is never the pin.
    /// </remarks>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="observed"/> is null.</exception>
    public PinnedFontRefusal? Match(PinnedFontObservation observed)
    {
        ArgumentNullException.ThrowIfNull(observed);

        return string.IsNullOrEmpty(observed.NameTableVersion) || double.IsNaN(observed.HeadRevision)
            ? PinnedFontRefusal.ParseFailure
            : !string.Equals(Family, observed.Family, StringComparison.Ordinal)
                ? PinnedFontRefusal.FamilyMismatch
                : !string.Equals(PostScriptName, observed.PostScriptName, StringComparison.Ordinal)
                    ? PinnedFontRefusal.PostScriptMismatch
                    : !string.Equals(NameTableVersion, observed.NameTableVersion, StringComparison.Ordinal)
                        ? PinnedFontRefusal.VersionMismatch
                        : Math.Abs(HeadRevision - observed.HeadRevision) > 0.00005
                            ? PinnedFontRefusal.VersionMismatch
                            : UnitsPerEm != observed.UnitsPerEm
                                ? PinnedFontRefusal.UnitsPerEmMismatch
                                : string.Equals(Sha256, observed.Sha256, StringComparison.Ordinal)
                                    ? null
                                    : PinnedFontRefusal.VersionMismatch;
    }
}

/// <summary>
/// The single pinned-face authority (ADR-0070 D5): the four Office-installed
/// Aptos faces the scene's advance table was generated from. The renderer and
/// the font-presence probe both consume this set, so a probe that disagrees
/// with the renderer is structurally impossible.
/// </summary>
/// <remarks>
/// The identities are transcribed from ADR-0070 D2's table. Changing a face,
/// its version, or its hash is a measured change: the pin, the advance table,
/// and the affected golden baselines move together with human review — never a
/// silent font swap.
/// </remarks>
public static class PinnedFontSet
{
    /// <summary>
    /// Gets the four pinned faces: Aptos Regular/Bold and Aptos Narrow
    /// Regular/Bold, exactly as recorded in ADR-0070 D2.
    /// </summary>
    public static IReadOnlyList<PinnedFontFace> Faces { get; } =
    [
        new(
            Family: "Aptos",
            PostScriptName: "Aptos",
            FileName: "30153066857.ttf",
            NameTableVersion: "Version 2.01;O365",
            HeadRevision: 2.0100,
            UnitsPerEm: 2048,
            Sha256: "95980114FCFD42F2F9C446DAE429B70582BF2F03097D68433EA9E7D85A49DA0B",
            IsBold: false),
        new(
            Family: "Aptos",
            PostScriptName: "Aptos-Bold",
            FileName: "31531277363.ttf",
            NameTableVersion: "Version 2.00;O365",
            HeadRevision: 2.0000,
            UnitsPerEm: 2048,
            Sha256: "B4124A6C17E73189574D9E51ADDA221F236D1DF901F1BBCD57E9A09AE5B18599",
            IsBold: true),
        new(
            Family: "Aptos Narrow",
            PostScriptName: "Aptos-Narrow",
            FileName: "27970306811.ttf",
            NameTableVersion: "Version 2.01;O365",
            HeadRevision: 2.0100,
            UnitsPerEm: 2048,
            Sha256: "DED7515AA064578485A09C11A798A05D0303BDACDD7F2DE92D0521F3BC7793BA",
            IsBold: false),
        new(
            Family: "Aptos Narrow",
            PostScriptName: "AptosNarrow-2bc0000-560000",
            FileName: "24328775610.ttf",
            NameTableVersion: "Version 2.00;O365",
            HeadRevision: 2.0000,
            UnitsPerEm: 2048,
            Sha256: "D1E33B7FFA6C481063DF87516D0C2976DCE00C717663C7943C2742E3E4BCAA99",
            IsBold: true),
    ];

    /// <summary>
    /// Finds the pinned face for a scene style's family name and weight.
    /// </summary>
    /// <param name="fontFamily">The resolved family name carried by the scene style.</param>
    /// <param name="bold">Whether the resolved style is bold.</param>
    /// <returns>The pinned face, or <see langword="null"/> when the family is not pinned.</returns>
    /// <remarks>
    /// A null return is deliberately not exception-throwing: the resolver maps
    /// an unpinned family to the typed <see cref="PinnedFontRefusal.FamilyMismatch"/>
    /// refusal, which the renderer surfaces before any text is drawn.
    /// </remarks>
    public static PinnedFontFace? FindFace(string fontFamily, bool bold)
    {
        foreach (PinnedFontFace face in Faces)
        {
            if (face.IsBold == bold && string.Equals(face.Family, fontFamily, StringComparison.Ordinal))
            {
                return face;
            }
        }

        return null;
    }
}
