namespace GanttCreator.AddIn;

/// <summary>
/// The picture each lock toggle shows, so the glyph matches the state (D1, owner ruling
/// 2026-11-14).
/// </summary>
/// <remarks>
/// <para>
/// <b>This closes the half of R5.18f that measurement had blocked.</b> The approved design was
/// a closed padlock while locked and an open padlock while unlocked, and it could not ship
/// because no open-padlock <c>imageMso</c> exists: the only padlock ids present in both local
/// Office sources are <c>Lock</c> and <c>LockCell</c>. The label carried the state instead.
/// The owner then ruled that Lucide is available for every ribbon icon (D7), which removes the
/// constraint — a <c>getImage</c> callback returns the same <c>IPictureDisp</c> the column
/// dropdown already produces, so the open padlock is an asset rather than a gallery id.
/// </para>
/// <para>
/// <b>The glyphs are committed PNGs</b>, rasterised from Lucide's <c>lock</c> and
/// <c>lock-open</c> geometry by <c>scripts/gen-column-visibility-icons.ps1</c>, and decoded
/// through <see cref="RibbonImages"/> like every other ribbon picture. Nothing renders a path
/// at run time.
/// </para>
/// <para>
/// <b>The label stays.</b> The image and the label state the same fact deliberately: a 16 px
/// glyph is ambiguous at a glance, invisible to a screen reader, and one of the two padlocks
/// is a small difference at 100 % scaling, so the word is what makes the state unambiguous.
/// </para>
/// </remarks>
internal static class LockGlyphImages
{
    /// <summary>The embedded suffix of the closed padlock, shown while the lock is applied.</summary>
    internal const string LockedResourceSuffix = "lock-locked-32.png";

    /// <summary>The embedded suffix of the open padlock, shown while the lock is released.</summary>
    internal const string UnlockedResourceSuffix = "lock-unlocked-32.png";

    /// <summary>
    /// Returns the picture for a lock state.
    /// </summary>
    /// <param name="locked">Whether the control's state is "locked".</param>
    /// <returns>The <c>IPictureDisp</c> wrapper, or null when the glyph is unavailable.</returns>
    internal static object? For(bool locked) =>
        RibbonImages.For(locked ? LockedResourceSuffix : UnlockedResourceSuffix);
}
