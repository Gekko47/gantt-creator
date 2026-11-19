namespace GanttCreator.Office;

/// <summary>
/// The clipboard state observed at one moment: its sequence number and the complete format
/// name set (R6.6).
/// </summary>
/// <param name="SequenceNumber">
/// <c>GetClipboardSequenceNumber</c> at the read. A change between two reads means the
/// clipboard content was replaced, not merely that a format stopped rendering.
/// </param>
/// <param name="FormatNames">Every format the clipboard reports, in enumeration order.</param>
public sealed record ClipboardFormatSnapshot(uint SequenceNumber, IReadOnlyList<string> FormatNames)
{
    /// <summary>Whether a format with the given name is present, ordinal case-insensitive.</summary>
    /// <param name="name">The format name to look for.</param>
    /// <returns><see langword="true"/> when present.</returns>
    public bool Contains(string name) =>
        FormatNames.Any(format => string.Equals(format, name, StringComparison.OrdinalIgnoreCase));
}

/// <summary>
/// The outcome of one clipboard read: a snapshot or a typed failure reason (R6.6).
/// </summary>
/// <param name="Succeeded">Whether the clipboard could be opened and enumerated.</param>
/// <param name="Snapshot">The observed state when <paramref name="Succeeded"/> is true.</param>
/// <param name="Failure">
/// A stable diagnostic token when <paramref name="Succeeded"/> is false. A busy or unavailable
/// clipboard is a distinct typed outcome, never a hang and never an exception crossing the
/// stage.
/// </param>
public sealed record ClipboardFormatRead(
    bool Succeeded,
    ClipboardFormatSnapshot? Snapshot,
    string? Failure)
{
    /// <summary>Builds a successful result.</summary>
    /// <param name="snapshot">The observed clipboard state.</param>
    /// <returns>The result.</returns>
    public static ClipboardFormatRead ForSuccess(ClipboardFormatSnapshot snapshot) =>
        new(true, snapshot, null);

    /// <summary>Builds a failed result.</summary>
    /// <param name="failure">The diagnostic token.</param>
    /// <returns>The result.</returns>
    public static ClipboardFormatRead ForFailure(string failure) =>
        new(false, null, failure);
}

/// <summary>
/// Narrow port over the one clipboard observation the editable export needs: enumerate the
/// format names the host left on the clipboard after a shape copy (R6.6).
/// </summary>
/// <remarks>
/// <para>
/// <b>Observe only.</b> An implementation opens the clipboard, enumerates, and closes it
/// without emptying or writing. Nothing parses a drawing payload; the format names are the
/// whole contract (work-item exclusion).
/// </para>
/// <para>
/// <b>A busy clipboard is typed, never thrown.</b> Important-D2 requires the unavailable state
/// to be a distinct failure with an actionable message, so an implementation returns
/// <see cref="ClipboardFormatRead.ForFailure"/> rather than raising.
/// </para>
/// <para>
/// <b>Immediate and bounded.</b> The read runs synchronously on the STA/main thread right
/// after the COM copy call - no polling loops, no sleeps (important-D3).
/// </para>
/// </remarks>
public interface IClipboardFormatReader
{
    /// <summary>
    /// Reads the current clipboard state, returning a typed failure instead of throwing when
    /// the clipboard is busy or unavailable.
    /// </summary>
    /// <returns>The snapshot or the typed failure.</returns>
    ClipboardFormatRead Read();
}
