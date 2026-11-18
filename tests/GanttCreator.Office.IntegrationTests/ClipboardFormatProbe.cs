using System.Runtime.InteropServices;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// One clipboard format present in a <see cref="ClipboardSnapshot"/>.
/// </summary>
/// <param name="Id">The Win32 clipboard format identifier.</param>
/// <param name="Name">
/// The standard <c>CF_*</c> name, one of the candidate Office format names, or a
/// stable <c>format-0x…</c> fallback for an unnamed registered format.
/// </param>
internal sealed record ClipboardFormatEntry(uint Id, string Name);

/// <summary>
/// The clipboard state observed at one moment: its sequence number and the
/// complete format set.
/// </summary>
/// <param name="SequenceNumber">
/// <c>GetClipboardSequenceNumber</c> at the read. A change between two reads means
/// the clipboard content was replaced, not merely that a format stopped rendering.
/// </param>
/// <param name="Formats">Every format the clipboard reports, in enumeration order.</param>
internal sealed record ClipboardSnapshot(uint SequenceNumber, IReadOnlyList<ClipboardFormatEntry> Formats)
{
    /// <summary>Whether a format with the given name is present, ordinal case-insensitive.</summary>
    /// <param name="name">The format name to look for.</param>
    /// <returns><see langword="true"/> when present.</returns>
    internal bool Contains(string name) => Formats.Any(
        format => string.Equals(format.Name, name, StringComparison.OrdinalIgnoreCase));

    /// <summary>The format names, in enumeration order, for the evidence record.</summary>
    internal IReadOnlyList<string> Names => [.. Formats.Select(format => format.Name)];
}

/// <summary>
/// The outcome of one clipboard read: a snapshot or a typed failure reason.
/// </summary>
/// <param name="Succeeded">Whether the clipboard could be opened and enumerated.</param>
/// <param name="Snapshot">The observed state when <paramref name="Succeeded"/> is true.</param>
/// <param name="Failure">
/// A stable diagnostic token when <paramref name="Succeeded"/> is false. A busy or
/// unavailable clipboard is a distinct typed outcome, never a hang and never an
/// exception crossing the test body.
/// </param>
internal sealed record ClipboardReadResult(bool Succeeded, ClipboardSnapshot? Snapshot, string? Failure)
{
    /// <summary>Builds a successful result.</summary>
    /// <param name="snapshot">The observed clipboard state.</param>
    /// <returns>The result.</returns>
    internal static ClipboardReadResult ForSuccess(ClipboardSnapshot snapshot) => new(true, snapshot, null);

    /// <summary>Builds a failed result.</summary>
    /// <param name="failure">The diagnostic token.</param>
    /// <returns>The result.</returns>
    internal static ClipboardReadResult ForFailure(string failure) => new(false, null, failure);
}

/// <summary>
/// Reads the Windows clipboard's format set for the R6.7 copy-after-delete spike.
/// </summary>
/// <remarks>
/// <para>
/// This helper only observes: it opens the clipboard, enumerates, and closes it
/// without emptying or writing. It resolves names two ways — the documented
/// <c>CF_*</c> identifiers directly, and registered Office format names by calling
/// <c>RegisterClipboardFormatW</c> for each documented candidate (registering an
/// already-registered name returns its existing identifier, so presence is a
/// deterministic comparison against the enumerated set rather than a guess from a
/// name buffer).
/// </para>
/// <para>
/// The Win32 calls are virtual seams in the <see cref="OfficeFixture"/> style, so a
/// contract test can drive the busy-clipboard failure path and the format-naming
/// table without owning a real clipboard.
/// </para>
/// </remarks>
internal partial class ClipboardFormatProbe
{
    /// <summary>
    /// The format name Excel and PowerPoint use for an Office drawing/shape copy.
    /// Recorded in the evidence matrix; presence is asserted by the spike bodies.
    /// </summary>
    internal const string OfficeDrawingShapeFormatName = "Office Drawing Shape Format";

    /// <summary>
    /// The alternative Excel drawing-object format name (the
    /// <c>Worksheet.PasteSpecial</c> "Microsoft Office Drawing Object" string
    /// equivalent documented by Microsoft).
    /// </summary>
    internal const string MicrosoftOfficeDrawingObjectName = "Microsoft Office Drawing Object";

    /// <summary>
    /// The format name measured on Excel 16.0.20527.20196 for a shape copy. Recorded as a
    /// candidate so presence is assertable even before the native name is resolved.
    /// </summary>
    internal const string Excel2007InternalShapeFormatName = "Excel 2007 Internal Shape";

    /// <summary>The Office Art GVML drawing clipboard format measured on the same host.</summary>
    internal const string ArtGvmlClipFormatName = "Art::GVML ClipFormat";

    /// <summary>
    /// Candidate registered format names the probe resolves to identifiers. The list
    /// is deliberately explicit: an unnamed registered format still appears in the
    /// snapshot by its stable id, so nothing is hidden by an incomplete list.
    /// </summary>
    internal static readonly string[] CandidateFormatNames =
    [
        OfficeDrawingShapeFormatName,
        MicrosoftOfficeDrawingObjectName,
        Excel2007InternalShapeFormatName,
        ArtGvmlClipFormatName,
        "Embed Source",
        "Embedded Object",
        "Object Descriptor",
        "Link Source",
        "EnhancedMetafile",
        "MetaFilePict",
        "PNG",
        "JFIF",
        "GIF",
        "HTML Format",
        "Rich Text Format",
    ];

    private Dictionary<uint, string>? _registeredNames;

    /// <summary>Opens the clipboard for reading. Virtual so a contract test can force the busy path.</summary>
    /// <returns><see langword="true"/> when the clipboard was opened.</returns>
    internal virtual bool TryOpenClipboard() => NativeMethods.OpenClipboard(IntPtr.Zero);

    /// <summary>Closes a clipboard opened by <see cref="TryOpenClipboard"/>.</summary>
    /// <returns>The Win32 result, ignored by the read but observable to a test seam.</returns>
    internal virtual bool CloseClipboardHandle() => NativeMethods.CloseClipboard();

    /// <summary>Enumerates the format after <paramref name="previous"/>; zero ends the enumeration.</summary>
    /// <param name="previous">The last format identifier returned, or zero to start.</param>
    /// <returns>The next format identifier, or zero.</returns>
    internal virtual uint EnumerateFormats(uint previous) => NativeMethods.EnumClipboardFormats(previous);

    /// <summary>Reads the clipboard sequence number, which changes when the content is replaced.</summary>
    /// <returns>The current sequence number.</returns>
    internal virtual uint ReadSequenceNumber() => NativeMethods.GetClipboardSequenceNumber();

    /// <summary>Registers a format name and returns its identifier, or zero on failure.</summary>
    /// <param name="name">The format name.</param>
    /// <returns>The registered format identifier.</returns>
    internal virtual uint RegisterFormat(string name) => NativeMethods.RegisterClipboardFormat(name);

    /// <summary>
    /// Reads the host's registered name for a format identifier, or <see langword="null"/>
    /// when the format is a standard one or unnamed. This is what makes the evidence record
    /// name the host's own Office formats instead of guessing from a candidate list.
    /// </summary>
    /// <param name="id">The clipboard format identifier.</param>
    /// <returns>The registered name, or <see langword="null"/>.</returns>
    internal virtual string? ReadFormatName(uint id)
    {
        const int bufferCharacters = 256;
        IntPtr buffer = Marshal.AllocHGlobal(bufferCharacters * sizeof(char));
        try
        {
            int length = NativeMethods.GetClipboardFormatName(id, buffer, bufferCharacters);
            return length > 0 ? Marshal.PtrToStringUni(buffer, length) : null;
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    /// <summary>
    /// Reads the current clipboard state, returning a typed failure instead of
    /// throwing when the clipboard is busy or unavailable.
    /// </summary>
    /// <returns>The snapshot or the typed failure.</returns>
    internal ClipboardReadResult Read()
    {
        if (!TryOpenClipboard())
        {
            return ClipboardReadResult.ForFailure(
                DescribeOpenFailure(Marshal.GetLastPInvokeError()));
        }

        try
        {
            uint sequenceNumber = ReadSequenceNumber();
            List<ClipboardFormatEntry> formats = [];
            uint current = 0;
            while ((current = EnumerateFormats(current)) != 0)
            {
                formats.Add(new ClipboardFormatEntry(current, DescribeFormat(current)));
            }

            return ClipboardReadResult.ForSuccess(new ClipboardSnapshot(sequenceNumber, formats));
        }
        finally
        {
            _ = CloseClipboardHandle();
        }
    }

    /// <summary>
    /// Resolves a format identifier to a standard name, a candidate Office name, or
    /// a stable <c>format-0x…</c> fallback.
    /// </summary>
    /// <param name="id">The clipboard format identifier.</param>
    /// <returns>A non-empty display name.</returns>
    internal string DescribeFormat(uint id)
    {
        string? standard = KnownStandardFormatName(id);
        if (standard is not null)
        {
            return standard;
        }

        // Not a standard CF_* id: prefer the host's own registered name (the evidence
        // record must name what Excel actually put on the clipboard), then the candidate
        // table, then a stable id-derived name so an unexpected format is still recorded.
        string? nativeName = ReadFormatName(id);
        if (!string.IsNullOrEmpty(nativeName))
        {
            return nativeName;
        }

        Dictionary<uint, string> registered = _registeredNames ??= BuildRegisteredNames();
        return registered.TryGetValue(id, out string? candidate) ? candidate : $"format-0x{id:X4}";
    }

    /// <summary>
    /// The documented <c>CF_*</c> identifier names, or <see langword="null"/> for an
    /// identifier that is not a standard format.
    /// </summary>
    /// <param name="id">The clipboard format identifier.</param>
    /// <returns>The standard name, or <see langword="null"/>.</returns>
    internal static string? KnownStandardFormatName(uint id) => id switch
    {
        1 => "CF_TEXT",
        2 => "CF_BITMAP",
        3 => "CF_METAFILEPICT",
        4 => "CF_SYLK",
        5 => "CF_DIF",
        6 => "CF_TIFF",
        7 => "CF_OEMTEXT",
        8 => "CF_DIB",
        9 => "CF_PALETTE",
        10 => "CF_PENDATA",
        11 => "CF_RIFF",
        12 => "CF_WAVE",
        13 => "CF_UNICODETEXT",
        14 => "CF_ENHMETAFILE",
        15 => "CF_HDROP",
        16 => "CF_LOCALE",
        17 => "CF_DIBV5",
        _ => null,
    };

    /// <summary>
    /// A stable diagnostic token for a failed clipboard open.
    /// </summary>
    /// <param name="win32Error">The last P/Invoke error, or zero when none was reported.</param>
    /// <returns>The token.</returns>
    internal static string DescribeOpenFailure(int win32Error) => win32Error == 0
        ? "clipboard-open-failed-without-error"
        : $"clipboard-unavailable:0x{win32Error:X8}";

    private Dictionary<uint, string> BuildRegisteredNames()
    {
        Dictionary<uint, string> names = [];
        foreach (string name in CandidateFormatNames)
        {
            uint id = RegisterFormat(name);
            if (id != 0)
            {
                _ = names.TryAdd(id, name);
            }
        }

        return names;
    }

    private static partial class NativeMethods
    {
        [LibraryImport("user32.dll", SetLastError = true)]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static partial bool OpenClipboard(IntPtr hWndNewOwner);

        [LibraryImport("user32.dll", SetLastError = true)]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static partial bool CloseClipboard();

        [LibraryImport("user32.dll", SetLastError = true)]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        internal static partial uint EnumClipboardFormats(uint format);

        [LibraryImport("user32.dll")]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        internal static partial uint GetClipboardSequenceNumber();

        [LibraryImport("user32.dll", EntryPoint = "GetClipboardFormatNameW", SetLastError = true)]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        internal static partial int GetClipboardFormatName(uint format, IntPtr lpszFormatName, int cchMaxCount);

        [LibraryImport(
            "user32.dll",
            EntryPoint = "RegisterClipboardFormatW",
            StringMarshalling = StringMarshalling.Utf16)]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        internal static partial uint RegisterClipboardFormat(string format);
    }
}
