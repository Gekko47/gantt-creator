using System.Globalization;
using System.Runtime.InteropServices;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="IClipboardFormatReader"/> over the Windows clipboard: opens, enumerates
/// and closes without writing, and resolves each format to the host's own registered name
/// (R6.6).
/// </summary>
/// <remarks>
/// <para>
/// <b>The Win32 route, not Windows Forms.</b> The Office project has no
/// <c>UseWindowsForms</c> and needs none: the <c>user32</c> surface below is exactly what the
/// R6.7 spike measured the evidence with, adds no dependency, and keeps every call behind an
/// <c>internal virtual</c> seam so contract tests drive the busy path without a real
/// clipboard. (The AddIn's Windows Forms usage is for ribbon images, a different boundary.)
/// </para>
/// <para>
/// <b>Names come from the host, not a guess.</b> Registered formats resolve through
/// <c>GetClipboardFormatNameW</c> - this is why the measured evidence names
/// <c>Excel 2007 Internal Shape</c> rather than the old "Office Drawing Shape Format"
/// candidate. <c>RegisterClipboardFormat</c> is a last-resort mapping for a name the
/// host did not report, and an unnamed format still appears by its stable id.
/// </para>
/// <para>
/// <b>Enumeration termination.</b> Once the clipboard is open, a zero from
/// <c>EnumClipboardFormats</c> is the end of the enumeration (the documented
/// <c>ERROR_SUCCESS</c> case); the reader does not add a <c>GetLastError</c> distinction the
/// measured probe did not use. A future build that exposes enumeration failure is a
/// divergence record for R10.1, not a silent change here.
/// </para>
/// <para>
/// <b>COM ownership.</b> This adapter touches no Office object at all - the clipboard is a
/// window-station resource, released by <c>CloseClipboard</c> in a <c>finally</c>.
/// </para>
/// </remarks>
public partial class WindowsClipboardFormatReader : IClipboardFormatReader
{
    /// <summary>
    /// Candidate registered format names the reader resolves to identifiers for its fallback
    /// table. The list is deliberately explicit: an unnamed registered format still appears in
    /// the snapshot by its stable id, so nothing is hidden by an incomplete list.
    /// </summary>
    internal static readonly string[] CandidateFormatNames =
    [
        "Excel 2007 Internal Shape",
        "Art::GVML ClipFormat",
        "Office Drawing Shape Format",
        "Microsoft Office Drawing Object",
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

    /// <summary>Reads the last P/Invoke error. Virtual so a contract test pins the failure token.</summary>
    /// <returns>The Win32 error code, or zero when none was reported.</returns>
    internal virtual int ReadLastPInvokeError() => Marshal.GetLastPInvokeError();

    /// <summary>Registers a format name and returns its identifier, or zero on failure.</summary>
    /// <param name="name">The format name.</param>
    /// <returns>The registered format identifier.</returns>
    internal virtual uint RegisterFormat(string name) => NativeMethods.RegisterClipboardFormat(name);

    /// <summary>
    /// Reads the host's registered name for a format identifier, or <see langword="null"/>
    /// when the format is a standard one or unnamed.
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

    /// <inheritdoc />
    public ClipboardFormatRead Read()
    {
        if (!TryOpenClipboard())
        {
            return ClipboardFormatRead.ForFailure(DescribeOpenFailure(ReadLastPInvokeError()));
        }

        try
        {
            uint sequenceNumber = ReadSequenceNumber();
            List<string> formats = [];
            uint current = 0;
            while ((current = EnumerateFormats(current)) != 0)
            {
                formats.Add(DescribeFormat(current));
            }

            return ClipboardFormatRead.ForSuccess(new ClipboardFormatSnapshot(sequenceNumber, formats));
        }
        finally
        {
            _ = CloseClipboardHandle();
        }
    }

    /// <summary>
    /// Resolves a format identifier to a standard name, a candidate Office name, or a stable
    /// <c>format-0x…</c> fallback.
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

        // Not a standard CF_* id: prefer the host's own registered name (the contract must
        // name what Excel actually put on the clipboard), then the candidate table, then a
        // stable id-derived name so an unexpected format is still recorded.
        string? nativeName = ReadFormatName(id);
        if (!string.IsNullOrEmpty(nativeName))
        {
            return nativeName;
        }

        Dictionary<uint, string> registered = _registeredNames ??= BuildRegisteredNames();
        return registered.TryGetValue(id, out string? candidate)
            ? candidate
            : string.Create(CultureInfo.InvariantCulture, $"format-0x{id:X4}");
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
        : string.Create(CultureInfo.InvariantCulture, $"clipboard-unavailable:0x{win32Error:X8}");

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

        [LibraryImport("user32.dll", EntryPoint = "RegisterClipboardFormatW", SetLastError = true, StringMarshalling = StringMarshalling.Utf16)]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        internal static partial uint RegisterClipboardFormat(string lpszFormat);
    }
}
