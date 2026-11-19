namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R6.6: the Win32 clipboard reader is a typed observation - a busy clipboard is a failure
/// token rather than an exception, names resolve from the host, and the handle is always
/// closed.
/// </summary>
/// <remarks>
/// Every host call is a subclassed <c>internal virtual</c> seam, so these facts own no real
/// clipboard and cannot interfere with another test's clipboard state.
/// </remarks>
public sealed class WindowsClipboardFormatReaderTests
{
    [Fact]
    public void A_busy_clipboard_is_a_typed_failure_and_the_handle_is_not_closed()
    {
        ScriptedReader reader = new() { Open = false, Error = 5 };

        ClipboardFormatRead read = reader.Read();

        Assert.False(read.Succeeded);
        Assert.Null(read.Snapshot);
        Assert.Equal("clipboard-unavailable:0x00000005", read.Failure);
        Assert.Equal(["open"], reader.Calls);
    }

    [Fact]
    public void Enumeration_resolves_standard_formats_and_closes_the_clipboard()
    {
        ScriptedReader reader = new()
        {
            Sequence = 548,
            Formats = new Queue<uint>([3, 17]),
        };

        ClipboardFormatRead read = reader.Read();

        Assert.True(read.Succeeded, read.Failure);
        ClipboardFormatSnapshot snapshot = Assert.IsType<ClipboardFormatSnapshot>(read.Snapshot);
        Assert.Equal(548u, snapshot.SequenceNumber);
        Assert.Equal(["CF_METAFILEPICT", "CF_DIBV5"], snapshot.FormatNames);
        Assert.Equal(["open", "close"], reader.Calls);
    }

    [Fact]
    public void A_registered_format_resolves_to_the_host_name()
    {
        ScriptedReader reader = new()
        {
            Formats = new Queue<uint>([0xC123]),
            Names = new Dictionary<uint, string> { [0xC123] = "Excel 2007 Internal Shape" },
        };

        ClipboardFormatRead read = reader.Read();

        ClipboardFormatSnapshot snapshot = Assert.IsType<ClipboardFormatSnapshot>(read.Snapshot);
        Assert.True(snapshot.Contains("excel 2007 internal shape"));
    }

    [Fact]
    public void An_unnamed_registered_format_still_appears_by_its_stable_id()
    {
        ScriptedReader reader = new() { Formats = new Queue<uint>([0xC124]) };

        ClipboardFormatRead read = reader.Read();

        ClipboardFormatSnapshot snapshot = Assert.IsType<ClipboardFormatSnapshot>(read.Snapshot);
        Assert.Equal(["format-0xC124"], snapshot.FormatNames);
    }

    [Fact]
    public void A_throwing_enumeration_still_closes_the_clipboard()
    {
        ScriptedReader reader = new() { ThrowDuringEnumeration = true };

        _ = Assert.Throws<InvalidOperationException>(reader.Read);

        Assert.Equal(["open", "close"], reader.Calls);
    }

    [Fact]
    public void A_zero_error_open_failure_has_its_own_stable_token() =>
        Assert.Equal("clipboard-open-failed-without-error", WindowsClipboardFormatReader.DescribeOpenFailure(0));

    private sealed class ScriptedReader : WindowsClipboardFormatReader
    {
        public bool Open { get; init; } = true;

        public int Error { get; init; }

        public uint Sequence { get; init; } = 1;

        public Queue<uint> Formats { get; init; } = new();

        public Dictionary<uint, string> Names { get; init; } = [];

        public bool ThrowDuringEnumeration { get; init; }

        public List<string> Calls { get; } = [];

        internal override bool TryOpenClipboard()
        {
            Calls.Add("open");
            return Open;
        }

        internal override bool CloseClipboardHandle()
        {
            Calls.Add("close");
            return true;
        }

        internal override int ReadLastPInvokeError() => Error;

        internal override uint ReadSequenceNumber() => Sequence;

        internal override uint EnumerateFormats(uint previous) => ThrowDuringEnumeration
            ? throw new InvalidOperationException("enumeration failed")
            : Formats.Count > 0 ? Formats.Dequeue() : 0;

        internal override string? ReadFormatName(uint id) =>
            Names.TryGetValue(id, out string? name) ? name : null;

        internal override uint RegisterFormat(string name) => 0;
    }
}
