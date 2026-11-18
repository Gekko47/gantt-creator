using System.Globalization;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// Host-free facts for <see cref="ClipboardFormatProbe"/>. They carry no
/// <c>OfficeIntegration</c> trait, so the ordinary test step runs them: the probe's
/// failure path and name resolution are testable without Excel, PowerPoint, or a
/// populated clipboard.
/// </summary>
public class ClipboardFormatProbeTests
{
    [Fact]
    public void Read_reports_a_typed_failure_when_the_clipboard_cannot_be_opened()
    {
        var probe = new FakeProbe(opens: false, formats: [], registered: [], sequence: 0);

        ClipboardReadResult result = probe.Read();

        Assert.False(result.Succeeded);
        Assert.Null(result.Snapshot);
        Assert.NotNull(result.Failure);
        Assert.StartsWith("clipboard-", result.Failure, StringComparison.Ordinal);
        Assert.Equal(0, probe.CloseCount);
    }

    [Fact]
    public void Read_enumerates_every_format_and_resolves_standard_and_candidate_names()
    {
        uint officeDrawingId = 0xC123;
        uint[] formats = [13, officeDrawingId];
        Dictionary<string, uint> registered = new(StringComparer.Ordinal)
        {
            [ClipboardFormatProbe.OfficeDrawingShapeFormatName] = officeDrawingId,
        };
        var probe = new FakeProbe(opens: true, formats: formats, registered: registered, sequence: 42);

        ClipboardReadResult result = probe.Read();

        Assert.True(result.Succeeded, result.Failure);
        ClipboardSnapshot snapshot = Assert.IsType<ClipboardSnapshot>(result.Snapshot);
        Assert.Equal(42u, snapshot.SequenceNumber);
        Assert.Equal(["CF_UNICODETEXT", ClipboardFormatProbe.OfficeDrawingShapeFormatName], snapshot.Names);
        Assert.True(snapshot.Contains("office drawing shape format"));
        Assert.False(snapshot.Contains(ClipboardFormatProbe.MicrosoftOfficeDrawingObjectName));
        Assert.Equal(1, probe.CloseCount);
    }

    [Fact]
    public void DescribeFormat_falls_back_to_a_stable_id_for_an_unregistered_format()
    {
        var probe = new FakeProbe(opens: true, formats: [], registered: [], sequence: 0);

        Assert.Equal(
            "format-0xDEAD",
            probe.DescribeFormat(0xDEAD),
            StringComparer.Ordinal);
    }

    [Fact]
    public void DescribeFormat_prefers_the_host_registered_name_over_the_candidate_table()
    {
        uint hostFormatId = 0xC0F5;
        Dictionary<string, uint> registered = new(StringComparer.Ordinal)
        {
            [ClipboardFormatProbe.OfficeDrawingShapeFormatName] = 0xC123,
        };
        var probe = new FakeProbe(
            opens: true,
            formats: [],
            registered: registered,
            sequence: 0,
            nativeNames: new Dictionary<uint, string>
            {
                [hostFormatId] = "Microsoft Office Drawing Format",
            });

        Assert.Equal("Microsoft Office Drawing Format", probe.DescribeFormat(hostFormatId));
    }

    [Fact]
    public void KnownStandardFormatName_maps_standard_ids_and_returns_null_for_a_registered_id()
    {
        Assert.Equal("CF_UNICODETEXT", ClipboardFormatProbe.KnownStandardFormatName(13));
        Assert.Equal("CF_DIB", ClipboardFormatProbe.KnownStandardFormatName(8));
        Assert.Null(ClipboardFormatProbe.KnownStandardFormatName(0xC123));
    }

    [Fact]
    public void DescribeOpenFailure_names_the_win32_error_or_the_absent_one()
    {
        Assert.Equal(
            "clipboard-unavailable:0x00000005",
            ClipboardFormatProbe.DescribeOpenFailure(5));
        Assert.Equal(
            "clipboard-open-failed-without-error",
            ClipboardFormatProbe.DescribeOpenFailure(0));
    }

    [Fact]
    public void Contains_is_case_insensitive_and_false_for_a_format_that_is_absent()
    {
        var snapshot = new ClipboardSnapshot(
            7,
            [
                new ClipboardFormatEntry(13, "CF_UNICODETEXT"),
                new ClipboardFormatEntry(0xC123, ClipboardFormatProbe.OfficeDrawingShapeFormatName),
            ]);

        Assert.True(snapshot.Contains("cf_unicodetext"));
        Assert.True(snapshot.Contains(ClipboardFormatProbe.OfficeDrawingShapeFormatName));
        Assert.False(snapshot.Contains("Embed Source"));
    }

    /// <summary>
    /// The live probe is allowed to fail on a machine whose clipboard is busy or
    /// unavailable; it is never allowed to throw out of <see cref="ClipboardFormatProbe.Read"/>.
    /// </summary>
    [Fact]
    public void Live_read_returns_a_typed_outcome_and_never_throws()
    {
        var probe = new ClipboardFormatProbe();

        ClipboardReadResult result = probe.Read();

        if (result.Succeeded)
        {
            Assert.NotNull(result.Snapshot);
        }
        else
        {
            Assert.False(string.IsNullOrWhiteSpace(result.Failure));
        }

        // The snapshot's format names are only consumed by the evidence record; assert
        // the record shape here so a later rename cannot silently empty it.
        if (result.Snapshot is { } snapshot)
        {
            foreach (ClipboardFormatEntry entry in snapshot.Formats)
            {
                Assert.False(string.IsNullOrWhiteSpace(entry.Name));
                Assert.True(
                    entry.Id > 0,
                    string.Create(CultureInfo.InvariantCulture, $"format id {entry.Id} is not positive"));
            }
        }
    }

    /// <summary>
    /// A probe whose Win32 seams are scripted, mirroring the
    /// <c>internal virtual</c> seam pattern the Office harness already uses.
    /// </summary>
    private sealed class FakeProbe(
        bool opens,
        uint[] formats,
        Dictionary<string, uint> registered,
        uint sequence,
        Dictionary<uint, string>? nativeNames = null) : ClipboardFormatProbe
    {
        internal int CloseCount { get; private set; }

        internal override string? ReadFormatName(uint id) =>
            nativeNames is not null && nativeNames.TryGetValue(id, out string? name) ? name : null;

        internal override bool TryOpenClipboard() => opens;

        internal override bool CloseClipboardHandle()
        {
            CloseCount++;
            return true;
        }

        internal override uint EnumerateFormats(uint previous)
        {
            if (previous == 0)
            {
                return formats.Length > 0 ? formats[0] : 0;
            }

            int index = Array.IndexOf(formats, previous);
            return index >= 0 && index + 1 < formats.Length ? formats[index + 1] : 0;
        }

        internal override uint ReadSequenceNumber() => sequence;

        internal override uint RegisterFormat(string name) =>
            registered.TryGetValue(name, out uint id) ? id : 0;
    }
}
