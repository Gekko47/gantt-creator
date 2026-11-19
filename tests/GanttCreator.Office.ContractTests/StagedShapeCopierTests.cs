namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R6.6: success is a verified drawing format, never a non-throwing copy; a busy clipboard is
/// a distinct typed refusal; and the sequence is copy-then-verify.
/// </summary>
/// <remarks>
/// Every refusal arm has a positive test that constructs the bad input (AGENTS.md validator
/// rule), over the shared <see cref="FakeShapeCopyPort"/> and
/// <see cref="FakeClipboardFormatReader"/>.
/// </remarks>
public sealed class StagedShapeCopierTests
{
    private static StagedGroupOutcome VerifiedGroup(string? name = null) =>
        new(name ?? ShapeStagingTag.GroupName, [], null, null, null);

    [Fact]
    public void The_copy_runs_before_the_read_and_only_a_verified_pair_is_success()
    {
        List<string> sequence = [];
        FakeShapeCopyPort port = new() { AfterCopy = () => sequence.Add("copy") };
        FakeClipboardFormatReader reader = new()
        {
            Result = FakeClipboardFormatReader.Snapshot(
                "Excel 2007 Internal Shape",
                "Art::GVML ClipFormat",
                "PNG"),
            AfterRead = () => sequence.Add("read"),
        };

        StagingCopyResult result = StagedShapeCopier.Copy(VerifiedGroup(), port, reader);

        Assert.True(result.Succeeded, "Copy refused: " + result.Refusal);
        Assert.Null(result.Refusal);
        Assert.Null(result.PortRefusal);
        Assert.Equal(["copy", "read"], sequence);
        Assert.Equal([ShapeStagingTag.GroupName], port.RequestedNames);
        Assert.Contains("Excel 2007 Internal Shape", result.Detail, StringComparison.Ordinal);
    }

    [Fact]
    public void An_absent_drawing_format_refuses_and_never_reports_success()
    {
        FakeShapeCopyPort port = new();
        FakeClipboardFormatReader missingArt = new()
        {
            Result = FakeClipboardFormatReader.Snapshot("Excel 2007 Internal Shape", "PNG"),
        };

        StagingCopyResult first = StagedShapeCopier.Copy(VerifiedGroup(), port, missingArt);

        Assert.False(first.Succeeded);
        Assert.Equal(StagedCopyRefusal.ExpectedFormatAbsent, first.Refusal);
        Assert.Contains("Art::GVML ClipFormat", first.Detail, StringComparison.Ordinal);
        Assert.Equal(1, missingArt.ReadCalls);

        FakeClipboardFormatReader missingShape = new()
        {
            Result = FakeClipboardFormatReader.Snapshot("Art::GVML ClipFormat"),
        };

        StagingCopyResult second = StagedShapeCopier.Copy(VerifiedGroup(), port, missingShape);

        Assert.False(second.Succeeded);
        Assert.Equal(StagedCopyRefusal.ExpectedFormatAbsent, second.Refusal);
        Assert.Contains("Excel 2007 Internal Shape", second.Detail, StringComparison.Ordinal);
    }

    [Fact]
    public void A_busy_clipboard_is_a_distinct_typed_refusal_with_an_actionable_detail()
    {
        FakeShapeCopyPort port = new();
        FakeClipboardFormatReader reader = new()
        {
            Result = ClipboardFormatRead.ForFailure("clipboard-unavailable:0x00000005"),
        };

        StagingCopyResult result = StagedShapeCopier.Copy(VerifiedGroup(), port, reader);

        Assert.False(result.Succeeded);
        Assert.Equal(StagedCopyRefusal.ClipboardUnavailable, result.Refusal);
        Assert.Contains("clipboard-unavailable:0x00000005", result.Detail, StringComparison.Ordinal);
        Assert.Contains("try again", result.Detail, StringComparison.Ordinal);
    }

    [Fact]
    public void A_copy_port_refusal_carries_the_port_reason_and_never_reads_the_clipboard()
    {
        FakeShapeCopyPort port = new() { Refusal = ShapeWriteRefusal.NotFound };
        FakeClipboardFormatReader reader = new();

        StagingCopyResult result = StagedShapeCopier.Copy(VerifiedGroup(), port, reader);

        Assert.False(result.Succeeded);
        Assert.Equal(StagedCopyRefusal.CopyRefused, result.Refusal);
        Assert.Equal(ShapeWriteRefusal.NotFound, result.PortRefusal);
        Assert.Equal(0, reader.ReadCalls);
    }

    [Fact]
    public void A_blank_group_name_is_refused_before_any_host_call()
    {
        FakeShapeCopyPort port = new();
        FakeClipboardFormatReader reader = new();

        StagingCopyResult result = StagedShapeCopier.Copy(VerifiedGroup("  "), port, reader);

        Assert.False(result.Succeeded);
        Assert.Equal(StagedCopyRefusal.BlankGroupName, result.Refusal);
        Assert.Empty(port.RequestedNames);
        Assert.Equal(0, reader.ReadCalls);
    }

    [Fact]
    public void Null_inputs_are_refused_without_touching_anything()
    {
        FakeShapeCopyPort port = new();
        FakeClipboardFormatReader reader = new();

        Assert.Equal(StagedCopyRefusal.NullOutcome, StagedShapeCopier.Copy(null, port, reader).Refusal);
        Assert.Equal(StagedCopyRefusal.NullPort, StagedShapeCopier.Copy(VerifiedGroup(), null, reader).Refusal);
        Assert.Equal(StagedCopyRefusal.NullReader, StagedShapeCopier.Copy(VerifiedGroup(), port, null).Refusal);

        Assert.Empty(port.RequestedNames);
        Assert.Equal(0, reader.ReadCalls);
    }

    [Fact]
    public void The_required_formats_are_the_measured_drawing_pair()
    {
        Assert.Equal(
            ["Excel 2007 Internal Shape", "Art::GVML ClipFormat"],
            StagedShapeCopier.RequiredDrawingFormatNames);
    }
}
