namespace GanttCreator.Core.Tests;

/// <summary>
/// The PNG-export settings contract (schema version 28, R8.8): one stored form per target,
/// a fail-open consumer parse, and a strict parse for the settings reader.
/// </summary>
public sealed class GanttPngExportTests
{
    [Fact]
    public void Format_writes_one_stored_form_per_target()
    {
        Assert.Equal(GanttPngExport.ClipboardValue, GanttPngExport.Format(PngExportTarget.Clipboard));
        Assert.Equal(GanttPngExport.FileValue, GanttPngExport.Format(PngExportTarget.File));
    }

    [Theory]
    [InlineData(null, PngExportTarget.Clipboard)]
    [InlineData("", PngExportTarget.Clipboard)]
    [InlineData("garbage", PngExportTarget.Clipboard)]
    [InlineData("clipboard", PngExportTarget.Clipboard)]
    [InlineData("file", PngExportTarget.File)]
    [InlineData(" FILE ", PngExportTarget.File)]
    public void Parse_reads_the_stored_target_and_fails_open_to_the_clipboard(
        string? stored,
        PngExportTarget expected)
    {
        Assert.Equal(expected, GanttPngExport.Parse(stored));
    }

    [Theory]
    [InlineData("clipboard", true)]
    [InlineData("file", true)]
    [InlineData(" File ", true)]
    [InlineData(null, false)]
    [InlineData("", false)]
    [InlineData("garbage", false)]
    public void TryParse_is_strict_for_the_settings_reader(string? stored, bool expected)
    {
        Assert.Equal(expected, GanttPngExport.TryParse(stored, out _));
    }
}
