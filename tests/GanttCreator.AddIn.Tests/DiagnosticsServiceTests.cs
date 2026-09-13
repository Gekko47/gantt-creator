using System.Globalization;
using System.IO;
using System.Windows.Forms;
using ExcelDna.Integration;
using GanttCreator.Core;
using GanttCreator.Core.Logging;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contract tests for <see cref="DiagnosticsService"/>: singleton lifetime,
/// identifier gathering, content building, log writing, and hyperlink handling.
/// None of these require Excel.
/// </summary>
public class DiagnosticsServiceTests
{
    [Fact]
    public void Instance_returns_same_singleton()
    {
        var first = DiagnosticsService.Instance;
        var second = DiagnosticsService.Instance;

        Assert.Same(first, second);
    }

    [Fact]
    public void Reset_clears_singleton()
    {
        var before = DiagnosticsService.Instance;
        DiagnosticsService.Reset();
        var after = DiagnosticsService.Instance;

        Assert.NotSame(before, after);
        Assert.NotNull(after);
    }

    [Fact]
    public void SetLog_rejects_null()
    {
        var service = DiagnosticsService.Instance;

        Assert.Throws<ArgumentNullException>(() => service.SetLog(null!));
    }

    [Fact]
    public void LogFilePath_returns_null_when_no_log_is_injected()
    {
        DiagnosticsService.Reset();
        var service = DiagnosticsService.Instance;

        Assert.Null(service.LogFilePath);
    }

    [Fact]
    public void LogFilePath_returns_active_log_path()
    {
        var log = new RollingLog(
            Path.Combine(Path.GetTempPath(), "GanttCreatorTests", Guid.NewGuid().ToString("N"), "logs"),
            baseName: "test-diagnostics"
        );

        try
        {
            var service = DiagnosticsService.Instance;
            service.SetLog(log);

            string? path = service.LogFilePath;

            Assert.NotNull(path);
            Assert.EndsWith(".log", path, StringComparison.Ordinal);
            Assert.Contains("test-diagnostics.log", path, StringComparison.Ordinal);
        }
        finally
        {
            log.Dispose();
            DiagnosticsService.Reset();
        }
    }

    [Fact]
    public void LogFilePath_returns_null_when_log_has_failed()
    {
        // Create a file path (not a directory) to cause RollingLog initialization
        // to fail deterministically: Directory.CreateDirectory throws when the
        // target path is an existing file, and RollingLog catches that and
        // latches MarkFailed so LogFilePath returns null.
        var tempDir = Path.Combine(Path.GetTempPath(), "GanttCreatorTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var filePath = Path.Combine(tempDir, "not-a-directory");
        File.WriteAllText(filePath, "placeholder");

        try
        {
            var service = DiagnosticsService.Instance;

            using var log = new RollingLog(filePath, baseName: "test-failed");
            Assert.True(log.IsFailed, "RollingLog should latch a failure when the log directory is an existing file.");

            service.SetLog(log);

            Assert.Null(service.LogFilePath);
        }
        finally
        {
            try
            {
                File.Delete(filePath);
                Directory.Delete(tempDir);
            }
#pragma warning disable CA1031 // Cleanup is best-effort; surface nothing to the caller.
            catch { }
#pragma warning restore CA1031
            DiagnosticsService.Reset();
        }
    }

    [Fact]
    public void GatherIdentifiers_returns_all_four_fields()
    {
        var identity = DiagnosticsService.GatherIdentifiers();
        Assert.NotNull(identity);
        Assert.NotNull(identity.AddInVersion);
        Assert.NotNull(identity.ExcelVersion);
        Assert.NotNull(identity.ProcessBitness);
        Assert.NotNull(identity.XllFileName);
    }

    [Fact]
    public void GatherIdentifiers_normalizes_unknown_fields()
    {
        var identity = DiagnosticsService.GatherIdentifiers();
        Assert.NotNull(identity.AddInVersion);
        Assert.NotEmpty(identity.AddInVersion);
        Assert.NotNull(identity.ExcelVersion);
        Assert.NotEmpty(identity.ExcelVersion);
        Assert.NotNull(identity.ProcessBitness);
        Assert.NotEmpty(identity.ProcessBitness);
        Assert.NotNull(identity.XllFileName);
        Assert.NotEmpty(identity.XllFileName);
    }

    [Fact]
    public void BuildContent_includes_identifier_fields()
    {
        var identity = new AddInIdentity("1.0.0", "16.0", "x64", "test.xll");
        var content = DiagnosticsService.BuildContent(identity, "file:///test.log");
        Assert.Contains("Add-in Version:", content, StringComparison.Ordinal);
        Assert.Contains("Excel Version:", content, StringComparison.Ordinal);
        Assert.Contains("Process:", content, StringComparison.Ordinal);
        Assert.Contains("XLL File:", content, StringComparison.Ordinal);
        Assert.Contains("file:///test.log", content, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildContent_includes_file_hyperlink()
    {
        var identity = new AddInIdentity("1.0.0", "16.0", "x64", "test.xll");
        var content = DiagnosticsService.BuildContent(identity, "C:\\Logs\\test.log");

        Assert.Contains("file:///C:/Logs/test.log", content, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildContent_handles_missing_log_path()
    {
        var identity = new AddInIdentity("1.0.0", "16.0", "x64", "test.xll");
        var content = DiagnosticsService.BuildContent(identity, null);

        Assert.Contains("not available", content, StringComparison.Ordinal);
    }

    [Fact]
    public void ShowDiagnostics_writes_log_record()
    {
        var written = new List<string>();
        var mockLog = new Mock<IRollingLog>();
        mockLog.Setup(l => l.IsFailed).Returns(false);
        mockLog.Setup(l => l.LogFilePath).Returns(Path.Combine(Path.GetTempPath(), "test-log.log"));
        mockLog.Setup(l => l.Write(It.IsAny<string>())).Callback<string>(msg => written.Add(msg));
        mockLog
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>(
                (fmt, args) =>
                {
                    var message = string.Format(CultureInfo.InvariantCulture, fmt, args);
                    written.Add(message);
                }
            );
        mockLog.Setup(l => l.Dispose()).Verifiable();

        var service = DiagnosticsService.Instance;
        service.SetLog(mockLog.Object);

        try
        {
            var identity = DiagnosticsService.GatherIdentifiers();
            Assert.NotNull(identity);

            var content = DiagnosticsService.BuildContent(identity, "file:///test.log");
            Assert.Contains("Add-in Version:", content, StringComparison.Ordinal);
            Assert.Contains("Excel Version:", content, StringComparison.Ordinal);
            Assert.Contains("Process:", content, StringComparison.Ordinal);
            Assert.Contains("XLL File:", content, StringComparison.Ordinal);
            Assert.Contains("file:///test.log", content, StringComparison.Ordinal);

            service.WriteDiagnosticRecord(identity);
            mockLog.Verify(
                l => l.Write(It.Is<string>(m => m.Contains("Diagnostics:", StringComparison.Ordinal)), It.IsAny<object?[]>()),
                Times.Once
            );
        }
        finally
        {
            DiagnosticsService.Reset();
        }
    }

    [Fact]
    public void OpenLogFile_invokes_process_start()
    {
        var path = "file:///C:/Test/log.txt";

        DiagnosticsService.OpenLogFile(path);

        // No exception expected; the call is non-fatal.
        Assert.True(true);
    }

    // ------------------------------------------------------------------
    // TaskDialog page contract (regression: the previous comctl32 P/Invoke
    // threw EntryPointNotFoundException on every ribbon click; the managed
    // WinForms TaskDialog page replaces it).
    // ------------------------------------------------------------------

    [Fact]
    public void CreatePage_wires_caption_heading_text_links_and_close_button()
    {
        var page = TaskDialogApi.CreatePage("Title", "Heading", "Content", _ => { });

        Assert.Equal("Title", page.Caption);
        Assert.Equal("Heading", page.Heading);
        Assert.Equal("Content", page.Text);
        Assert.True(page.EnableLinks, "The content carries <a href> markup, so EnableLinks must be on.");
        Assert.Contains(TaskDialogButton.Close, page.Buttons);
    }

    [Fact]
    public void LinkClicked_handler_invokes_open_action_with_link_href()
    {
        // TaskDialogLinkClickedEventArgs(String) is the documented public
        // constructor; this reproduces the event without showing a dialog.
        var args = new TaskDialogLinkClickedEventArgs("file:///C:/Logs/test.log");

        string? opened = null;
        TaskDialogApi.HandleLinkClicked(args, link => opened = link);

        Assert.Equal("file:///C:/Logs/test.log", opened);
    }

    [Fact]
    public void HandleLinkClicked_rejects_null_args()
    {
        Assert.Throws<ArgumentNullException>(
            () => TaskDialogApi.HandleLinkClicked(null!, _ => { }));
    }

    [Fact]
    public void StripLinkMarkup_removes_anchor_markup_for_the_fallback_dialog()
    {
        const string content =
            "Add-in Version: 1.0.0\r\n\r\nLog File: <a href=\"file:///C:/Logs/test log.log\">Open the log file</a>\r\n";

        string plain = DiagnosticsService.StripLinkMarkup(content);

        Assert.DoesNotContain("<a href=", plain, StringComparison.Ordinal);
        Assert.DoesNotContain("</a>", plain, StringComparison.Ordinal);
        Assert.Contains("Log File: Open the log file", plain, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildContent_escapes_access_key_ampersands()
    {
        var identity = new AddInIdentity("1&0.0", "16&0", "x64", "test&file.xll");

        string content = DiagnosticsService.BuildContent(identity, null);

        Assert.Contains("Add-in Version: 1&&0.0", content, StringComparison.Ordinal);
        Assert.Contains("test&&file.xll", content, StringComparison.Ordinal);
    }

    [Fact]
    public void WriteDiagnosticsError_writes_the_failure_to_the_log()
    {
        var written = new List<string>();
        var mockLog = new Mock<IRollingLog>();
        mockLog.Setup(l => l.IsFailed).Returns(false);
        mockLog.Setup(l => l.LogFilePath).Returns(Path.Combine(Path.GetTempPath(), "test-log.log"));
        mockLog
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>(
                (fmt, args) => written.Add(string.Format(CultureInfo.InvariantCulture, fmt, args))
            );

        var service = DiagnosticsService.Instance;
        service.SetLog(mockLog.Object);

        try
        {
            DiagnosticsService.WriteDiagnosticsError(new InvalidOperationException("simulated"));

            Assert.Contains(written, m => m.Contains("OnDiagnosticsClick failed", StringComparison.Ordinal));
        }
        finally
        {
            DiagnosticsService.Reset();
        }
    }

    [Fact]
    public void WriteDiagnosticsError_never_throws_even_without_a_log()
    {
        DiagnosticsService.Reset();

        // No log injected: the record degrades to nothing and must not throw.
        DiagnosticsService.WriteDiagnosticsError(new InvalidOperationException("simulated"));
    }
}
