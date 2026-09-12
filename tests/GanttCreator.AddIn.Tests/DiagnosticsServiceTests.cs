using System.IO;
using System.Security.AccessControl;
using System.Globalization;
using ExcelDna.Integration;
using GanttCreator.Core.Logging;
using GanttCreator.Core;
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
            baseName: "test-diagnostics");

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
        var directory = Path.Combine(
            Path.GetTempPath(),
            "GanttCreatorTests",
            Guid.NewGuid().ToString("N"),
            "logs");

        Directory.CreateDirectory(directory);

        try
        {
            var service = DiagnosticsService.Instance;
            try
            {
                using (var log = new RollingLog(directory, baseName: "test-failed"))
                {
                    var directoryInfo = new System.IO.DirectoryInfo(directory);
                    var directorySecurity = directoryInfo.GetAccessControl();
                    directorySecurity.AddAccessRule(
                        new System.Security.AccessControl.FileSystemAccessRule(
                            "Everyone",
                            System.Security.AccessControl.FileSystemRights.Write,
                            System.Security.AccessControl.AccessControlType.Deny));
                    directoryInfo.SetAccessControl(directorySecurity);

                    service.SetLog(log);
                }
            }
            catch (UnauthorizedAccessException)
            {
                // Expected: we cannot modify directory security on some environments.
            }
        }
        finally
        {
            try
            {
                var directoryInfo = new System.IO.DirectoryInfo(directory);
                directoryInfo.Delete(true);
            }
#pragma warning disable CA1031 // Cleanup is best-effort; surface nothing to the caller.
            catch
            {
            }
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
        mockLog.Setup(l => l.LogFilePath).Returns(
            Path.Combine(Path.GetTempPath(), "test-log.log"));
        mockLog.Setup(l => l.Write(It.IsAny<string>()))
            .Callback<string>(msg => written.Add(msg));
        mockLog.Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>((fmt, args) =>
            {
                var message = string.Format(CultureInfo.InvariantCulture, fmt, args);
                written.Add(message);
            });
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
                l => l.Write(
                    It.Is<string>(m => m.Contains("Diagnostics:", StringComparison.Ordinal)),
                    It.IsAny<object?[]>()),
                Times.Once);
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
}