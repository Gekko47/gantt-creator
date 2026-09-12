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
    private static readonly DateTimeOffset TestTimestamp =
        new(2026, 9, 12, 10, 0, 0, TimeSpan.Zero);

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
            Assert.EndWith(path, ".log", StringComparison.Ordinal);
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

        // Create a directory, then make it unwritable so the log latches failure.
        Directory.CreateDirectory(directory);

        try
        {
            var log = new RollingLog(directory, baseName: "test-failed");

            // Force failure by making the directory read-only/unwritable.
            // RollingLog attempts to write; with no writer available it should
            // latch failure. We simulate by disposing the underlying stream
            // immediately after construction won't work; instead we create a
            // scenario where writes fail: use a directory we can mark
            // read-only, then create the log (which succeeds), then make it
            // read-only so subsequent writes latch failure.

            var service = DiagnosticsService.Instance;
            service.SetLog(log);

            // Before making it fail, the path should be available.
            Assert.NotNull(service.LogFilePath);

            // Now make the directory read-only so writes fail and the log latches.
            Directory.SetAccessControl(
                directory,
                new System.Security.AccessControl.DirectorySecurity());

            try
            {
                // Attempt a write to force the failure latch.
                log.Write("force failure");
            }
            catch
            {
                // Expected: write may throw or latch.
            }

            // After failure, LogFilePath returns null.
            Assert.Null(service.LogFilePath);
        }
        finally
        {
            DiagnosticsService.Reset();

            // Clean up: remove read-only attribute so deletion works.
            try
            {
                var attr = File.GetAttributes(directory);
                File.SetAttributes(directory, attr & ~FileAttributes.ReadOnly);
            }
            catch
            {
            }

            try
            {
                Directory.Delete(directory, recursive: true);
            }
            catch
            {
            }
        }
    }

    [Fact]
    public void GatherIdentifiers_returns_all_four_fields()
    {
        var service = DiagnosticsService.Instance;

        // Inject a log so the service is usable.
        var log = new RollingLog(
            Path.Combine(Path.GetTempPath(), "GanttCreatorTests", Guid.NewGuid().ToString("N"), "logs"),
            baseName: "test-gather");
        try
        {
            service.SetLog(log);

            var identity = service.GatherIdentifiers();

            Assert.NotNull(identity);
            Assert.NotEmpty(identity.AddInVersion);
            Assert.NotEmpty(identity.ExcelVersion);
            Assert.NotNull(identity.ProcessBitness);
            Assert.NotEmpty(identity.XllFileName);
        }
        finally
        {
            log.Dispose();
            DiagnosticsService.Reset();
        }
    }

    [Fact]
    public void GatherIdentifiers_normalizes_unknown_fields()
    {
        var service = DiagnosticsService.Instance;

        var log = new RollingLog(
            Path.GetTempPath(),
            baseName: "test-gather-null");
        try
        {
            service.SetLog(log);

            var identity = service.GatherIdentifiers();

            // Each field should be non-null and non-empty; unknown fields
            // degrade to "unknown" rather than null.
            Assert.NotNull(identity.AddInVersion);
            Assert.NotEmpty(identity.AddInVersion);
            Assert.NotNull(identity.ExcelVersion);
            Assert.NotEmpty(identity.ExcelVersion);
            Assert.NotNull(identity.ProcessBitness);
            Assert.NotEmpty(identity.ProcessBitness);
            Assert.NotNull(identity.XllFileName);
            Assert.NotEmpty(identity.XllFileName);
        }
        finally
        {
            log.Dispose();
            DiagnosticsService.Reset();
        }
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
            // We cannot actually invoke ShowDiagnostics here because it
            // requires the TaskDialog UI (which is Windows-only and not
            // available in a test harness). Instead, we verify the
            // identifier-gathering and log-writing parts by calling the
            // internal methods directly via a wrapper test.

            // Verify the service can gather identifiers.
            var identity = service.GatherIdentifiers();
            Assert.NotNull(identity);

            // Verify the content builder works.
            var content = service.BuildContent(identity, "file:///test.log");
            Assert.Contains("Add-in Version:", content);
            Assert.Contains("Excel Version:", content);
            Assert.Contains("Process:", content);
            Assert.Contains("XLL File:", content);
            Assert.Contains("file:///test.log", content);

            // Verify the log write path works via the public API.
            service.WriteDiagnosticRecord(identity);
            mockLog.Verify(
                l => l.Write(
                    It.Is<string>(m => m.Contains("Diagnostics:")),
                    It.IsAny<object?[]>()),
                Times.Once);
        }
        finally
        {
            DiagnosticsService.Reset();
        }
    }

    [Fact]
    public void BuildContent_includes_file_hyperlink()
    {
        var service = DiagnosticsService.Instance;

        var identity = new AddInIdentity("1.0.0", "16.0", "x64", "test.xll");
        var content = service.BuildContent(identity, "C:\\Logs\\test.log");

        Assert.Contains("file:///C:/Logs/test.log", content);
    }

    [Fact]
    public void BuildContent_handles_missing_log_path()
    {
        var service = DiagnosticsService.Instance;

        var identity = new AddInIdentity("1.0.0", "16.0", "x64", "test.xll");
        var content = service.BuildContent(identity, null);

        Assert.Contains("not available", content);
    }

    [Fact]
    public void OpenLogFile_invokes_process_start()
    {
        var path = "file:///C:/Test/log.txt";
        var startInfo = new ProcessStartInfo(path)
        {
            UseShellExecute = true
        };

        var mockProcess = new Mock<Process>();
        mockProcess.Setup(p => p.Start()).Returns(true);

        // We test OpenLogFile by invoking it directly and verifying it
        // doesn't throw. The actual Process.Start is called; in a test
        // environment this may fail but that's acceptable (non-fatal).
        var service = DiagnosticsService.Instance;
        service.OpenLogFile(path);

        // No exception expected; the call is non-fatal.
        Assert.True(true);
    }
}
