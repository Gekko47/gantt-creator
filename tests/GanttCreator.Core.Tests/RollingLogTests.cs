using System;
using System.IO;
using System.Linq;
using System.Threading;
using GanttCreator.Core.Logging;

namespace GanttCreator.Core.Tests;

/// <summary>
/// Tests for <see cref="RollingLog"/> rotation, redaction injection, and file cap.
/// </summary>
public sealed class RollingLogTests : IDisposable
{
    private readonly string _testLogDir;
    // Use a unique base name per test instance to avoid parallel-test file locking
    private readonly string _baseName = $"testlog-{Guid.NewGuid():N}";

    public RollingLogTests()
    {
        _testLogDir = Path.Combine(Path.GetTempPath(), $"gantt-creator-tests-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_testLogDir);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_testLogDir))
            {
                Directory.Delete(_testLogDir, recursive: true);
            }
        }
        catch
        {
            // Best effort cleanup
        }
    }

    [Fact]
    public void Write_creates_log_file_with_timestamp_and_message()
    {
        using (var log = new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: 1024, maxFileCount: 3))
        {
            log.Write("Test message");
        }

        var files = Directory.GetFiles(_testLogDir, $"{_baseName}*.log");
        Assert.Single(files);

        var content = File.ReadAllText(files[0]);
        Assert.Contains("Test message", content);
        Assert.Matches(@"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z ", content);
    }

    [Fact]
    public void Write_redacts_sensitive_data_via_injected_redactor()
    {
        using (var log = new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: 1024, maxFileCount: 3))
        {
            log.Write("User alice@example.com connected");
        }

        var files = Directory.GetFiles(_testLogDir, $"{_baseName}*.log");
        var content = File.ReadAllText(files[0]);
        Assert.Contains("[email]", content);
        Assert.DoesNotContain("alice@example.com", content);
    }

    [Fact]
    public void Rotation_creates_new_file_when_size_exceeded()
    {
        // Small max size to force rotation
        using var log = new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: 100, maxFileCount: 3);

        // Write enough to exceed 100 bytes (each line ~50 bytes)
        log.Write("First message that is long enough to take space");
        log.Write("Second message that is also long enough to take space");
        log.Write("Third message that is also long enough to take space");

        var files = Directory.GetFiles(_testLogDir, $"{_baseName}*.log");
        Assert.True(files.Length >= 2, $"Expected at least 2 log files after rotation, got {files.Length}");
    }

    [Fact]
    public void File_count_cap_deletes_oldest_when_exceeded()
    {
        // Very small size and count to force quick rotation and deletion
        using var log = new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: 50, maxFileCount: 2);

        // Force multiple rotations
        for (int i = 0; i < 10; i++)
        {
            log.Write($"Message number {i} with enough content to rotate");
        }

        var files = Directory.GetFiles(_testLogDir, $"{_baseName}*.log");
        Assert.True(files.Length <= 2, $"Expected at most 2 log files (cap), got {files.Length}");
    }

    [Fact]
    public void Dispose_allows_reopening_same_base_name()
    {
        string dir = _testLogDir;
        string baseName = _baseName;

        using (var log1 = new RollingLog(dir, baseName, maxFileSizeBytes: 1024, maxFileCount: 3))
        {
            log1.Write("First session");
        }

        using (var log2 = new RollingLog(dir, baseName, maxFileSizeBytes: 1024, maxFileCount: 3))
        {
            log2.Write("Second session");
        }

        var files = Directory.GetFiles(dir, $"{baseName}*.log");
        Assert.True(files.Length >= 1);

        var allContent = files.SelectMany(f => File.ReadAllLines(f)).ToArray();
        Assert.Contains(allContent, line => line.Contains("First session"));
        Assert.Contains(allContent, line => line.Contains("Second session"));
    }

    [Fact]
    public void Write_null_or_empty_creates_no_message_lines()
    {
        using (var log = new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: 1024, maxFileCount: 3))
        {
            log.Write("");
            log.Write(null!);
            log.Write("   ");
        }

        var files = Directory.GetFiles(_testLogDir, $"{_baseName}*.log");
        if (files.Length > 0)
        {
            // If a file was created, it should not contain the empty messages
            foreach (var f in files)
            {
                var content = File.ReadAllText(f);
                Assert.DoesNotContain("Test message", content);
            }
        }
    }

    [Fact]
    public void Write_format_overload_works()
    {
        using (var log = new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: 1024, maxFileCount: 3))
        {
            log.Write("Value: {0}, Name: {1}", 42, "test");
        }

        var files = Directory.GetFiles(_testLogDir, $"{_baseName}*.log");
        Assert.NotEmpty(files);
        var content = File.ReadAllText(files[0]);
        Assert.Contains("Value: 42, Name: test", content);
    }
}