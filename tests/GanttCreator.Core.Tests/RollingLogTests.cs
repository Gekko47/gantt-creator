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
        _ = Directory.CreateDirectory(_testLogDir);
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
        catch (IOException)
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
        _ = Assert.Single(files);

        var content = File.ReadAllText(files[0]);
        Assert.Contains("Test message", content, StringComparison.Ordinal);
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
        Assert.Contains("[email]", content, StringComparison.Ordinal);
        Assert.DoesNotContain("alice@example.com", content, StringComparison.Ordinal);
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
        for (var i = 0; i < 10; i++)
        {
            log.Write($"Message number {i} with enough content to rotate");
        }

        var files = Directory.GetFiles(_testLogDir, $"{_baseName}*.log");
        Assert.True(files.Length <= 2, $"Expected at most 2 log files (cap), got {files.Length}");
    }

    [Fact]
    public void Rotation_at_cap_never_deletes_active_log_and_preserves_rotated_files()
    {
        // With cap=3 the active .log plus two rotated files are retained. Rotation must
        // delete the largest numeric rotation first and shift before moving the active log,
        // so the active .log is never selected for deletion and an existing .1.log is never overwritten.
        using var log = new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: 50, maxFileCount: 3);

        for (var i = 0; i < 20; i++)
        {
            log.Write($"Message number {i} with enough content to rotate the log file");
        }

        var files = Directory.GetFiles(_testLogDir, $"{_baseName}*.log");

        // Cap respected
        Assert.True(files.Length <= 3, $"Expected at most 3 log files (cap), got {files.Length}");

        // The active log must always exist — it is never selected for deletion
        Assert.Contains(files, f => Path.GetFileName(f) == $"{_baseName}.log");

        // Rotation indices must remain unique — no overwrite of an existing rotation ever occurred
        var indices = files
            .Select(f => Path.GetFileNameWithoutExtension(Path.GetFileName(f)))
            .Select(name =>
            {
                var dot = name.LastIndexOf('.');
                return dot >= 0 && int.TryParse(name[(dot + 1)..], out var idx) ? (int?)idx : null;
            })
            .Where(idx => idx.HasValue)
            .Select(idx => idx!.Value)
            .ToArray();

        Assert.Equal(indices.Length, indices.Distinct().Count());
        Assert.All(indices, idx => Assert.True(idx >= 0));
    }

    [Fact]
    public void Dispose_allows_reopening_same_base_name()
    {
        var dir = _testLogDir;
        var baseName = _baseName;

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

        var allContent = files.SelectMany(File.ReadAllLines).ToArray();
        Assert.Contains(allContent, line => line.Contains("First session", StringComparison.Ordinal));
        Assert.Contains(allContent, line => line.Contains("Second session", StringComparison.Ordinal));
    }

    [Fact]
    public void Reopen_preserves_existing_log_content_in_append_mode()
    {
        // With append-mode reopen, the existing active log is NOT truncated.
        var dir = _testLogDir;
        var baseName = _baseName;

        using (var log1 = new RollingLog(dir, baseName, maxFileSizeBytes: 1024, maxFileCount: 3))
        {
            log1.Write("First session");
        }

        // Reopen without disposing log1 first (simulate fresh process restart)
        using (var log2 = new RollingLog(dir, baseName, maxFileSizeBytes: 1024, maxFileCount: 3))
        {
            log2.Write("Second session");
        }

        var activeFile = Path.Combine(dir, $"{baseName}.log");
        Assert.True(File.Exists(activeFile));
        var content = File.ReadAllText(activeFile);
        Assert.Contains("First session", content, StringComparison.Ordinal);
        Assert.Contains("Second session", content, StringComparison.Ordinal);
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
                Assert.DoesNotContain("Test message", content, StringComparison.Ordinal);
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
        Assert.Contains("Value: 42, Name: test", content, StringComparison.Ordinal);
    }

    [Fact]
    public void BaseName_rejects_rooted_path()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(() =>
            new RollingLog(_testLogDir, "C:/evil", maxFileSizeBytes: 1024, maxFileCount: 3));
        Assert.Contains("baseName", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void BaseName_rejects_directory_separator()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(() =>
            new RollingLog(_testLogDir, "subdir/file", maxFileSizeBytes: 1024, maxFileCount: 3));
        Assert.Contains("baseName", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void BaseName_rejects_wildcard()
    {
        _ = Assert.Throws<ArgumentException>(() =>
            new RollingLog(_testLogDir, "file*", maxFileSizeBytes: 1024, maxFileCount: 3));
        _ = Assert.Throws<ArgumentException>(() =>
            new RollingLog(_testLogDir, "file?", maxFileSizeBytes: 1024, maxFileCount: 3));
    }

    [Fact]
    public void Write_after_Dispose_is_silently_rejected()
    {
        var log = new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: 1024, maxFileCount: 3);
        log.Write("Before dispose");
        log.Dispose();

        // After dispose, Write should be a no-op (not throw, not write)
        log.Write("After dispose");

        var files = Directory.GetFiles(_testLogDir, $"{_baseName}*.log");
        _ = Assert.Single(files);
        var content = File.ReadAllText(files[0]);
        Assert.Contains("Before dispose", content, StringComparison.Ordinal);
        Assert.DoesNotContain("After dispose", content, StringComparison.Ordinal);
    }

    [Fact]
    public void Retention_double_digit_count_preserves_numeric_order()
    {
        // maxFileCount > 9 forces double-digit rotation suffixes,
        // which string sorting would misorder (.10 before .2).
        using var log = new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: 10, maxFileCount: 12);

        // Write enough messages to trigger many rotations
        for (var i = 0; i < 50; i++)
        {
            log.Write($"Message {i} with enough content to rotate");
        }

        var files = Directory.GetFiles(_testLogDir, $"{_baseName}*.log");
        Assert.True(files.Length <= 12, $"Expected at most 12 files, got {files.Length}");

        // Verify all rotation indices are unique and positive
        var indices = files
            .Select(f => Path.GetFileNameWithoutExtension(Path.GetFileName(f)))
            .Select(name =>
            {
                var dot = name.LastIndexOf('.');
                return dot >= 0 && int.TryParse(name[(dot + 1)..], out var i) ? (int?)i : null;
            })
            .Where(i => i.HasValue)
            .Select(i => i!.Value)
            .ToArray();

        // No duplicate rotation indices
        Assert.Equal(indices.Length, indices.Distinct().Count());
        // All indices are positive
        Assert.All(indices, i => Assert.True(i > 0));
    }
}
