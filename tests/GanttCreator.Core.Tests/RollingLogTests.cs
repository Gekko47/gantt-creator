using System.Globalization;
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
        catch (UnauthorizedAccessException)
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
    public void Log_file_tracks_utf8_byte_count_exactly_without_preamble()
    {
        // Regression: RollingLog must write UTF-8 without a preamble. Encoding.UTF8
        // is a BOM-emitting UTF8Encoding in .NET 10, so the real file size must equal
        // the size tracked via Encoding.UTF8.GetByteCount or rotation drifts.
        var fixedTime = new DateTimeOffset(2030, 1, 2, 3, 4, 5, 678, TimeSpan.Zero);
        var message = "Preamble-free byte accounting";
        var line = string.Format(System.Globalization.CultureInfo.InvariantCulture,
            "{0:yyyy-MM-ddTHH:mm:ss.fffZ} {1}{2}", fixedTime, message, Environment.NewLine);
        var lineBytes = (long)System.Text.Encoding.UTF8.GetByteCount(line);

        using (var log1 = new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: 1024, maxFileCount: 3,
            timeProvider: new FixedTimeProvider(fixedTime)))
        {
            log1.Write(message);
        }

        var activeFile = Path.Combine(_testLogDir, $"{_baseName}.log");
        var activeInfo = new FileInfo(activeFile);
        Assert.Equal(lineBytes, (long)activeInfo.Length);

        // Append reopen must not inject a preamble either; the reopened file size
        // stays equal to the sum of the UTF-8 line byte counts.
        using (var log2 = new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: 1024, maxFileCount: 3,
            timeProvider: new FixedTimeProvider(fixedTime)))
        {
            log2.Write(message);
        }

        var reopenedInfo = new FileInfo(activeFile);
        Assert.Equal(2 * lineBytes, (long)reopenedInfo.Length);
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
    public void File_count_cap_of_one_deletes_active_log_on_rotation()
    {
        // With cap=1 there is no rotated slot. Rotation must delete the active
        // .log instead of moving it to .1.log, otherwise the directory would
        // hold .log plus .1.log, exceeding the cap.
        using var log = new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: 50, maxFileCount: 1);

        // Force multiple rotations
        for (var i = 0; i < 10; i++)
        {
            log.Write($"Message number {i} with enough content to rotate");
        }

        var files = Directory.GetFiles(_testLogDir, $"{_baseName}*.log");

        // Cap respected: at most one file (the active .log)
        Assert.True(files.Length <= 1, $"Expected at most 1 log file (cap), got {files.Length}");

        // The active log must always exist after rotation
        Assert.Contains(files, f => Path.GetFileName(f) == $"{_baseName}.log");

        // No rotated file may ever be created with cap=1
        Assert.DoesNotContain(files, f => Path.GetFileName(f) == $"{_baseName}.1.log");
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

        // Reopen the same base name after log1 has been disposed; the
        // constructor must append rather than truncate. (Previously this
        // comment claimed we simulated a fresh process restart without
        // disposing log1 first, which the scoped `using` did not do.)
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
    public void Write_null_or_empty_is_skipped_and_whitespace_is_written()
    {
        using (var log = new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: 1024, maxFileCount: 3))
        {
            log.Write("");
            log.Write(null!);
            log.Write("   ");
        }

        var files = Directory.GetFiles(_testLogDir, $"{_baseName}*.log");
        _ = Assert.Single(files);
        var lines = File.ReadAllLines(files[0]);
        _ = Assert.Single(lines);
        Assert.EndsWith("   ", lines[0], StringComparison.Ordinal);
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
        Assert.Equal("baseName", ex.ParamName);
    }

    [Fact]
    public void BaseName_rejects_directory_separator()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(() =>
            new RollingLog(_testLogDir, "subdir/file", maxFileSizeBytes: 1024, maxFileCount: 3));
        Assert.Contains("baseName", ex.Message, StringComparison.Ordinal);
        Assert.Equal("baseName", ex.ParamName);
    }

    [Fact]
    public void BaseName_rejects_wildcard()
    {
        ArgumentException ex1 = Assert.Throws<ArgumentException>(() =>
            new RollingLog(_testLogDir, "file*", maxFileSizeBytes: 1024, maxFileCount: 3));
        Assert.Equal("baseName", ex1.ParamName);

        ArgumentException ex2 = Assert.Throws<ArgumentException>(() =>
            new RollingLog(_testLogDir, "file?", maxFileSizeBytes: 1024, maxFileCount: 3));
        Assert.Equal("baseName", ex2.ParamName);
    }

    [Fact]
    public void Constructor_rejects_whitespace_only_log_directory()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(() =>
            new RollingLog("   ", _baseName, maxFileSizeBytes: 1024, maxFileCount: 3));
        Assert.Equal("logDirectory", ex.ParamName);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Constructor_rejects_non_positive_max_file_size_bytes(long maxFileSizeBytes)
    {
        ArgumentOutOfRangeException ex = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: maxFileSizeBytes, maxFileCount: 3));
        Assert.Equal("maxFileSizeBytes", ex.ParamName);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Constructor_rejects_non_positive_max_file_count(int maxFileCount)
    {
        ArgumentOutOfRangeException ex = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: 1024, maxFileCount: maxFileCount));
        Assert.Equal("maxFileCount", ex.ParamName);
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

    [Fact]
    public void Write_format_failure_is_contained_without_latching_failure()
    {
        using var log = new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: 1024, maxFileCount: 3);

        // Unclosed format item throws FormatException -- must be contained
        // to the current call without latching the permanent failure flag.
        Exception ex = Record.Exception(() => log.Write("{0", "test"));
        Assert.Null(ex);
        Assert.False(log.IsFailed);

        // Subsequent writes still proceed after the skipped malformed message.
        log.Write("Should appear");
        log.Dispose();

        var files = Directory.GetFiles(_testLogDir, $"{_baseName}*.log");
        _ = Assert.Single(files);
        var content = File.ReadAllText(files[0]);
        Assert.Contains("Should appear", content, StringComparison.Ordinal);
    }

    [Fact]
    public void Write_redaction_failure_is_contained_and_disables_writes()
    {
        using var log = new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: 1024, maxFileCount: 3, redactor: new FailingRedactor());

        Exception ex = Record.Exception(() => log.Write("Test message"));
        Assert.Null(ex);

        // After failure, writes are silently disabled
        log.Write("Should not appear");

        var files = Directory.GetFiles(_testLogDir, $"{_baseName}*.log");
        _ = Assert.Single(files);
        var content = File.ReadAllText(files[0]);
        Assert.DoesNotContain("Test message", content, StringComparison.Ordinal);
        Assert.DoesNotContain("Should not appear", content, StringComparison.Ordinal);
    }

    [Fact]
    public void Write_rotation_failure_is_contained_and_disables_writes()
    {
        using var log = new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: 50, maxFileCount: 5);

        // Write a small message to create the active log file and advance its size
        log.Write("Initial");

        // Open an exclusive handle on .1.log to block the rotation that
        // moves the active log into the .1 slot.
        var rotatedPath = Path.Combine(_testLogDir, $"{_baseName}.1.log");
        using var blockingHandle = new FileStream(
            rotatedPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);

        // A message large enough to force rotation; PerformRotation cannot
        // move files because .1.log is locked -- failure is contained.
        Exception ex = Record.Exception(() => log.Write("Second message long enough to exceed the 50-byte limit"));
        Assert.Null(ex);

        // After failure, writes are silently disabled
        log.Write("Should not appear");

        var activePath = Path.Combine(_testLogDir, $"{_baseName}.log");
        Assert.True(File.Exists(activePath));
        var content = File.ReadAllText(activePath);
        Assert.Contains("Initial", content, StringComparison.Ordinal);
        Assert.DoesNotContain("Should not appear", content, StringComparison.Ordinal);
    }

    private sealed class FailingRedactor : IRedactor
    {
        [return: System.Diagnostics.CodeAnalysis.NotNullIfNotNull(nameof(input))]
        public string? Redact(string? input) => throw new InvalidOperationException("Simulated redactor failure");
    }

    private sealed class FixedTimeProvider(DateTimeOffset utcNow) : TimeProvider
    {
        public override DateTimeOffset GetUtcNow() => utcNow;
    }

    [Fact]
    public void Write_uses_injected_time_provider_for_timestamp()
    {
        var fixedTime = new DateTimeOffset(2030, 1, 2, 3, 4, 5, 678, TimeSpan.Zero);
        using (var log = new RollingLog(
            _testLogDir, _baseName, maxFileSizeBytes: 1024, maxFileCount: 3,
            timeProvider: new FixedTimeProvider(fixedTime)))
        {
            log.Write("Clock test message");
        }

        var files = Directory.GetFiles(_testLogDir, $"{_baseName}*.log");
        _ = Assert.Single(files);
        var content = File.ReadAllText(files[0]);
        Assert.StartsWith("2030-01-02T03:04:05.678Z Clock test message", content, StringComparison.Ordinal);
    }

    [Fact]
    public void IsFailed_is_false_until_a_failure_is_latched()
    {
        using var log = new RollingLog(
            _testLogDir, _baseName, maxFileSizeBytes: 1024, maxFileCount: 3,
            redactor: new FailingRedactor());

        Assert.False(log.IsFailed);
        log.Write("Trigger failure");
        Assert.True(log.IsFailed);
    }

    [Fact]
    public void IsFailed_is_exposed_through_the_IRollingLog_abstraction()
    {
        // Consumers that program against the abstraction must be able to
        // detect the permanent failure latch without downcasting.
        // CA1859 suggests the concrete type for performance; the whole point
        // of this test is the interface reference, so the suggestion is
        // intentionally declined here.
#pragma warning disable CA1859
        IRollingLog log = new RollingLog(
            _testLogDir, _baseName, maxFileSizeBytes: 1024, maxFileCount: 3,
            redactor: new FailingRedactor());
#pragma warning restore CA1859
        try
        {
            Assert.False(log.IsFailed);
            log.Write("Trigger failure");
            Assert.True(log.IsFailed);
        }
        finally
        {
            log.Dispose();
        }
    }

    [Fact]
    public void Rotation_never_touches_files_that_merely_share_the_base_name_prefix()
    {
        // Both files match the rotation enumeration glob "{base}*.log" but
        // are not rotation files of this log; rotation must leave them alone.
        var foreign1 = Path.Combine(_testLogDir, $"{_baseName}2.log");
        var foreign2 = Path.Combine(_testLogDir, $"{_baseName}-backup.log");
        File.WriteAllText(foreign1, "FOREIGN-1-CONTENT");
        File.WriteAllText(foreign2, "FOREIGN-2-CONTENT");

        using var log = new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: 50, maxFileCount: 3);
        for (var i = 0; i < 10; i++)
        {
            log.Write($"Message number {i} with enough content to rotate");
        }

        Assert.Equal("FOREIGN-1-CONTENT", File.ReadAllText(foreign1));
        Assert.Equal("FOREIGN-2-CONTENT", File.ReadAllText(foreign2));
    }

    [Fact]
    public void Constructor_latches_failure_when_log_directory_path_points_to_an_existing_file()
    {
        // Regression: Directory.CreateDirectory on a path occupied by an
        // existing file throws IOException. The constructor must contain
        // that failure through the permanent latch (like every other
        // logging failure) instead of escaping into product code.
        var occupiedPath = Path.Combine(_testLogDir, $"occupied-{Guid.NewGuid():N}");
        File.WriteAllText(occupiedPath, "this is a file, not a directory");

        using var log = new RollingLog(
            occupiedPath, _baseName, maxFileSizeBytes: 1024, maxFileCount: 3);

        Assert.True(log.IsFailed);
        log.Write("must be discarded");
        Assert.True(log.IsFailed);
    }

    [Fact]
    public void Write_format_overload_is_culture_invariant()
    {
        // Checklist A culture-roundtrip: the formatted overload must use the
        // invariant culture regardless of the ambient locale.
        CultureInfo previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
            using (var log = new RollingLog(_testLogDir, _baseName, maxFileSizeBytes: 1024, maxFileCount: 3))
            {
                log.Write("Value: {0}", 1.5);
            }

            var files = Directory.GetFiles(_testLogDir, $"{_baseName}*.log");
            _ = Assert.Single(files);
            var content = File.ReadAllText(files[0]);
            Assert.Contains("Value: 1.5", content, StringComparison.Ordinal);
        }
        finally
        {
            CultureInfo.CurrentCulture = previous;
        }
    }
}
