namespace GanttCreator.Core.Logging;

/// <summary>
/// File-based rolling log with size-based rotation and file count cap.
/// Messages are redacted by an injected <see cref="IRedactor"/> before writing.
/// </summary>
public sealed class RollingLog : IRollingLog
{
    private readonly string _logDirectory;
    private readonly string _baseName;
    private readonly long _maxFileSizeBytes;
    private readonly int _maxFileCount;
    private readonly IRedactor _redactor;
    private readonly TimeProvider _timeProvider;
    private readonly Lock _gate = new();
    private StreamWriter? _currentWriter;
    private long _currentFileSize;
    private bool _disposed;

    // Volatile so the lock-free IsFailed read reliably observes the latched
    // flag across threads without acquiring _gate. Every write happens under
    // _gate (or during construction, which has exclusive access).
    private volatile bool _failed;

    /// <summary>
    /// Creates a new rolling log.
    /// </summary>
    /// <param name="logDirectory">Directory where log files are stored.</param>
    /// <param name="baseName">Base file name (without extension). Default: "gantt-creator".</param>
    /// <param name="maxFileSizeBytes">Maximum size of each log file before rotation. Default: 1 MB.</param>
    /// <param name="maxFileCount">Maximum number of log files to retain. Default: 5.</param>
    /// <param name="redactor">Redactor for privacy-sensitive data. Default: <see cref="Redactor"/>.</param>
    /// <param name="timeProvider">Clock boundary for timestamps and testability. Default: <see cref="TimeProvider.System"/>.</param>
    public RollingLog(
        string logDirectory,
        string baseName = "gantt-creator",
        long maxFileSizeBytes = 1_048_576,
        int maxFileCount = 5,
        IRedactor? redactor = null,
        TimeProvider? timeProvider = null)
    {
        if (string.IsNullOrWhiteSpace(logDirectory))
        {
            throw new ArgumentException("Log directory must not be empty.", nameof(logDirectory));
        }
        if (maxFileSizeBytes <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(maxFileSizeBytes), "Must be positive.");
        }
        if (maxFileCount <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(maxFileCount), "Must be positive.");
        }

        // Validate baseName BEFORE any file I/O so validation errors throw ArgumentException,
        // not IOException from the OS when attempting to create files with invalid names.
        _baseName = string.IsNullOrWhiteSpace(baseName) ? "gantt-creator" : ValidateBaseName(baseName);
        _logDirectory = logDirectory;
        _maxFileSizeBytes = maxFileSizeBytes;
        _maxFileCount = maxFileCount;
        _redactor = redactor ?? new Redactor();
        _timeProvider = timeProvider ?? TimeProvider.System;

        _ = Directory.CreateDirectory(_logDirectory);
        RotateIfNeeded();
    }

#pragma warning disable IDE0046  // 'if' statement can be simplified

    // IDE0046 is suppressed here for separate guard clauses on purpose: the
    // rooted-path check and the invalid-filename-char check are two
    // independent failure modes, and one conditional expression would hide
    // which guard fired. Scoped to this method only; every other file still
    // enforces the rule (see Directory.Build.props).
    private static string ValidateBaseName(string baseName)
    {
        if (Path.IsPathRooted(baseName))
        {
            throw new ArgumentException(
                "baseName must not contain path separators, rooted paths, or wildcard characters.",
                nameof(baseName));
        }
        if (baseName.AsSpan().IndexOfAny(Path.GetInvalidFileNameChars()) is >= 0)
        {
            throw new ArgumentException(
                "baseName must not contain path separators, rooted paths, or wildcard characters.",
                nameof(baseName));
        }
        // Reject wildcard characters independently of Path.GetInvalidFileNameChars()
        // so '*' and '?' can never be interpreted as a glob by GetRotatedFilesOldestFirst,
        // even on platforms where they are not in the invalid-file-name set.
        if (baseName.Contains('*', StringComparison.Ordinal) || baseName.Contains('?', StringComparison.Ordinal))
        {
            throw new ArgumentException(
                "baseName must not contain path separators, rooted paths, or wildcard characters.",
                nameof(baseName));
        }
        return baseName;
#pragma warning restore IDE0046
    }

    /// <summary>
    /// Indicates the logger has latched a failure (formatting, redaction,
    /// rotation, or write) and will discard all further writes. The latch is
    /// permanent by design: log failures must never propagate into product
    /// code paths. Diagnostics should consult this property to detect that
    /// logging has stopped.
    /// </summary>
    public bool IsFailed => _failed;

    /// <summary>
    /// Writes a message to the log. The message is automatically redacted.
    /// </summary>
    /// <param name="message">The log message.</param>
    public void Write(string message) => WriteCore(message);

    /// <summary>
    /// Writes a formatted message to the log.
    /// </summary>
    /// <param name="format">Composite format string.</param>
    /// <param name="args">Format arguments.</param>
    public void Write(string format, params object?[] args)
    {
        var message = format;
        try
        {
            if (args is { Length: > 0 })
            {
                message = string.Format(System.Globalization.CultureInfo.InvariantCulture, format, args);
            }
        }
        // CA1031: A logger must never propagate formatting exceptions to callers.
        // Any exception type (FormatException, NullReferenceException from a bad
        // arg, etc.) must be contained and disable further writes.
#pragma warning disable CA1031
        catch
#pragma warning restore CA1031
        {
            _gate.Enter();
            try
            {
                MarkFailed();
            }
            finally
            {
                _gate.Exit();
            }
            return;
        }

        WriteCore(message);
    }

    private void WriteCore(string message)
    {
        if (string.IsNullOrEmpty(message))
        {
            return;
        }

        _gate.Enter();
        try
        {
            if (_disposed || _failed)
            {
                return;
            }

            var redacted = _redactor.Redact(message) ?? string.Empty;
            DateTimeOffset now = _timeProvider.GetUtcNow();
            var line = string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:yyyy-MM-ddTHH:mm:ss.fffZ} {1}{2}", now, redacted, Environment.NewLine);
            var bytes = System.Text.Encoding.UTF8.GetByteCount(line);

            RotateIfNeeded(bytes);
            if (_failed)
            {
                return;
            }

            _currentWriter?.Write(line);
            _currentWriter?.Flush();
            _currentFileSize += bytes;
        }
        // CA1031: A logger must never propagate write/flush/rotation exceptions.
#pragma warning disable CA1031
        catch
#pragma warning restore CA1031
        {
            MarkFailed();
        }
        finally
        {
            _gate.Exit();
        }
    }

    /// <summary>
    /// Marks the logger as failed, preventing all further writes.
    /// Releases the current writer to avoid leaking file handles.
    /// Callers must hold <see cref="_gate"/>, except the constructor
    /// which has exclusive access during initialization.
    /// </summary>
    private void MarkFailed()
    {
        _failed = true;
        try
        {
            _currentWriter?.Dispose();
        }
        // CA1031: In a failure state, a disposal error must not mask the
        // original failure or propagate to callers.
#pragma warning disable CA1031
        catch
#pragma warning restore CA1031
        {
            // Intentionally swallowed
        }
        _currentWriter = null;
    }

    private void RotateIfNeeded(long incomingBytes = 0)
    {
        if (_failed || _disposed)
        {
            return;
        }

        try
        {
            PerformRotation(incomingBytes);
        }
        // CA1031: Rotation I/O failures (move, delete, create) must be
        // contained and disable further writes, never propagated.
#pragma warning disable CA1031
        catch
#pragma warning restore CA1031
        {
            MarkFailed();
        }
    }

    // PerformRotation contains the unchanged rotation logic, extracted so
    // that RotateIfNeeded can wrap it in a try-catch and contain any I/O
    // failure (file move, delete, create) without propagating to callers.
    private void PerformRotation(long incomingBytes)
    {
        if (_currentWriter is null || _currentFileSize + incomingBytes > _maxFileSizeBytes)
        {
            _currentWriter?.Dispose();
            _currentWriter = null;

            var activeLogPath = Path.Combine(_logDirectory, $"{_baseName}.log");

            // If the active log already exists, open in append mode and only rotate if it exceeds the limit.
            // This preserves existing log content when reopening a log with the same base name.
            if (File.Exists(activeLogPath) && _currentFileSize == 0)
            {
                var fileInfo = new FileInfo(activeLogPath);
                if (fileInfo.Length + incomingBytes <= _maxFileSizeBytes)
                {
                    // File exists and has room - open in append mode
                    var appendStream = new FileStream(activeLogPath, FileMode.Append, FileAccess.Write, FileShare.Read);
                    // UTF-8 without a BOM so the reopened file stays byte-exact with
                    // the tracked _currentFileSize; Encoding.UTF8 would emit a preamble.
                    _currentWriter = new StreamWriter(appendStream, new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                    _currentFileSize = fileInfo.Length;
                    return;
                }
            }

            // Rotate existing files in the correct order:
            //   1. Delete the largest numeric rotation(s) if at cap (never the active .log).
            //   2. Shift existing numeric rotations in descending order (.N -> .N+1, ..., .1 -> .2).
            //   3. Move the active .log to .1.log only after shifting, so an existing .1.log is never overwritten.
            // Get only rotated files (the active .log is never included:
            // TryGetRotationIndex returns false for it), ordered descending
            // by rotation index so the oldest (highest-numbered) file comes first.
            List<string> rotatedFiles = GetRotatedFilesOldestFirst();

            // After rotation we'll have (rotatedFiles.Count + 1) rotated files plus the active .log,
            // so delete from the top when rotatedFiles.Count + 2 exceeds the cap.
            var toDelete = rotatedFiles.Count + 2 - _maxFileCount;
            for (var i = 0; i < toDelete && i < rotatedFiles.Count; i++)
            {
                DeleteFile(rotatedFiles[i]);
            }

            // Re-get remaining rotated files after deletion, still descending
            // rotation index (oldest first).
            if (toDelete > 0)
            {
                rotatedFiles = GetRotatedFilesOldestFirst();
            }

            // Shift existing numeric rotations in descending order (.N -> .N+1, ..., .1 -> .2).
            // Descending order ensures we never overwrite a file that hasn't been moved yet.
            foreach (var file in rotatedFiles)
            {
                // Helper output is guaranteed to carry a rotation index.
                _ = TryGetRotationIndex(file, _baseName, out var currentRotation);
                var newRotation = currentRotation + 1;
                var newPath = Path.Combine(_logDirectory, $"{_baseName}.{newRotation}.log");
                File.Move(file, newPath);
            }

            // Move active .log to .1.log only after shifting, so an existing .1.log (now .2) is never overwritten.
            // With a cap of 1 there is no rotated slot, so the active log is deleted instead.
            if (File.Exists(activeLogPath))
            {
                if (_maxFileCount == 1)
                {
                    DeleteFile(activeLogPath);
                }
                else
                {
                    File.Move(activeLogPath, Path.Combine(_logDirectory, $"{_baseName}.1.log"));
                }
            }

            // Create new active log file
            var newActivePath = Path.Combine(_logDirectory, $"{_baseName}.log");
            var fileStream = new FileStream(newActivePath, FileMode.Create, FileAccess.Write, FileShare.Read);
            // UTF-8 without a BOM so the created file's size equals the tracked
            // Encoding.UTF8 byte count and rotation remains byte-exact.
            _currentWriter = new StreamWriter(fileStream, new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            _currentFileSize = 0;
        }
    }

    /// <summary>
    /// Enumerates the rotated files for this base name ordered by
    /// descending rotation index: the oldest (highest-numbered) file
    /// first, down to <c>{baseName}.1.log</c>. The active
    /// <c>{baseName}.log</c> is never included. Only files of the exact
    /// shape <c>{baseName}.{index}.log</c> with ASCII-digit indexes are
    /// matched, so unrelated files that merely share the base-name prefix
    /// (for example <c>{baseName}-backup.log</c> or <c>{baseName}2.log</c>
    /// written by another tool) are never shifted or deleted by rotation.
    /// </summary>
    private List<string> GetRotatedFilesOldestFirst()
    {
        return [.. Directory.GetFiles(_logDirectory, $"{_baseName}*.log")
            .Select(f => (Path: f, Index: TryGetRotationIndex(f, _baseName, out var i) ? i : (int?)null))
            .Where(x => x.Index is not null)
            .OrderByDescending(x => x.Index!.Value)
            .Select(x => x.Path)];
    }

    private static bool TryGetRotationIndex(string path, string baseName, out int index)
    {
        index = 0;
        var fileName = Path.GetFileName(path);
        const string suffix = ".log";

        // Strip the suffix FIRST, then the "{baseName}." prefix. For the
        // active "{baseName}.log" the prefix and suffix overlap (its dot is
        // the dot of ".log"), so slicing before stripping the suffix would
        // produce an inverted range and throw.
        if (!fileName.EndsWith(suffix, StringComparison.Ordinal))
        {
            return false;
        }

        var stem = fileName[..^suffix.Length];
        var prefix = $"{baseName}.";
        if (!stem.StartsWith(prefix, StringComparison.Ordinal))
        {
            return false;
        }

        var middle = stem[prefix.Length..];
        if (middle.Length == 0)
        {
            // The active {baseName}.log itself; never a rotated file.
            return false;
        }

        // ASCII digits only: rejects signs, separators, and culture-specific
        // digit shapes; the subsequent parse is then culture-safe. Overflowing
        // values fail the parse and the file is left untouched (safe direction).
        foreach (var c in middle)
        {
            if (c is < '0' or > '9')
            {
                return false;
            }
        }

        return int.TryParse(middle, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out index);
    }

    private static void DeleteFile(string path) => File.Delete(path);

    /// <summary>
    /// Disposes the rolling log, releasing the current writer.
    /// </summary>
    public void Dispose()
    {
        _gate.Enter();
        try
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            try
            {
                _currentWriter?.Dispose();
            }
            // CA1031: a disposal error during teardown must not propagate
            // from Dispose nor mask an earlier latched failure.
#pragma warning disable CA1031
            catch
#pragma warning restore CA1031
            {
                // Contained: nothing actionable while tearing down.
            }
            _currentWriter = null;
        }
        finally
        {
            _gate.Exit();
        }
    }
}
