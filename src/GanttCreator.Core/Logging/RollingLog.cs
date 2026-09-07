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
    private readonly Lock _gate = new();
    private StreamWriter? _currentWriter;
    private string _currentFilePath = string.Empty;
    private long _currentFileSize;
    private bool _disposed;
    private bool _failed;

    /// <summary>
    /// Creates a new rolling log.
    /// </summary>
    /// <param name="logDirectory">Directory where log files are stored.</param>
    /// <param name="baseName">Base file name (without extension). Default: "gantt-creator".</param>
    /// <param name="maxFileSizeBytes">Maximum size of each log file before rotation. Default: 1 MB.</param>
    /// <param name="maxFileCount">Maximum number of log files to retain. Default: 5.</param>
    /// <param name="redactor">Redactor for privacy-sensitive data. Default: <see cref="Redactor"/>.</param>
    public RollingLog(
        string logDirectory,
        string baseName = "gantt-creator",
        long maxFileSizeBytes = 1_048_576,
        int maxFileCount = 5,
        IRedactor? redactor = null)
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

        _ = Directory.CreateDirectory(_logDirectory);
        RotateIfNeeded();
    }

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
        return baseName;
    }

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
            var line = $"{DateTime.UtcNow:yyyy-MM-ddTHH:mm:ss.fffZ} {redacted}{Environment.NewLine}";
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
                    _currentFilePath = activeLogPath;
                    var appendStream = new FileStream(_currentFilePath, FileMode.Append, FileAccess.Write, FileShare.Read);
                    _currentWriter = new StreamWriter(appendStream, System.Text.Encoding.UTF8);
                    _currentFileSize = fileInfo.Length;
                    return;
                }
            }

            // Rotate existing files in the correct order:
            //   1. Delete the largest numeric rotation(s) if at cap (never the active .log).
            //   2. Shift existing numeric rotations in descending order (.N -> .N+1, ..., .1 -> .2).
            //   3. Move the active .log to .1.log only after shifting, so an existing .1.log is never overwritten.
            // Get only rotated files (exclude the active .log, which has int.MaxValue rotation index),
            // sorted descending so the largest numeric rotation comes first.
            var rotatedFiles = Directory.GetFiles(_logDirectory, $"{_baseName}*.log")
                .Where(f => GetRotationIndex(f) != int.MaxValue)
                .OrderByDescending(GetRotationIndex)
                .ToArray();

            // After rotation we'll have (rotatedFiles.Length + 1) rotated files plus the active .log,
            // so delete from the top when rotatedFiles.Length + 2 exceeds the cap.
            var toDelete = rotatedFiles.Length + 2 - _maxFileCount;
            for (var i = 0; i < toDelete && i < rotatedFiles.Length; i++)
            {
                DeleteFile(rotatedFiles[i]);
            }

            // Re-get remaining rotated files after deletion, still in descending order.
            if (toDelete > 0)
            {
                // IDE0305: Collection initialization can be simplified - explicit LINQ chain for clarity
#pragma warning disable IDE0305
                rotatedFiles = Directory.GetFiles(_logDirectory, $"{_baseName}*.log")
                    .Where(f => GetRotationIndex(f) != int.MaxValue)
                    .OrderByDescending(GetRotationIndex)
                    .ToArray();
#pragma warning restore IDE0305
            }

            // Shift existing numeric rotations in descending order (.N -> .N+1, ..., .1 -> .2).
            // Descending order ensures we never overwrite a file that hasn't been moved yet.
            foreach (var file in rotatedFiles)
            {
                var currentRotation = GetRotationIndex(file);
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
            _currentFilePath = Path.Combine(_logDirectory, $"{_baseName}.log");
            var fileStream = new FileStream(_currentFilePath, FileMode.Create, FileAccess.Write, FileShare.Read);
            _currentWriter = new StreamWriter(fileStream, System.Text.Encoding.UTF8);
            _currentFileSize = 0;
        }
    }

    private static int GetRotationIndex(string path)
    {
        var fileName = Path.GetFileName(path);
        var baseName = Path.GetFileNameWithoutExtension(fileName);
        var lastDot = baseName.LastIndexOf('.');
        return lastDot >= 0 && int.TryParse(baseName[(lastDot + 1)..], out var index) ? index : int.MaxValue;
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
            _currentWriter?.Dispose();
            _currentWriter = null;
        }
        finally
        {
            _gate.Exit();
        }
    }
}
