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

    private static string ValidateBaseName(string name)
    {
        if (Path.IsPathRooted(name))
        {
            throw new ArgumentException(
                "baseName must not contain path separators, rooted paths, or wildcard characters.",
                nameof(name));
        }
        if (name.AsSpan().IndexOfAny(Path.GetInvalidFileNameChars()) is >= 0)
        {
            throw new ArgumentException(
                "baseName must not contain path separators, rooted paths, or wildcard characters.",
                nameof(name));
        }
        return name;
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
        if (args is { Length: > 0 })
        {
            WriteCore(string.Format(System.Globalization.CultureInfo.InvariantCulture, format, args));
        }
        else
        {
            WriteCore(format);
        }
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
            if (_disposed)
            {
                return;
            }

            var redacted = _redactor.Redact(message);
            var line = $"{DateTime.UtcNow:yyyy-MM-ddTHH:mm:ss.fffZ} {redacted}{Environment.NewLine}";
            var bytes = System.Text.Encoding.UTF8.GetByteCount(line);

            RotateIfNeeded(bytes);
            _currentWriter?.Write(line);
            _currentWriter?.Flush();
            _currentFileSize += bytes;
        }
        finally
        {
            _gate.Exit();
        }
    }

    private void RotateIfNeeded(long incomingBytes = 0)
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
            if (File.Exists(activeLogPath))
            {
                var newPath = Path.Combine(_logDirectory, $"{_baseName}.1.log");
                File.Move(activeLogPath, newPath);
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

    private static void DeleteFile(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch (IOException)
        {
            throw;
        }
        catch (UnauthorizedAccessException)
        {
            throw;
        }
    }

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
