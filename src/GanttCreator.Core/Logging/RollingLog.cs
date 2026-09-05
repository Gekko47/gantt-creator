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

        _logDirectory = logDirectory;
        _baseName = string.IsNullOrWhiteSpace(baseName) ? "gantt-creator" : baseName;
        _maxFileSizeBytes = maxFileSizeBytes;
        _maxFileCount = maxFileCount;
        _redactor = redactor ?? new Redactor();

        _ = Directory.CreateDirectory(_logDirectory);
        RotateIfNeeded();
    }

    /// <summary>
    /// Writes a message to the log. The message is automatically redacted.
    /// </summary>
    /// <param name="message">The log message.</param>
    public void Write(string message)
    {
        WriteCore(message);
    }

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

        var redacted = _redactor.Redact(message);
        var line = $"{DateTime.UtcNow:yyyy-MM-ddTHH:mm:ss.fffZ} {redacted}{Environment.NewLine}";
        var bytes = System.Text.Encoding.UTF8.GetByteCount(line);

        _gate.Enter();
        try
        {
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

            // Rotate existing files: delete oldest if at cap, shift others
            var files = Directory.GetFiles(_logDirectory, $"{_baseName}*.log")
                .OrderByDescending(f => f)
                .ToArray();

            if (files.Length >= _maxFileCount)
            {
                for (int i = _maxFileCount - 1; i < files.Length; i++)
                {
                    try
                    {
                        File.Delete(files[i]);
                    }
                    catch (IOException)
                    {
                        // best effort
                    }
                    catch (UnauthorizedAccessException)
                    {
                        // best effort
                    }
                }
            }

            for (int i = files.Length - 1; i >= 0; i--)
            {
                var newName = $"{_baseName}.{files.Length - i}.log";
                var newPath = Path.Combine(_logDirectory, newName);
                try
                {
                    File.Move(files[i], newPath, overwrite: true);
                }
                catch (IOException)
                {
                    // best effort
                }
                catch (UnauthorizedAccessException)
                {
                    // best effort
                }
            }

            _currentFilePath = Path.Combine(_logDirectory, $"{_baseName}.log");
            var fileStream = new FileStream(_currentFilePath, FileMode.Create, FileAccess.Write, FileShare.Read);
            _currentWriter = new StreamWriter(fileStream, System.Text.Encoding.UTF8);
            _currentFileSize = 0;
        }
    }

    /// <summary>
    /// Disposes the rolling log, releasing the current writer.
    /// </summary>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }
        _disposed = true;
        _currentWriter?.Dispose();
        _currentWriter = null;
    }
}