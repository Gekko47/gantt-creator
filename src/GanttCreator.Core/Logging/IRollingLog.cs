namespace GanttCreator.Core.Logging;

/// <summary>
/// Rolling log abstraction: size-based rotation with configurable file count cap.
/// </summary>
public interface IRollingLog : IDisposable
{
    /// <summary>
    /// Indicates the logger has latched a failure (redaction,
    /// rotation, or write) and will discard all further writes. The latch is
    /// permanent by design: log failures must never propagate into product
    /// code paths. Diagnostics should consult this property to detect that
    /// logging has stopped.
    /// </summary>
    bool IsFailed { get; }

    /// <summary>
    /// The full path to the active log file, or null when the log has not been
    /// initialized or has failed during initialization. This is an informational
    /// property for diagnostics; it does not affect write behaviour.
    /// </summary>
    string? LogFilePath { get; }

    /// <summary>
    /// Writes a message to the log. The message is automatically redacted.
    /// </summary>
    /// <param name="message">The log message.</param>
    void Write(string message);

    /// <summary>
    /// Writes a formatted message to the log.
    /// A malformed format string skips only that message without latching
    /// <see cref="IsFailed"/>; subsequent writes still proceed.
    /// </summary>
    /// <param name="format">Composite format string.</param>
    /// <param name="args">Format arguments.</param>
    void Write(string format, params object?[] args);
}
