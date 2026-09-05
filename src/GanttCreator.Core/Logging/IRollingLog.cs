namespace GanttCreator.Core.Logging;

/// <summary>
/// Rolling log abstraction: size-based rotation with configurable file count cap.
/// </summary>
public interface IRollingLog : IDisposable
{
    /// <summary>
    /// Writes a message to the log. The message is automatically redacted.
    /// </summary>
    /// <param name="message">The log message.</param>
    void Write(string message);

    /// <summary>
    /// Writes a formatted message to the log.
    /// </summary>
    /// <param name="format">Composite format string.</param>
    /// <param name="args">Format arguments.</param>
    void Write(string format, params object?[] args);
}