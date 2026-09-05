namespace GanttCreator.Core.Logging;

/// <summary>
/// Redacts privacy-sensitive data from log messages before they are persisted.
/// </summary>
public interface IRedactor
{
    /// <summary>
    /// Redacts known sensitive patterns in the input text.
    /// </summary>
    /// <param name="input">The raw log message.</param>
    /// <returns>The redacted message.</returns>
    string Redact(string input);
}