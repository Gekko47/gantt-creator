namespace GanttCreator.Core.Logging;

/// <summary>
/// Redacts privacy-sensitive data from log messages before they are persisted.
/// </summary>
public interface IRedactor
{
    /// <summary>
    /// Redacts known sensitive patterns in the input text.
    /// </summary>
    /// <param name="input">The raw log message. May be <c>null</c>, in which case <c>null</c> is returned.</param>
    /// <returns>The redacted message, or <c>null</c> when <paramref name="input"/> is <c>null</c>.</returns>
    [return: System.Diagnostics.CodeAnalysis.NotNullIfNotNull(nameof(input))]
    string? Redact(string? input);
}
