using System.Text.RegularExpressions;

namespace GanttCreator.Core.Logging;

/// <summary>
/// Default redaction implementation. Replaces emails, absolute file paths,
/// ISO dates, GUIDs, and long hex tokens with stable placeholder tokens.
/// </summary>
public sealed partial class Redactor : IRedactor
{
    [GeneratedRegex(@"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b", RegexOptions.Compiled | RegexOptions.CultureInvariant)]
    private static partial Regex EmailPattern();

    [GeneratedRegex(@"(?:\\\\[^\s]+|[A-Za-z]:[\\/][^\s]+)", RegexOptions.Compiled | RegexOptions.CultureInvariant)]
    private static partial Regex WindowsPathPattern();

    [GeneratedRegex(@"/(?:[^\s/\\]+/)*[A-Za-z0-9._-]+(?:\.[A-Za-z0-9]+)?", RegexOptions.Compiled | RegexOptions.CultureInvariant)]
    private static partial Regex UnixPathPattern();

    [GeneratedRegex(@"\b\d{4}-\d{2}-\d{2}[T\s]\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?\b", RegexOptions.Compiled | RegexOptions.CultureInvariant)]
    private static partial Regex IsoDatePattern();

    [GeneratedRegex(@"\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b", RegexOptions.Compiled | RegexOptions.CultureInvariant)]
    private static partial Regex GuidPattern();

    [GeneratedRegex(@"\b[0-9a-fA-F]{16,}\b", RegexOptions.Compiled | RegexOptions.CultureInvariant)]
    private static partial Regex LongHexTokenPattern();

    /// <summary>
    /// Redacts known sensitive patterns (emails, paths, dates, GUIDs, hex tokens)
    /// in the input text, replacing them with stable placeholder tokens.
    /// </summary>
    /// <param name="input">The raw text to redact.</param>
    /// <returns>The redacted text with sensitive patterns replaced.</returns>
    public string Redact(string input)
    {
        if (string.IsNullOrEmpty(input))
        {
            return input;
        }

        var result = input;

        result = EmailPattern().Replace(result, "[email]");
        result = WindowsPathPattern().Replace(result, "[path]");
        result = UnixPathPattern().Replace(result, "[path]");
        result = IsoDatePattern().Replace(result, "[datetime]");
        result = GuidPattern().Replace(result, "[guid]");
        result = LongHexTokenPattern().Replace(result, "[token]");

        return result;
    }
}
