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

    // The leading lookbehind requires the '/' to begin a token: it must not
    // be preceded by a word character or another slash. Without it, prose
    // like "and/or", ratios like "3/4", and dates like "12/31/2026" were
    // redacted as [path], destroying log usefulness. Absolute Unix paths in
    // real messages follow whitespace, punctuation, or the start of the text.
    [GeneratedRegex(@"(?<![\w/])/(?:[^\s/\\]+/)*[A-Za-z0-9._-]+(?:\.[A-Za-z0-9]+)?", RegexOptions.Compiled | RegexOptions.CultureInvariant)]
    private static partial Regex UnixPathPattern();

    [GeneratedRegex(@"\b\d{4}-\d{2}-\d{2}(?:[T\s]\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?)?\b", RegexOptions.Compiled | RegexOptions.CultureInvariant)]
    private static partial Regex IsoDatePattern();

    [GeneratedRegex(@"\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b", RegexOptions.Compiled | RegexOptions.CultureInvariant)]
    private static partial Regex GuidPattern();

    [GeneratedRegex(@"\b[0-9a-fA-F]{16,}\b", RegexOptions.Compiled | RegexOptions.CultureInvariant)]
    private static partial Regex LongHexTokenPattern();

    /// <summary>
    /// Redacts known sensitive patterns (emails, paths, dates, GUIDs, hex tokens)
    /// in the input text, replacing them with stable placeholder tokens.
    /// </summary>
    /// <param name="input">The raw text to redact. May be <c>null</c>, in which case <c>null</c> is returned.</param>
    /// <returns>The redacted text with sensitive patterns replaced, or <c>null</c> when <paramref name="input"/> is <c>null</c>.</returns>
    [return: System.Diagnostics.CodeAnalysis.NotNullIfNotNull(nameof(input))]
    public string? Redact(string? input)
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
