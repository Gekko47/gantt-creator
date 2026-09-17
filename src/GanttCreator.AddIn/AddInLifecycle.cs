using GanttCreator.Core.Logging;

namespace GanttCreator.AddIn;

/// <summary>
/// Writes the <c>AutoOpen</c>/<c>AutoClose</c> lifecycle records to the
/// add-in rolling log. Records carry identifiers only — never workbook
/// content — and every field missing a value is normalized to
/// <c>"unknown"</c> so the record shape is stable.
/// </summary>
/// <param name="log">The rolling log to write to.</param>
/// <exception cref="ArgumentNullException"><paramref name="log"/> is <see langword="null"/>.</exception>
public sealed class AddInLifecycle(IRollingLog log)
{
    private readonly IRollingLog _log = log ?? throw new ArgumentNullException(nameof(log));

    /// <summary>
    /// Writes exactly one open record carrying the identity fields plus the
    /// per-load session token (work item R1.6 D5).
    /// </summary>
    /// <param name="identity">Load-session identifiers.</param>
    /// <exception cref="ArgumentNullException"><paramref name="identity"/> is <see langword="null"/>.</exception>
    public void LogOpen(AddInIdentity identity)
    {
        ArgumentNullException.ThrowIfNull(identity);
        _log.Write(
            "open addin-version={0} excel-version={1} process={2} xll={3} session={4}",
            OrUnknown(identity.AddInVersion),
            OrUnknown(identity.ExcelVersion),
            OrUnknown(identity.ProcessBitness),
            OrUnknown(identity.XllFileName),
            OrUnknown(identity.SessionToken));
    }

    /// <summary>
    /// Writes exactly one close record carrying the session token of the
    /// load it closes (work item R1.6 D5), so the <c>open</c>/<c>close</c>
    /// pair of one load is correlatable in the shared log.
    /// </summary>
    /// <param name="sessionToken">The session token of the load being closed.</param>
    public void LogClose(string? sessionToken) => _log.Write("close session={0}", OrUnknown(sessionToken));

    private static string OrUnknown(string? value) =>
        string.IsNullOrWhiteSpace(value) ? "unknown" : value;
}
