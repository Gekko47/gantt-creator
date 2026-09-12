using GanttCreator.Core.Logging;

namespace GanttCreator.AddIn;

/// <summary>
/// Creates the add-in rolling log under the per-user local application
/// data directory, with a temp-directory fallback. The location is offline,
/// user-local, and independent of any workbook (the log must never be
/// written to a worksheet). Write failures latch
/// <see cref="IRollingLog.IsFailed"/> inside Core's <see cref="RollingLog"/>
/// and never propagate.
/// </summary>
public static class AddInLogFactory
{
    /// <summary>The product directory written under the resolved base directory.</summary>
    private const string _productDirectory = "GanttCreator";

    /// <summary>The log sub-directory written under the product directory.</summary>
    private const string _logSubdirectory = "logs";

    /// <summary>
    /// Log file base name: the active file is <c>gantt-creator-addin.log</c>
    /// and rotations are <c>gantt-creator-addin.N.log</c>.
    /// </summary>
    private const string _baseName = "gantt-creator-addin";

    /// <summary>
    /// Creates the add-in rolling log under
    /// <c>%LOCALAPPDATA%\GanttCreator\logs</c>. When the
    /// <c>%LOCALAPPDATA%</c> path is unavailable, the system temp directory
    /// is used instead.
    /// </summary>
    /// <returns>A rolling log ready for writes; write failures latch <see cref="IRollingLog.IsFailed"/> and never throw.</returns>
    public static IRollingLog Create()
    {
        var baseDirectory = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        IRollingLog log = Create(baseDirectory);
        if (!string.IsNullOrWhiteSpace(baseDirectory) && log.IsFailed)
        {
            log.Dispose();
            log = Create(Path.GetTempPath());
        }
        return log;
    }

    /// <summary>
    /// Creates the add-in rolling log under the given base directory.
    /// A <see langword="null"/> or whitespace base directory selects the
    /// system temp fallback.
    /// </summary>
    /// <param name="baseDirectory">Base directory for the log tree; <see langword="null"/> or whitespace selects the temp fallback.</param>
    /// <returns>A rolling log ready for writes; write failures latch <see cref="IRollingLog.IsFailed"/> and never throw.</returns>
    public static IRollingLog Create(string? baseDirectory)
    {
        var root = string.IsNullOrWhiteSpace(baseDirectory)
            ? Path.GetTempPath()
            : baseDirectory;
        return new RollingLog(Path.Combine(root, _productDirectory, _logSubdirectory), baseName: _baseName);
    }
}
