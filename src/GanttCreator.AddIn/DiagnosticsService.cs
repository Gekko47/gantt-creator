using System.Diagnostics;
using System.Reflection;
using ExcelDna.Integration;
using GanttCreator.Core.Logging;

// Expose internals to the test project so contract tests can verify the
// log-writing path without requiring the TaskDialog UI.
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("GanttCreator.AddIn.Tests")]

namespace GanttCreator.AddIn;

/// <summary>
/// Project-wide diagnostics service. Gathers add-in identifiers, writes a
/// diagnostic record to the rolling log, and shows a Windows TaskDialog with
/// a clickable hyperlink to the active log file. Intended to be the single
/// diagnostics entry point for the add-in.
/// </summary>
/// <remarks>
/// The service is a singleton per Excel session. <see cref="Instance"/> is
/// set once at first use and reset by <see cref="Reset"/>. The log reference
/// is injected via <see cref="SetLog"/> so the service does not own the log
/// lifetime (the log is owned by <see cref="AddInHost"/>).
/// </remarks>
public sealed class DiagnosticsService
{
    private static DiagnosticsService? _instance;

    /// <summary>
    /// The singleton instance for the current Excel session. Returns null
    /// until <see cref="SetLog"/> is called.
    /// </summary>
    public static DiagnosticsService Instance
    {
        get
        {
            if (_instance is null)
            {
                _instance = new DiagnosticsService();
            }
            return _instance;
        }
    }

    private IRollingLog? _log;

    private DiagnosticsService()
    {
    }

    /// <summary>
    /// Injects the rolling log reference. Called once per Excel session by
    /// <see cref="AddInHost"/> after the log is created.
    /// </summary>
    /// <param name="log">The rolling log.</param>
    public void SetLog(IRollingLog log)
    {
        _log = log ?? throw new ArgumentNullException(nameof(log));
    }

    /// <summary>
    /// Resets the singleton and clears the log reference, so the next Excel
    /// session starts fresh. Called by <see cref="AddInHost.AutoClose"/>.
    /// </summary>
    public static void Reset()
    {
        _instance?._log = null;
        _instance = null;
    }

    /// <summary>
    /// Returns the full path to the active log file, or null when no log has
    /// been injected or the log has failed/disposed.
    /// </summary>
    public string? LogFilePath => _log?.LogFilePath;

    /// <summary>
    /// Shows the diagnostics TaskDialog with identifier summary and a
    /// clickable hyperlink to the active log file. Writes a diagnostic
    /// record to the log before showing the dialog.
    /// </summary>
    public void ShowDiagnostics()
    {
        // Gather identifiers for display and logging.
        var identifiers = GatherIdentifiers();

        // Write a diagnostic record to the log (non-fatal: if the log has
        // failed, this is a no-op via RollingLog's own guards).
        WriteDiagnosticRecord(identifiers);

        // Build the hyperlink-enabled content. The hyperlink is a file:// URI
        // so TaskDialog parses it as a clickable link.
        string content = BuildContent(identifiers);

        // Show the dialog.
        int buttonId = TaskDialogApi.ShowWithHyperlink(
            title: "Gantt Creator Diagnostics",
            mainInstruction: "Add-in diagnostic information",
            content: content,
            openFileAction: path => OpenLogFile(path));

        // Log which button was clicked (informational; non-fatal if log failed).
        if (_log is not null)
        {
            _log.Write("Diagnostics dialog closed; button ID = {0}.", buttonId);
        }
    }

    /// <summary>
    /// Gathers the four add-in identifiers used for diagnostics.
    /// </summary>
    /// <returns>The identifier set.</returns>
    public AddInIdentity GatherIdentifiers()
    {
        var excelVersion = "unknown";
        var xllFileName = "unknown";
        var xllPath = "unknown";

        try
        {
            excelVersion = ExcelDnaUtil.ExcelVersion.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture);
        }
        catch
        {
            // Intentionally empty: identifier degrades to "unknown".
        }

        try
        {
            xllPath = ExcelDnaUtil.XllPath;
            xllFileName = Path.GetFileName(xllPath);
            if (string.IsNullOrWhiteSpace(xllFileName))
            {
                xllFileName = "unknown";
            }
        }
        catch
        {
            // Intentionally empty: identifier degrades to "unknown".
        }

        return new AddInIdentity(
            VersionInfo.SemanticVersion,
            excelVersion,
            Environment.Is64BitProcess ? "x64" : "x86",
            xllFileName);
    }

    /// <summary>
    /// Builds the TaskDialog content with the identifier summary and a
    /// clickable hyperlink to the active log file.
    /// </summary>
    /// <param name="identifiers">The add-in identifiers.</param>
    /// <returns>The content string with a file:// hyperlink.</returns>
    public string BuildContent(AddInIdentity identifiers, string? logFilePath = null)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Add-in Version: " + identifiers.AddInVersion);
        sb.AppendLine("Excel Version:  " + identifiers.ExcelVersion);
        sb.AppendLine("Process:        " + identifiers.ProcessBitness);
        sb.AppendLine("XLL File:       " + identifiers.XllFileName);

        if (!string.IsNullOrWhiteSpace(logFilePath))
        {
            sb.AppendLine();
            // Use a file:// URI so TaskDialog parses it as a hyperlink.
            // Normalize to use forward slashes and percent-encode spaces.
            string uri = new Uri(logFilePath).AbsoluteUri;
            sb.AppendLine("Log File: " + uri);
        }
        else
        {
            sb.AppendLine();
            sb.AppendLine("Log File: not available");
        }

        return sb.ToString();
    }

    /// <summary>
    /// Opens the log file using the system default handler (typically the
    /// associated text editor or the Windows file preview). Non-fatal: a
    /// failure to open the file does not surface to the user.
    /// </summary>
    /// <param name="path">The log file path.</param>
    public void OpenLogFile(string path)
    {
        try
        {
            Process.Start(new ProcessStartInfo(path)
            {
                UseShellExecute = true
            });
        }
        catch
        {
            // Intentionally empty: non-fatal.
        }
    }

    /// <summary>
    /// Writes a diagnostic record to the log, using the same format as
    /// <see cref="ShowDiagnostics"/>. Public so contract tests can verify
    /// the log-writing path without invoking the TaskDialog UI.
    /// </summary>
    public void WriteDiagnosticRecord(AddInIdentity identifiers)
    {
        if (_log is null)
        {
            return;
        }

        _log.Write(
            "Diagnostics: addin-version={0} excel-version={1} process={2} xll={3}",
            identifiers.AddInVersion,
            identifiers.ExcelVersion,
            identifiers.ProcessBitness,
            identifiers.XllFileName);
    }
}
