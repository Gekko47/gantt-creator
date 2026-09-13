using System.Diagnostics;
using System.Text;
using ExcelDna.Integration;
using GanttCreator.Core;
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
    /// The singleton instance for the current Excel session. Creates and
    /// returns a <see cref="DiagnosticsService"/> on first access; the
    /// injected log may be absent until <see cref="SetLog"/> is called.
    /// </summary>
    public static DiagnosticsService Instance
    {
        get
        {
            _instance ??= new DiagnosticsService();
            return _instance;
        }
    }

    private IRollingLog? _log;

    private DiagnosticsService() { }

    /// <summary>
    /// Injects the rolling log reference. Called once per Excel session by
    /// <see cref="AddInHost"/> after the log is created.
    /// </summary>
    /// <param name="log">The rolling log.</param>
    public void SetLog(IRollingLog log) => _log = log ?? throw new ArgumentNullException(nameof(log));

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
        const string Title = "Gantt Creator Diagnostics";
        const string MainInstruction = "Add-in diagnostic information";

        // Gather identifiers for display and logging.
        AddInIdentity identifiers = GatherIdentifiers();

        // Write a diagnostic record to the log (non-fatal: if the log has
        // failed, this is a no-op via RollingLog's own guards).
        WriteDiagnosticRecord(identifiers);

        // Build the hyperlink-enabled content. The hyperlink is a file:// URI
        // anchor (<a href="...">) so the TaskDialog renders it as a link.
        var content = BuildContent(identifiers, LogFilePath);

        try
        {
            // Show the dialog.
            var buttonId = TaskDialogApi.ShowWithHyperlink(
                title: Title,
                mainInstruction: MainInstruction,
                content: content,
                openFileAction: OpenLogFile
            );

            // Log which button was clicked (informational; non-fatal if log failed).
            _log?.Write("Diagnostics dialog closed; button ID = {0}.", buttonId);
        }
#pragma warning disable CA1031
        catch (Exception ex)
        {
            // CA1031: The diagnostics command must always show the user
            // *some* dialog. The managed TaskDialog can fail outside the
            // Excel main STA thread or on a malformed page configuration, so
            // any failure falls back to a plain MessageBox carrying the same
            // information, degrading visibility instead of hiding it.
            try
            {
                _log?.Write("Task dialog failed; falling back to MessageBox: {0}", ex.Message);
            }
            catch
            {
                // Intentionally empty: the fallback dialog must still show.
            }
            ShowFallbackDialog(Title, BuildFallbackText(identifiers, LogFilePath));
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Fallback dialog used when the TaskDialog cannot be shown. Displays the
    /// dedicated plain text built by <see cref="BuildFallbackText"/> (the four
    /// identifiers plus the log file path).
    /// </summary>
    /// <param name="title">The dialog title.</param>
    /// <param name="text">The plain fallback text.</param>
    private static void ShowFallbackDialog(string title, string text)
    {
        _ = MessageBox.Show(
            text,
            title,
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    /// <summary>
    /// Doubles each <c>&amp;</c> so the string survives the TaskDialog's
    /// access-key (mnemonic) interpretation, which applies when
    /// <see cref="TaskDialogPage.EnableLinks"/> is on and at least one link
    /// is present. Internal so the escaping contract is testable.
    /// </summary>
    /// <param name="text">Plain-text fragment to escape.</param>
    /// <returns>The escaped fragment.</returns>
    internal static string EscapeAccessKeys(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        return text.Replace("&", "&&", StringComparison.Ordinal);
    }

    /// <summary>
    /// Gathers the four add-in identifiers used for diagnostics.
    /// </summary>
    /// <returns>The identifier set.</returns>
    public static AddInIdentity GatherIdentifiers()
    {
        var excelVersion = "unknown";
        var xllFileName = "unknown";

#pragma warning disable CA1031
        try
        {
            excelVersion = ExcelDnaUtil.ExcelVersion.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture);
        }
        catch
#pragma warning restore CA1031
        {
            // Intentionally empty: identifier degrades to "unknown".
        }

#pragma warning disable CA1031
        try
        {
            var xllPath = ExcelDnaUtil.XllPath;
            xllFileName = Path.GetFileName(xllPath);
            if (string.IsNullOrWhiteSpace(xllFileName))
            {
                xllFileName = "unknown";
            }
        }
        catch
#pragma warning restore CA1031
        {
            // Intentionally empty: identifier degrades to "unknown".
        }

        return new AddInIdentity(VersionInfo.SemanticVersion, excelVersion, Environment.Is64BitProcess ? "x64" : "x86", xllFileName);
    }

    /// <summary>
    /// Builds the TaskDialog content with the identifier summary and a
    /// clickable hyperlink to the active log file.
    /// </summary>
    /// <param name="identifiers">The add-in identifiers.</param>
    /// <param name="logFilePath">The active log file path, or null when unavailable.</param>
    /// <returns>The content string with a file:// hyperlink.</returns>
    public static string BuildContent(AddInIdentity identifiers, string? logFilePath = null)
    {
        ArgumentNullException.ThrowIfNull(identifiers);

        StringBuilder sb = new();
        // EnableLinks turns '&' into an access-key (mnemonic) prefix once at
        // least one link is present, so escape it in the plain-text lines.
        AppendIdentifierLines(sb, identifiers, escapeAccessKeys: true);

        if (!string.IsNullOrWhiteSpace(logFilePath))
        {
            _ = sb.AppendLine();
            // Render the file:// URI as an anchor so the TaskDialog shows it
            // as a clickable hyperlink (requires TaskDialogPage.EnableLinks).
            var uri = new Uri(logFilePath).AbsoluteUri;
            _ = sb.AppendLine(
                System.Globalization.CultureInfo.InvariantCulture,
                $"Log File: <a href=\"{uri}\">Open the log file</a>");
        }
        else
        {
            _ = sb.AppendLine();
            _ = sb.AppendLine("Log File: not available");
        }

        return sb.ToString();
    }

    /// <summary>
    /// Builds the plain-text content for the MessageBox fallback: the four
    /// identifier lines plus the log file path as dedicated plain text, with
    /// no <c>&lt;a href&gt;</c> markup (a MessageBox cannot render links, so
    /// the actual path is shown). Internal so the fallback contract is
    /// testable.
    /// </summary>
    /// <param name="identifiers">The add-in identifiers.</param>
    /// <param name="logFilePath">The active log file path, or null when unavailable.</param>
    /// <returns>The plain-text fallback content.</returns>
    internal static string BuildFallbackText(AddInIdentity identifiers, string? logFilePath = null)
    {
        ArgumentNullException.ThrowIfNull(identifiers);

        StringBuilder sb = new();
        // The MessageBox has no access-key interpretation, so the identifiers
        // are written unescaped.
        AppendIdentifierLines(sb, identifiers, escapeAccessKeys: false);

        if (!string.IsNullOrWhiteSpace(logFilePath))
        {
            _ = sb.AppendLine();
            _ = sb.AppendLine("Log File: " + logFilePath);
        }
        else
        {
            _ = sb.AppendLine();
            _ = sb.AppendLine("Log File: not available");
        }

        return sb.ToString();
    }

    /// <summary>
    /// Appends the four identifier lines. <paramref name="escapeAccessKeys"/>
    /// controls whether <c>&amp;</c> is doubled for the TaskDialog's access-key
    /// interpretation (required on the hyperlink-enabled content, unnecessary
    /// for the plain MessageBox fallback).
    /// </summary>
    /// <param name="sb">The builder to append to.</param>
    /// <param name="identifiers">The add-in identifiers.</param>
    /// <param name="escapeAccessKeys">
    /// True to double <c>&amp;</c> (hyperlink-enabled TaskDialog content);
    /// false for plain text (MessageBox fallback).
    /// </param>
    private static void AppendIdentifierLines(StringBuilder sb, AddInIdentity identifiers, bool escapeAccessKeys)
    {
        AddLine("Add-in Version: ", identifiers.AddInVersion);
        AddLine("Excel Version:  ", identifiers.ExcelVersion);
        AddLine("Process:        ", identifiers.ProcessBitness);
        AddLine("XLL File:       ", identifiers.XllFileName);

        void AddLine(string label, string value)
        {
            var text = escapeAccessKeys ? EscapeAccessKeys(value) : value;
            _ = sb.AppendLine(label + text);
        }
    }

    /// <summary>
    /// Opens the log file using the system default handler (typically the
    /// associated text editor or the Windows file preview). Non-fatal: a
    /// failure to open the file does not surface to the user.
    /// </summary>
    /// <param name="path">The log file path.</param>
    public static void OpenLogFile(string path)
    {
#pragma warning disable CA1031
        try
        {
            _ = Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch
#pragma warning restore CA1031
        {
            // Intentionally empty: non-fatal.
        }
    }

    /// <summary>
    /// Records a diagnostics callback failure through the rolling log (if
    /// available) for diagnosis. Never throws: a log write failure degrades
    /// to no record, which is the safe direction inside an exception handler.
    /// </summary>
    /// <param name="ex">The exception that was caught.</param>
    public static void WriteDiagnosticsError(Exception ex)
    {
        ArgumentNullException.ThrowIfNull(ex);

        // CA1031: This runs inside a ribbon callback that is already handling
        // an exception; recording must never introduce a new failure path.
#pragma warning disable CA1031
        try
        {
            Instance._log?.Write("OnDiagnosticsClick failed: {0}", ex.ToString());
        }
        catch
        {
            // Intentionally empty: the record degrades to nothing.
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Writes a diagnostic record to the log, using the same format as
    /// <see cref="ShowDiagnostics"/>. Public so contract tests can verify
    /// the log-writing path without invoking the TaskDialog UI.
    /// </summary>
    /// <param name="identifiers">The add-in identifiers.</param>
    public void WriteDiagnosticRecord(AddInIdentity identifiers)
    {
        ArgumentNullException.ThrowIfNull(identifiers);

        _log?.Write(
            "Diagnostics: addin-version={0} excel-version={1} process={2} xll={3}",
            identifiers.AddInVersion,
            identifiers.ExcelVersion,
            identifiers.ProcessBitness,
            identifiers.XllFileName
        );
    }
}
