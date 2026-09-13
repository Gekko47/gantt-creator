using System.Runtime.InteropServices;

namespace GanttCreator.AddIn;
/// <summary>
/// Pure, UI-free translation of an unexpected exception into a concise,
/// user-safe message that contains the operation ID. This is the "translated
/// result" half of the command error boundary
/// (docs/03-ROADMAP.md R1.4). The translation never includes the exception
/// text, the stack, or machine paths — those belong only in the redacted
/// technical log record.
/// </summary>
internal static class CommandErrorTranslator
{
    /// <summary>
    /// Translates <paramref name="exception"/> into a user-safe message.
    /// Known exception categories map to a specific concise sentence; every
    /// other exception falls back to the generic sentence. The result always
    /// contains the operation ID and a pointer to the log file.
    /// </summary>
    /// <param name="exception">The caught exception.</param>
    /// <param name="operationId">The operation ID to reference in the message.</param>
    /// <returns>The user-safe message text.</returns>
    /// <exception cref="ArgumentNullException">An argument is <see langword="null"/>.</exception>
    internal static string Translate(Exception exception, string operationId)
    {
        ArgumentNullException.ThrowIfNull(exception);
        ArgumentNullException.ThrowIfNull(operationId);

        var category = Categorize(exception);
        return string.Create(
            System.Globalization.CultureInfo.InvariantCulture,
            $"""
            {category}
            Operation: {operationId}

            The action was cancelled and your workbook is unchanged.
            Technical details were written to the log file (see the Diagnostics button).
            """);
    }

    /// <summary>
    /// Maps the exception category to its concise user-safe sentence.
    /// Internal so the category mapping contract is testable in isolation.
    /// </summary>
    /// <param name="exception">The caught exception.</param>
    /// <returns>The user-safe category sentence.</returns>
    internal static string Categorize(Exception exception)
    {
        ArgumentNullException.ThrowIfNull(exception);

        // COM interop failures (the Excel/PowerPoint boundary rejects a call,
        // returns a failure HRESULT, or the object is disconnected). SEH is
        // included because some COM failures surface as structured errors.
        if (exception is COMException or SEHException)
        {
            return "Excel could not complete the operation.";
        }

        if (exception is UnauthorizedAccessException or System.Security.SecurityException)
        {
            return "Gantt Creator does not have permission to complete the action.";
        }

        if (exception is IOException)
        {
            return "A file needed by this action was not available.";
        }

        // IDE0046 is intentionally not applied to the earlier guards: each
        // category is a separate, documented branch. Only the final branch is
        // a binary choice between two returns.
        return exception is OperationCanceledException
            ? "The action was cancelled."
            : "An unexpected error occurred.";
    }
}
