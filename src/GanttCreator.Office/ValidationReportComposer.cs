using System.Globalization;
using System.Text;
using GanttCreator.Core;

namespace GanttCreator.Office;

/// <summary>
/// Pure helper that groups an already-sorted issue list into the per-cell note
/// payloads the reporter writes. Contains no COM: all grouping, ordering, and
/// text composition are unit-testable in isolation.
/// </summary>
/// <remarks>
/// <para>
/// The reporter contract expects issues in deterministic
/// <c>Severity → Row → Field → Code</c> order (see
/// <see cref="GanttValidationOutcome"/>). This composer assumes that order and
/// only groups consecutive issues on the same <c>(RowNumber, Field)</c> into a
/// single cell note whose lines keep that order.
/// </para>
/// <para>
/// Ownership is tagged by a leading text sentinel
/// (<see cref="NoteSentinelPrefix"/>) because Excel's <c>Comment.Author</c> is
/// read-only; the reporter recognises owned notes by this prefix rather than by
/// author.
/// </para>
/// </remarks>
internal static class ValidationReportComposer
{
    /// <summary>
    /// The text prefix every add-in-owned validation note begins with, used to
    /// recognise and clear only notes this add-in wrote. The reader owns notes
    /// it created (renderer rule E analog for cells).
    /// </summary>
    internal const string NoteSentinelPrefix = "Gantt Creator validation:";

    /// <summary>
    /// The maximum characters of a cell note body Excel will accept
    /// (<c>Range.AddComment</c> / <c>Comment.Text</c>). The composed note must
    /// not exceed this; the reporter asserts it here so a long issue set is
    /// never silently truncated by Excel.
    /// </summary>
    internal const int MaxNoteLength = 255;

    /// <summary>
    /// Groups consecutive same-cell issues into note payloads, in the supplied
    /// (already sorted) order.
    /// </summary>
    /// <param name="issues">
    /// Findings in <c>Severity → Row → Field → Code</c> order. Need not be
    /// pre-grouped.
    /// </param>
    /// <returns>
    /// Ordered note payloads, one per distinct <c>(RowNumber, Field)</c> cell,
    /// in first-appearance order. Each payload's <see cref="NotePayload.Text"/>
    /// begins with <see cref="NoteSentinelPrefix"/> followed by one
    /// <c>Severity: Message</c> line per issue.
    /// </returns>
    /// <exception cref="ArgumentNullException"><paramref name="issues"/> is <see lang="null"/>.</exception>
    public static IReadOnlyList<NotePayload> GroupIntoNotes(IReadOnlyList<GanttValidationIssue> issues)
    {
        ArgumentNullException.ThrowIfNull(issues);

        var result = new List<NotePayload>(issues.Count);
        var index = 0;
        while (index < issues.Count)
        {
            GanttValidationIssue current = issues[index];
            var lines = new List<string>(capacity: 4);

            while (index < issues.Count
                   && issues[index].RowNumber == current.RowNumber
                   && string.Equals(issues[index].Field, current.Field, StringComparison.Ordinal))
            {
                lines.Add($"{SeverityLabel(issues[index].Severity)}: {issues[index].Message}");
                index++;
            }

            var text = BuildNoteText(current.RowNumber, current.Field, lines);
            if (text.Length > MaxNoteLength)
            {
                text = text[..MaxNoteLength];
            }

            result.Add(new NotePayload(current.RowNumber, current.Field, text));
        }

        return result;
    }

    /// <summary>
    /// Returns <c>true</c> when the note text belongs to this add-in (starts
    /// with <see cref="NoteSentinelPrefix"/>).
    /// </summary>
    /// <param name="text">The note text, or <see lang="null"/>.</param>
    /// <returns>
    /// <see langword="true"/> when the text is non-empty and begins with the
    /// sentinel prefix.
    /// </returns>
    public static bool IsOwnedByAddIn(string? text) =>
        !string.IsNullOrEmpty(text)
        && text.StartsWith(NoteSentinelPrefix, StringComparison.Ordinal);

    /// <summary>
    /// Returns the part of a note text that the user wrote: everything before
    /// the first <see cref="NoteSentinelPrefix"/> occurrence, with trailing
    /// whitespace trimmed. Text with no sentinel is returned unchanged (trimmed).
    /// </summary>
    /// <remarks>
    /// Excel allows only one classic note per cell and <c>Range.AddComment</c>
    /// fails when a note already exists (MS Learn, <c>Range.AddComment</c>).
    /// The reporter therefore preserves the user's text verbatim as a prefix and
    /// replaces only the add-in section, using this method to strip a previous
    /// section before appending the current one.
    /// </remarks>
    /// <param name="text">The current note text, or <see langword="null"/>.</param>
    /// <returns>The user-authored prefix; empty when the note is purely the add-in's.</returns>
    public static string StripOwnedSection(string? text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }

        var marker = text.IndexOf(NoteSentinelPrefix, StringComparison.Ordinal);
        var userPart = marker < 0 ? text : text[..marker];
        return userPart.TrimEnd('\r', '\n', ' ');
    }

    /// <summary>
    /// Builds the full note text for one cell: the sentinel header line, a
    /// column hint, then one message line per issue, in order.
    /// </summary>
    /// <param name="rowNumber">The one-based body row the note anchors on.</param>
    /// <param name="fieldName">The column the note anchors on.</param>
    /// <param name="messageLines">The ordered <c>Severity: Message</c> lines.</param>
    /// <returns>The composed note text.</returns>
    private static string BuildNoteText(int rowNumber, string fieldName, List<string> messageLines)
    {
        var sb = new StringBuilder();
        _ = sb.Append(NoteSentinelPrefix).Append(' ');
        _ = sb.AppendFormat(CultureInfo.InvariantCulture, "row {0}, column '{1}':", rowNumber, fieldName);
        foreach (var line in messageLines)
        {
            _ = sb.Append(' ').Append(line);
        }

        return sb.ToString();
    }

    /// <summary>
    /// Maps a severity to its note-line label.
    /// </summary>
    private static string SeverityLabel(GanttValidationSeverity severity) => severity switch
    {
        GanttValidationSeverity.Error => "Error",
        GanttValidationSeverity.Warning => "Warning",
        _ => "Issue",
    };
}

/// <summary>
/// One cell note to write: its body row, the column it should anchor on, and the
/// full composed note text (already including the ownership sentinel).
/// </summary>
/// <param name="RowNumber">The one-based body-row index to anchor on.</param>
/// <param name="FieldName">The schema column name to anchor on (resolved to a table column by the reporter).</param>
/// <param name="Text">The full note text, including the ownership sentinel.</param>
internal sealed record NotePayload(int RowNumber, string FieldName, string Text);
