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
    /// The suffix appended when a note is clamped to
    /// <see cref="MaxNoteLength"/>, so a clipped note is visibly clipped instead
    /// of silently ending mid-sentence.
    /// </summary>
    internal const string TruncationMarker = "...";

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

            var text = Truncate(BuildNoteText(current.RowNumber, current.Field, lines));
            result.Add(new NotePayload(current.RowNumber, current.Field, text));
        }

        return result;
    }

    /// <summary>
    /// Returns <c>true</c> when the note text carries an add-in-owned section.
    /// The section is recognised only when the sentinel begins the text or begins
    /// a line within it, so a user note that merely mentions the phrase
    /// mid-sentence is never treated as add-in content and never truncated.
    /// </summary>
    /// <param name="text">The note text, or <see lang="null"/>.</param>
    /// <returns>
    /// <see langword="true"/> when the text starts a line with
    /// <see cref="NoteSentinelPrefix"/>.
    /// </returns>
    public static bool IsOwnedByAddIn(string? text) => FindOwnedSectionStart(text) >= 0;

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
        var marker = FindOwnedSectionStart(text);
        return marker < 0
            ? (string.IsNullOrEmpty(text) ? string.Empty : text.TrimEnd('\r', '\n', ' '))
            : text![..marker].TrimEnd('\r', '\n', ' ');
    }

    /// <summary>
    /// Locates the first line-anchored <see cref="NoteSentinelPrefix"/> in a note.
    /// </summary>
    /// <remarks>
    /// Ownership is a line-structured property, not a substring property: the
    /// add-in always writes the sentinel at the start of a line (either the whole
    /// note is the add-in's, or the section follows the user's text on its own
    /// line). Requiring that anchor is what makes user text safe — a note reading
    /// <c>"I use Gantt Creator validation: for my own reasons"</c> is the user's
    /// note and must survive untouched.
    /// </remarks>
    /// <param name="text">The note text, or <see lang="null"/>.</param>
    /// <returns>The start index of the owned section, or <c>-1</c> when absent.</returns>
    private static int FindOwnedSectionStart(string? text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return -1;
        }

        var searchFrom = 0;
        while (searchFrom <= text.Length - NoteSentinelPrefix.Length)
        {
            var index = text.IndexOf(NoteSentinelPrefix, searchFrom, StringComparison.Ordinal);
            if (index < 0)
            {
                return -1;
            }

            if (index == 0 || text[index - 1] is '\n' or '\r')
            {
                return index;
            }

            searchFrom = index + 1;
        }

        return -1;
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
        _ = sb.Append(NoteSentinelPrefix);
        _ = sb.AppendFormat(CultureInfo.InvariantCulture, " row {0}, column '{1}':", rowNumber, fieldName);
        foreach (var line in messageLines)
        {
            _ = sb.Append('\n').Append(line);
        }

        return sb.ToString();
    }

    /// <summary>
    /// Clamps a note body to <see cref="MaxNoteLength"/> with an explicit marker so a
    /// reader can tell the note was clipped, never splitting a surrogate pair and
    /// never leaving a dangling high surrogate that Excel would reject.
    /// </summary>
    /// <param name="text">The composed note text.</param>
    /// <returns>
    /// The text unchanged when it fits, otherwise the clamped prefix followed by
    /// <see cref="TruncationMarker"/>.
    /// </returns>
    internal static string Truncate(string text)
    {
        if (text.Length <= MaxNoteLength)
        {
            return text;
        }

        var budget = MaxNoteLength - TruncationMarker.Length;
        // A cut landing between a high and low surrogate would leave invalid
        // UTF-16; back up one unit so the last kept unit is a complete pair.
        if (char.IsHighSurrogate(text[budget - 1]))
        {
            budget--;
        }

        return string.Concat(text.AsSpan(0, budget), TruncationMarker);
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
