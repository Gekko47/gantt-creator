using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Office;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ValidationReportComposer"/>: the pure grouping,
/// ordering, sentinel, and truncation rules behind the R2.6 row-level report.
/// No Excel and no interop types are involved.
/// </summary>
/// <remarks>
/// AGENTS.md validator rule: every guard in the composer has a positive test
/// here that constructs the bad input and asserts the failure path.
/// </remarks>
public class ValidationReportComposerTests
{
    private static GanttValidationIssue Issue(
        int row,
        string field,
        string code,
        GanttValidationSeverity severity,
        string message) => new(row, field, code, severity, message);

    [Fact]
    public void GroupIntoNotes_groups_consecutive_same_cell_issues_into_one_note()
    {
        // Two findings on the same body row and column must collapse into a
        // single cell note, because Excel allows one classic note per cell.
        GanttValidationIssue first = Issue(
            3, "Start", "GC1001", GanttValidationSeverity.Error, "Start is not a date.");
        GanttValidationIssue second = Issue(
            3, "Start", "GC1002", GanttValidationSeverity.Warning, "Start is before the project window.");

        IReadOnlyList<NotePayload> notes = ValidationReportComposer.GroupIntoNotes([first, second]);

        NotePayload note = Assert.Single(notes);
        Assert.Equal(3, note.RowNumber);
        Assert.Equal("Start", note.FieldName);
        Assert.Contains("Start is not a date.", note.Text, StringComparison.Ordinal);
        Assert.Contains("Start is before the project window.", note.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void GroupIntoNotes_keeps_the_supplied_severity_row_field_code_order_within_a_cell()
    {
        // The validator guarantees Severity -> Row -> Field -> Code order; the
        // composer must not reorder, so the note lines follow the list order.
        GanttValidationIssue error = Issue(
            2, "Finish", "GC1003", GanttValidationSeverity.Error, "Finish is before Start.");
        GanttValidationIssue warning = Issue(
            2, "Finish", "GC1004", GanttValidationSeverity.Warning, "Finish is unusually long.");

        NotePayload note = Assert.Single(ValidationReportComposer.GroupIntoNotes([error, warning]));

        var errorIndex = note.Text.IndexOf("Finish is before Start.", StringComparison.Ordinal);
        var warningIndex = note.Text.IndexOf("Finish is unusually long.", StringComparison.Ordinal);
        Assert.True(
            errorIndex >= 0 && warningIndex > errorIndex,
            $"Expected the error line before the warning line. Text: {note.Text}");
    }

    [Fact]
    public void GroupIntoNotes_produces_one_note_per_cell_in_first_appearance_order()
    {
        GanttValidationIssue rowOneType = Issue(
            1, "Type", "GC2001", GanttValidationSeverity.Error, "Type is unknown.");
        GanttValidationIssue rowOneLane = Issue(
            1, "LaneId", "GC2002", GanttValidationSeverity.Error, "LaneId is blank.");
        GanttValidationIssue rowTwoType = Issue(
            2, "Type", "GC2003", GanttValidationSeverity.Warning, "Type is deprecated.");

        IReadOnlyList<NotePayload> notes =
            ValidationReportComposer.GroupIntoNotes([rowOneType, rowOneLane, rowTwoType]);

        Assert.Equal(3, notes.Count);
        Assert.Equal((1, "Type"), (notes[0].RowNumber, notes[0].FieldName));
        Assert.Equal((1, "LaneId"), (notes[1].RowNumber, notes[1].FieldName));
        Assert.Equal((2, "Type"), (notes[2].RowNumber, notes[2].FieldName));
    }

    [Fact]
    public void GroupIntoNotes_does_not_merge_issues_that_differ_only_by_severity_across_cells()
    {
        // Same severity pair on two different rows: separate cells, separate notes.
        GanttValidationIssue first = Issue(
            1, "Visible", "GC3001", GanttValidationSeverity.Error, "Visible is not a flag.");
        GanttValidationIssue second = Issue(
            2, "Visible", "GC3002", GanttValidationSeverity.Error, "Visible is not a flag.");

        IReadOnlyList<NotePayload> notes = ValidationReportComposer.GroupIntoNotes([first, second]);

        Assert.Equal(2, notes.Count);
        Assert.Equal(1, notes[0].RowNumber);
        Assert.Equal(2, notes[1].RowNumber);
    }

    [Fact]
    public void GroupIntoNotes_prefixes_every_note_with_the_ownership_sentinel()
    {
        // The reporter recognises its own notes by this prefix (Comment.Author is
        // read-only), so every written note must carry it.
        NotePayload note = Assert.Single(ValidationReportComposer.GroupIntoNotes(
            [Issue(1, "Id", "GC4001", GanttValidationSeverity.Error, "Id is blank.")]));

        Assert.StartsWith(ValidationReportComposer.NoteSentinelPrefix, note.Text, StringComparison.Ordinal);
        Assert.True(ValidationReportComposer.IsOwnedByAddIn(note.Text));
    }

    [Fact]
    public void GroupIntoNotes_names_the_row_and_column_in_each_note()
    {
        NotePayload note = Assert.Single(ValidationReportComposer.GroupIntoNotes(
            [Issue(7, "ParentId", "GC5001", GanttValidationSeverity.Error, "ParentId is unknown.")]));

        Assert.Contains("row 7", note.Text, StringComparison.Ordinal);
        Assert.Contains("ParentId", note.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void GroupIntoNotes_labels_the_severity_of_each_line()
    {
        NotePayload note = Assert.Single(ValidationReportComposer.GroupIntoNotes(
        [
            Issue(1, "Start", "GC6001", GanttValidationSeverity.Error, "Start is invalid."),
            Issue(1, "Start", "GC6002", GanttValidationSeverity.Warning, "Start is far outside the window."),
        ]));

        Assert.Contains("Error: Start is invalid.", note.Text, StringComparison.Ordinal);
        Assert.Contains("Warning: Start is far outside the window.", note.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void GroupIntoNotes_returns_empty_for_no_issues()
    {
        Assert.Empty(ValidationReportComposer.GroupIntoNotes([]));
    }

    [Fact]
    public void GroupIntoNotes_throws_for_a_null_issue_list()
    {
        var exception = Record.Exception(() => ValidationReportComposer.GroupIntoNotes(null!));

        Assert.IsType<ArgumentNullException>(exception);
    }

    [Fact]
    public void GroupIntoNotes_truncates_text_to_the_excel_note_limit()
    {
        // Range.AddComment rejects a body longer than 255 characters; the composer
        // truncates so a verbose issue set can never fail the write outright.
        var message = new string('x', 400);
        NotePayload note = Assert.Single(ValidationReportComposer.GroupIntoNotes(
            [Issue(1, "Description", "GC7001", GanttValidationSeverity.Error, message)]));

        Assert.Equal(ValidationReportComposer.MaxNoteLength, note.Text.Length);
        Assert.StartsWith(ValidationReportComposer.NoteSentinelPrefix, note.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void GroupIntoNotes_is_culture_invariant()
    {
        GanttValidationIssue issue = Issue(
            1234, "StackIndex", "GC8001", GanttValidationSeverity.Error, "StackIndex is not a number.");
        CultureInfo original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("tr-TR");
            NotePayload turkish = Assert.Single(ValidationReportComposer.GroupIntoNotes([issue]));

            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
            NotePayload invariant = Assert.Single(ValidationReportComposer.GroupIntoNotes([issue]));

            Assert.Equal(turkish.Text, invariant.Text);
            Assert.Contains("row 1234", turkish.Text, StringComparison.Ordinal);
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Fact]
    public void IsOwnedByAddIn_recognises_only_the_complete_generated_marker()
    {
        // Ownership needs the whole marker the composer writes. A bare prefix is
        // not proof of ownership: a user can legitimately start a line with the
        // same phrase, and treating that as add-in content would delete their text.
        Assert.True(ValidationReportComposer.IsOwnedByAddIn(
            ValidationReportComposer.NoteSentinelPrefix + " row 1, column 'Id': Error: Id is blank."));

        Assert.False(ValidationReportComposer.IsOwnedByAddIn(
            ValidationReportComposer.NoteSentinelPrefix));

        Assert.False(ValidationReportComposer.IsOwnedByAddIn(null));
        Assert.False(ValidationReportComposer.IsOwnedByAddIn(string.Empty));
        Assert.False(ValidationReportComposer.IsOwnedByAddIn("A note the user typed."));
        Assert.False(ValidationReportComposer.IsOwnedByAddIn("gannt creator validation: lowercase"));
    }

    [Fact]
    public void IsOwnedByAddIn_accepts_a_line_anchored_section_after_user_text()
    {
        // The add-in always writes its sentinel at the start of a line, so a
        // section following the user's text is still add-in owned.
        Assert.True(ValidationReportComposer.IsOwnedByAddIn(
            "My own note."
            + "\n"
            + ValidationReportComposer.NoteSentinelPrefix + " row 1, column 'Id': Error: Id is blank."));
    }

    [Fact]
    public void IsOwnedByAddIn_rejects_a_mid_sentence_mention()
    {
        // Regression: ownership was previously a substring test, so a user's own
        // note that merely mentioned the phrase was treated as add-in content and
        // truncated by StripOwnedSection, destroying their text.
        const string UserNote = "I use Gantt Creator validation: as my own checklist.";

        Assert.False(ValidationReportComposer.IsOwnedByAddIn(UserNote));
        Assert.Equal(UserNote, ValidationReportComposer.StripOwnedSection(UserNote));
    }

    [Fact]
    public void IsOwnedByAddIn_ignores_a_later_mid_sentence_mention_and_keeps_the_real_section()
    {
        // A note can mention the phrase in the user's part and still carry a real
        // owned section; only the line-anchored one counts.
        var text = "I use Gantt Creator validation: myself."
            + "\n"
            + ValidationReportComposer.NoteSentinelPrefix + " row 2, column 'Type': Error: unknown type.";

        Assert.True(ValidationReportComposer.IsOwnedByAddIn(text));
        Assert.Equal("I use Gantt Creator validation: myself.", ValidationReportComposer.StripOwnedSection(text));
    }

    [Fact]
    public void GroupIntoNotes_separates_message_lines_with_newlines()
    {
        // Line structure is what makes ownership detectable after user text is
        // prepended, so the composer must emit real line breaks.
        NotePayload note = Assert.Single(ValidationReportComposer.GroupIntoNotes(
        [
            Issue(1, "Start", "GC9101", GanttValidationSeverity.Error, "First problem."),
            Issue(1, "Start", "GC9102", GanttValidationSeverity.Warning, "Second problem."),
        ]));

        var lines = note.Text.Split('\n');
        Assert.Equal(3, lines.Length);
        Assert.StartsWith(ValidationReportComposer.NoteSentinelPrefix, lines[0], StringComparison.Ordinal);
        Assert.Equal("Error: First problem.", lines[1]);
        Assert.Equal("Warning: Second problem.", lines[2]);
    }

    [Fact]
    public void Truncate_marks_a_clipped_note_and_never_splits_a_surrogate_pair()
    {
        // One ASCII character, then a real astral character (surrogate pair)
        // repeated until the text overruns the limit. The leading ASCII unit
        // makes the clamp point land on a high surrogate, which the guard must
        // drop: cutting between a high and low surrogate would make the note
        // text invalid UTF-16 and Excel would reject it.
        var builder = new System.Text.StringBuilder();
        _ = builder.Append('x');
        while (builder.Length <= ValidationReportComposer.MaxNoteLength)
        {
            _ = builder.Append("\U0001F600"); // GRINNING FACE (astral, surrogate pair)
        }

        string clipped = ValidationReportComposer.Truncate(builder.ToString());

        // One unit shorter than the maximum: the orphaned high surrogate was
        // removed before the marker was appended.
        Assert.Equal(ValidationReportComposer.MaxNoteLength - 1, clipped.Length);
        Assert.EndsWith(ValidationReportComposer.TruncationMarker, clipped, StringComparison.Ordinal);
        // The last kept unit before the marker must be a complete pair: a low
        // surrogate preceded by its high surrogate, never a lone high surrogate.
        var markerStart = clipped.Length - ValidationReportComposer.TruncationMarker.Length;
        Assert.False(char.IsHighSurrogate(clipped[markerStart - 1]));
        Assert.False(char.IsLowSurrogate(clipped[0]));
    }

    [Fact]
    public void GroupIntoNotes_truncation_is_marked_and_within_the_limit()
    {
        var message = new string('x', 400);
        NotePayload note = Assert.Single(ValidationReportComposer.GroupIntoNotes(
            [Issue(1, "Description", "GC9201", GanttValidationSeverity.Error, message)]));

        Assert.Equal(ValidationReportComposer.MaxNoteLength, note.Text.Length);
        Assert.EndsWith(ValidationReportComposer.TruncationMarker, note.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void Truncate_leaves_text_within_the_limit_unchanged()
    {
        const string text = "short note";
        Assert.Equal(text, ValidationReportComposer.Truncate(text));
    }

    [Fact]
    public void StripOwnedSection_returns_the_user_text_before_the_marker()
    {
        // Excel allows one classic note per cell, so the add-in section is
        // appended to the user's text; stripping must return exactly their text.
        var text = "Reminder I typed myself."
            + Environment.NewLine
            + ValidationReportComposer.NoteSentinelPrefix
            + " row 1, column 'Start': Error: Start is invalid.";

        Assert.Equal("Reminder I typed myself.", ValidationReportComposer.StripOwnedSection(text));
    }

    [Fact]
    public void StripOwnedSection_returns_empty_for_a_purely_add_in_note()
    {
        var text = ValidationReportComposer.NoteSentinelPrefix + " row 1, column 'Id': Error: Id is blank.";

        Assert.Equal(string.Empty, ValidationReportComposer.StripOwnedSection(text));
    }

    [Fact]
    public void StripOwnedSection_returns_a_note_without_the_marker_unchanged()
    {
        Assert.Equal("Just my note.", ValidationReportComposer.StripOwnedSection("Just my note."));
    }

    [Fact]
    public void StripOwnedSection_trims_trailing_blank_lines_left_by_an_earlier_section()
    {
        var text = "My note." + Environment.NewLine + Environment.NewLine
            + ValidationReportComposer.NoteSentinelPrefix + " row 2, column 'Type': Error: unknown type.";

        Assert.Equal("My note.", ValidationReportComposer.StripOwnedSection(text));
    }

    [Fact]
    public void StripOwnedSection_handles_null_and_empty_text()
    {
        Assert.Equal(string.Empty, ValidationReportComposer.StripOwnedSection(null));
        Assert.Equal(string.Empty, ValidationReportComposer.StripOwnedSection(string.Empty));
    }

    [Fact]
    public void StripOwnedSection_keeps_a_user_line_that_only_begins_with_the_sentinel()
    {
        // The line anchor alone was not enough: a user note whose own line starts
        // with the phrase was treated as add-in content, so everything from that
        // line on was deleted and the user lost their text. Requiring the
        // complete generated-section marker is what preserves it.
        var text = "Chased with the vendor."
            + Environment.NewLine
            + ValidationReportComposer.NoteSentinelPrefix
            + " is how I describe this check to the team.";

        Assert.False(ValidationReportComposer.IsOwnedByAddIn(text));
        Assert.Equal(text, ValidationReportComposer.StripOwnedSection(text));
    }

    [Theory]
    [InlineData(" row 1, column 'Id':")]
    [InlineData(" row 12, column 'ParentId':")]
    public void IsOwnedByAddIn_accepts_a_complete_generated_section_marker(string headerSuffix)
    {
        // The positive control for the stricter marker check: the exact shape
        // BuildNoteText writes must still be recognised as owned, or the
        // reporter would stop clearing its own stale sections.
        var text = ValidationReportComposer.NoteSentinelPrefix + headerSuffix + " Error: something.";

        Assert.True(ValidationReportComposer.IsOwnedByAddIn(text));
        Assert.Equal(string.Empty, ValidationReportComposer.StripOwnedSection(text));
    }

    [Theory]
    [InlineData("")]                       // bare prefix, no header at all
    [InlineData(" row , column 'Id':")]   // missing row number
    [InlineData(" row 1 column 'Id':")]    // missing comma before column
    [InlineData(" row 1, column Id:")]     // missing quotes around the column
    [InlineData(" column 'Id':")]         // missing row segment
    [InlineData(" row 1, column 'Id")]    // no closing quote/colon on the header line
    [InlineData(" row 1, column 'Id\n'Id': tail")]  // header wraps; the closing
                                                    // marker is on a later line
    public void IsOwnedByAddIn_rejects_an_incomplete_marker(string headerSuffix)
    {
        // Each variant is a near-miss that a looser check would accept and then
        // silently delete a user's note over. With no complete marker found the
        // whole text is the user's, so stripping returns it unchanged.
        var text = "My own note." + Environment.NewLine + ValidationReportComposer.NoteSentinelPrefix + headerSuffix;

        Assert.False(ValidationReportComposer.IsOwnedByAddIn(text));
        Assert.Equal(text, ValidationReportComposer.StripOwnedSection(text));
    }

    [Fact]
    public void StripOwnedSection_skips_a_bare_prefix_and_still_strips_a_real_section_below_it()
    {
        // The scan must not stop at the first prefix occurrence: a user line that
        // merely starts with the phrase comes first, and the genuine generated
        // section after it still has to be removed.
        var userLine = ValidationReportComposer.NoteSentinelPrefix + " is my shorthand.";
        var section = ValidationReportComposer.NoteSentinelPrefix + " row 3, column 'Type': Error: unknown type.";
        var text = userLine + Environment.NewLine + section;

        Assert.Equal(userLine, ValidationReportComposer.StripOwnedSection(text));
    }

    [Fact]
    public void StripOwnedSection_then_recompose_is_idempotent()
    {
        // Two runs must not accumulate duplicate sections: stripping the result of
        // the first run and appending the second leaves exactly one section.
        const string userText = "Keep me.";
        GanttValidationIssue issue = Issue(
            1, "Start", "GC9501", GanttValidationSeverity.Error, "Start is invalid.");

        var first = userText + Environment.NewLine + Assert.Single(
            ValidationReportComposer.GroupIntoNotes([issue])).Text;
        var second = ValidationReportComposer.StripOwnedSection(first)
            + Environment.NewLine
            + Assert.Single(ValidationReportComposer.GroupIntoNotes([issue])).Text;

        Assert.Equal(first, second);
        var occurrences = second.Split(ValidationReportComposer.NoteSentinelPrefix).Length - 1;
        Assert.Equal(1, occurrences);
    }
}
