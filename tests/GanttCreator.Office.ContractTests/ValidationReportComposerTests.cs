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
    public void IsOwnedByAddIn_recognises_only_the_sentinel_prefix()
    {
        Assert.True(ValidationReportComposer.IsOwnedByAddIn(
            ValidationReportComposer.NoteSentinelPrefix + " row 1, column 'Id': Error: Id is blank."));
        Assert.True(ValidationReportComposer.IsOwnedByAddIn(
            ValidationReportComposer.NoteSentinelPrefix));

        Assert.False(ValidationReportComposer.IsOwnedByAddIn(null));
        Assert.False(ValidationReportComposer.IsOwnedByAddIn(string.Empty));
        Assert.False(ValidationReportComposer.IsOwnedByAddIn("A note the user typed."));
        Assert.False(ValidationReportComposer.IsOwnedByAddIn("gannt creator validation: lowercase"));
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
