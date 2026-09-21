using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contract tests for <see cref="ValidateSheetCommand"/> (work item R2.6): the
/// read → validate → report pipeline, the typed refusals, the single summary
/// message, and the CA1031 presenter boundary. None of these require Excel —
/// both ports are Moq proxies.
/// </summary>
/// <remarks>
/// AGENTS.md validator rule: every refusal path exercised here is a positive
/// test that constructs the bad input and asserts the typed refusal fires.
/// </remarks>
public class ValidateSheetCommandTests
{
    /// <summary>A well-formed <see cref="GanttRowId"/>: <c>G-</c> plus 32 hex characters.</summary>
    private static string NewId(char last) => "G-" + new string('0', 31) + last;

    private static GanttRowDto ValidRow(int rowNumber, char idSuffix) => new(
        rowNumber,
        NewId(idSuffix),
        NewId('f'),
        0,
        "As-Planned Activity",
        "Excavate",
        new DateOnly(2026, 9, 1),
        new DateOnly(2026, 9, 5),
        null,
        null,
        null,
        null,
        null,
        true,
        null);

    private static GanttRowDto InvalidRow(int rowNumber) => new(
        rowNumber,
        null,
        null,
        1,
        "Not a catalogued type",
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null);

    private static Mock<IGanttTableReader> ReaderWith(GanttTableReadOutcome outcome)
    {
        var reader = new Mock<IGanttTableReader>();
        _ = reader.Setup(r => r.Read()).Returns(outcome);
        return reader;
    }

    private static Mock<IGanttValidationReporter> ReporterWith(GanttValidationReportOutcome outcome)
    {
        var reporter = new Mock<IGanttValidationReporter>();
        _ = reporter
            .Setup(r => r.Report(It.IsAny<IReadOnlyList<GanttValidationIssue>>()))
            .Returns(outcome);
        return reporter;
    }

    private static readonly Mock<IGanttTableReader> UnusedReader = new();

    private static readonly Mock<IGanttValidationReporter> UnusedReporter = new();

    [Fact]
    public void Run_throws_for_a_null_reader()
    {
        var exception = Record.Exception(
            () => ValidateSheetCommand.Run(null!, UnusedReporter.Object, _ => { }));

        Assert.IsType<ArgumentNullException>(exception);
    }

    [Fact]
    public void Run_throws_for_a_null_reporter()
    {
        var exception = Record.Exception(
            () => ValidateSheetCommand.Run(UnusedReader.Object, null!, _ => { }));

        Assert.IsType<ArgumentNullException>(exception);
    }

    [Fact]
    public void Run_throws_for_a_null_presenter()
    {
        var exception = Record.Exception(
            () => ValidateSheetCommand.Run(UnusedReader.Object, UnusedReporter.Object, null!));

        Assert.IsType<ArgumentNullException>(exception);
    }

    [Fact]
    public void A_valid_table_presents_all_rows_valid_and_writes_no_note()
    {
        var messages = new List<string>();
        var reader = ReaderWith(GanttTableReadOutcome.Ok([ValidRow(1, '1'), ValidRow(2, '2')]));
        var reporter = ReporterWith(GanttValidationReportOutcome.Ok(0));

        ValidateSheetCommand.Run(reader.Object, reporter.Object, messages.Add);

        var message = Assert.Single(messages);
        Assert.Contains("all rows valid", message, StringComparison.OrdinalIgnoreCase);
        reporter.Verify(
            r => r.Report(It.IsAny<IReadOnlyList<GanttValidationIssue>>()),
            Times.Once);
    }

    [Fact]
    public void An_invalid_table_presents_the_counts_and_forwards_every_validator_issue()
    {
        var messages = new List<string>();
        var rows = new List<GanttRowDto> { ValidRow(1, '3'), InvalidRow(2) };
        GanttValidationOutcome expected = GanttRowValidator.Validate(rows);
        Assert.Contains(expected.Issues, issue => issue.Severity == GanttValidationSeverity.Error);

        var reader = ReaderWith(GanttTableReadOutcome.Ok(rows));
        var forwarded = new List<GanttValidationIssue>();
        var reporter = new Mock<IGanttValidationReporter>();
        _ = reporter
            .Setup(r => r.Report(It.IsAny<IReadOnlyList<GanttValidationIssue>>()))
            .Callback<IReadOnlyList<GanttValidationIssue>>(issues => forwarded.AddRange(issues))
            .Returns(GanttValidationReportOutcome.Ok(1));

        ValidateSheetCommand.Run(reader.Object, reporter.Object, messages.Add);

        Assert.Equal(expected.Issues, forwarded);
        var message = Assert.Single(messages);
        Assert.Contains("error(s)", message, StringComparison.Ordinal);
        Assert.Contains("tblGanttData", message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_read_refusal_presents_one_actionable_message_and_never_calls_the_reporter()
    {
        var messages = new List<string>();
        var reader = ReaderWith(
            GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.NoActiveWorkbook));
        var reporter = ReporterWith(GanttValidationReportOutcome.Ok(0));

        ValidateSheetCommand.Run(reader.Object, reporter.Object, messages.Add);

        var message = Assert.Single(messages);
        Assert.Contains("workbook", message, StringComparison.OrdinalIgnoreCase);
        reporter.Verify(
            r => r.Report(It.IsAny<IReadOnlyList<GanttValidationIssue>>()),
            Times.Never);
    }

    [Fact]
    public void A_table_missing_refusal_names_the_table_and_the_recovery_action()
    {
        var messages = new List<string>();
        var reader = ReaderWith(
            GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.TableMissing));

        ValidateSheetCommand.Run(reader.Object, ReporterWith(GanttValidationReportOutcome.Ok(0)).Object, messages.Add);

        var message = Assert.Single(messages);
        Assert.Contains("tblGanttData", message, StringComparison.Ordinal);
        Assert.Contains("Initialise sheet", message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_date_system_refusal_names_the_1904_limitation()
    {
        var messages = new List<string>();
        var reader = ReaderWith(
            GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.DateSystemUnsupported));

        ValidateSheetCommand.Run(reader.Object, ReporterWith(GanttValidationReportOutcome.Ok(0)).Object, messages.Add);

        var message = Assert.Single(messages);
        Assert.Contains("1904", message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_report_refusal_presents_one_actionable_message()
    {
        var messages = new List<string>();
        var reader = ReaderWith(GanttTableReadOutcome.Ok([ValidRow(1, '4')]));
        var reporter = ReporterWith(
            GanttValidationReportOutcome.Refused(GanttValidationReportRefusalReason.TableMissing));

        ValidateSheetCommand.Run(reader.Object, reporter.Object, messages.Add);

        var message = Assert.Single(messages);
        Assert.Contains("tblGanttData", message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_report_no_active_workbook_refusal_presents_one_actionable_message()
    {
        var messages = new List<string>();
        var reader = ReaderWith(GanttTableReadOutcome.Ok([ValidRow(1, '5')]));
        var reporter = ReporterWith(
            GanttValidationReportOutcome.Refused(GanttValidationReportRefusalReason.NoActiveWorkbook));

        ValidateSheetCommand.Run(reader.Object, reporter.Object, messages.Add);

        var message = Assert.Single(messages);
        Assert.Contains("workbook", message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void A_failing_presenter_degrades_to_no_message_instead_of_throwing()
    {
        // CA1031: the presenter is outside this command's contract — a dialog
        // that cannot render must not propagate into the Ribbon callback.
        var reader = ReaderWith(GanttTableReadOutcome.Ok([ValidRow(1, '6')]));
        var reporter = ReporterWith(GanttValidationReportOutcome.Ok(0));

        var exception = Record.Exception(() => ValidateSheetCommand.Run(
            reader.Object, reporter.Object, _ => throw new InvalidOperationException("dialog down")));

        Assert.Null(exception);
        reporter.Verify(
            r => r.Report(It.IsAny<IReadOnlyList<GanttValidationIssue>>()),
            Times.Once);
    }

    [Fact]
    public void A_failing_presenter_on_a_refusal_degrades_to_no_message()
    {
        var reader = ReaderWith(
            GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.NoActiveWorkbook));

        var exception = Record.Exception(() => ValidateSheetCommand.Run(
            reader.Object,
            UnusedReporter.Object,
            _ => throw new InvalidOperationException("dialog down")));

        Assert.Null(exception);
    }

    [Fact]
    public void SummaryMessage_reports_counts_without_schedule_content()
    {
        List<GanttValidationIssue> issues =
        [
            new(1, "Start", GanttValidationCodes.StartRequired, GanttValidationSeverity.Error, "Start is required."),
            new(1, "Finish", GanttValidationCodes.NotUsedByType, GanttValidationSeverity.Warning, "Finish is ignored."),
        ];

        var message = ValidateSheetCommand.SummaryMessage(issues, notesWritten: 1);

        Assert.Contains("1 error(s)", message, StringComparison.Ordinal);
        Assert.Contains("1 warning(s)", message, StringComparison.Ordinal);
        Assert.Contains("1 note(s)", message, StringComparison.Ordinal);
        Assert.DoesNotContain("Start is required.", message, StringComparison.Ordinal);
    }

    [Fact]
    public void SummaryMessage_labels_an_empty_issue_list_as_all_rows_valid()
    {
        var message = ValidateSheetCommand.SummaryMessage([], notesWritten: 0);

        Assert.Contains("all rows valid", message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Every_read_refusal_has_a_distinct_user_safe_message()
    {
        var seen = new List<string>();
        foreach (GanttTableReadRefusalReason reason in Enum.GetValues<GanttTableReadRefusalReason>())
        {
            var message = ValidateSheetCommand.TranslateReadRefusal(reason);
            Assert.False(string.IsNullOrWhiteSpace(message));
            Assert.DoesNotContain(message, seen, StringComparer.Ordinal);
            seen.Add(message);
        }
    }

    [Fact]
    public void Every_report_refusal_has_a_distinct_user_safe_message()
    {
        var seen = new List<string>();
        foreach (GanttValidationReportRefusalReason reason in Enum.GetValues<GanttValidationReportRefusalReason>())
        {
            var message = ValidateSheetCommand.TranslateReportRefusal(reason);
            Assert.False(string.IsNullOrWhiteSpace(message));
            Assert.DoesNotContain(message, seen, StringComparer.Ordinal);
            seen.Add(message);
        }
    }
}
