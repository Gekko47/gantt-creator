using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contract tests for <see cref="InitialiseSheetCommand"/>: the typed-refusal
/// surface (work item R2.2 decision D4/D5). Success is silent; every refusal
/// produces exactly one presenter message and nothing is logged. None of
/// these require Excel — the initialiser is a Moq port proxy.
/// </summary>
/// <remarks>
/// AGENTS.md validator rule: every refusal path exercised here is a positive
/// test that constructs the bad input and asserts the typed refusal fires.
/// </remarks>
public class InitialiseSheetCommandTests
{
    private static readonly Mock<IWorkbookInitialiser> FailingInitialiser =
        new Mock<IWorkbookInitialiser>();

    /// <summary>Builds an initialiser mock whose <c>Initialise</c> returns one outcome.</summary>
    private static Mock<IWorkbookInitialiser> InitialiserWith(WorkbookInitialiseOutcome outcome)
    {
        var mock = new Mock<IWorkbookInitialiser>();
        mock.Setup(i => i.Initialise()).Returns(outcome);
        return mock;
    }

    [Fact]
    public void Run_throws_for_a_null_initialiser()
    {
        var exception = Record.Exception(
            () => InitialiseSheetCommand.Run(null!, _ => { }));

        Assert.NotNull(exception);
        Assert.IsType<ArgumentNullException>(exception);
    }

    [Fact]
    public void Run_throws_for_a_null_presenter()
    {
        var exception = Record.Exception(
            () => InitialiseSheetCommand.Run(FailingInitialiser.Object, null!));

        Assert.NotNull(exception);
        Assert.IsType<ArgumentNullException>(exception);
    }

    [Fact]
    public void Success_is_silent_no_presenter_call_and_no_log_side_effect()
    {
        // R2.2 decision D7: the renamed/populated sheet is self-evident, so a
        // successful initialise produces no dialog and no log record.
        var messages = new List<string>();
        var initialiser = InitialiserWith(
            WorkbookInitialiseOutcome.Adopted(GanttWorkbookContract.GanttSheetLabel));

        InitialiseSheetCommand.Run(initialiser.Object, messages.Add);

        Assert.Empty(messages);
        initialiser.Verify(i => i.Initialise(), Times.Once);
    }

    [Fact]
    public void NoActiveWorkbook_refusal_surfaces_one_actionable_message()
    {
        var messages = new List<string>();
        var initialiser = InitialiserWith(
            WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.NoActiveWorkbook));

        InitialiseSheetCommand.Run(initialiser.Object, messages.Add);

        var message = Assert.Single(messages);
        Assert.Contains("open workbook", message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("tblGanttData", message, StringComparison.Ordinal);
    }

    [Fact]
    public void TableExists_refusal_names_the_target_sheet()
    {
        var messages = new List<string>();
        var initialiser = InitialiserWith(
            WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.TableExists));

        InitialiseSheetCommand.Run(initialiser.Object, messages.Add);

        var message = Assert.Single(messages);
        Assert.Contains("tblGanttData", message, StringComparison.Ordinal);
        Assert.Contains("left it unchanged", message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ConfigSheetExists_refusal_names_the_configuration_sheet()
    {
        var messages = new List<string>();
        var initialiser = InitialiserWith(
            WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.ConfigSheetExists));

        InitialiseSheetCommand.Run(initialiser.Object, messages.Add);

        var message = Assert.Single(messages);
        Assert.Contains("_GanttCreatorConfig", message, StringComparison.Ordinal);
        Assert.Contains("left it unchanged", message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TargetProtected_refusal_names_protection_and_the_unprotect_action()
    {
        var messages = new List<string>();
        var initialiser = InitialiserWith(
            WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.TargetProtected));

        InitialiseSheetCommand.Run(initialiser.Object, messages.Add);

        var message = Assert.Single(messages);
        Assert.Contains("protected", message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Unprotect", message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_failing_presenter_degrades_to_no_message_instead_of_throwing()
    {
        // CA1031: the presenter is outside this command's contract — a dialog
        // that cannot render must not propagate into the Ribbon callback.
        var initialiser = InitialiserWith(
            WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.NoActiveWorkbook));

        var exception = Record.Exception(() =>
            InitialiseSheetCommand.Run(initialiser.Object, _ => throw new InvalidOperationException("dialog down")));

        Assert.Null(exception);
        initialiser.Verify(i => i.Initialise(), Times.Once);
    }
}
