using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// R4.8A D1: the refresh command is a boundary, not an orchestrator. It invokes the
/// service exactly once and presents what comes back.
/// </summary>
/// <remarks>
/// The claim worth testing is the negative one — that the command holds no pipeline
/// of its own. A test cannot assert the absence of code directly, so it asserts the
/// observable consequences: exactly one invocation, no retry, and no presentation on
/// success.
/// </remarks>
public class RefreshSheetCommandTests
{
    private static GanttRefreshOutcome Success() =>
        GanttRefreshOutcome.Ok([], shapesWritten: 3, durationCellsWritten: 2, styleKeyCellsWritten: 0);

    private static GanttRefreshOutcome Refused(GanttRefreshRefusal refusal) =>
        GanttRefreshOutcome.Refused(refusal, "the chart could not be refreshed");

    /// <summary>
    /// The command invokes the orchestrator exactly once.
    /// </summary>
    /// <remarks>
    /// Once, not "at least once". A retry loop would re-run the whole pipeline against
    /// a chart the first attempt already partly wrote, which is how a transient COM
    /// refusal becomes a corrupted sheet.
    /// </remarks>
    [Fact]
    public void Run_invokes_the_orchestrator_exactly_once()
    {
        var orchestrator = new Mock<IGanttRefreshOrchestrator>();
        _ = orchestrator.Setup(o => o.Refresh()).Returns(Success());

        RefreshSheetCommand.Run(orchestrator.Object, _ => { });

        orchestrator.Verify(o => o.Refresh(), Times.Once);
    }

    /// <summary>A successful refresh presents nothing.</summary>
    [Fact]
    public void Run_presents_nothing_when_the_refresh_succeeded()
    {
        var orchestrator = new Mock<IGanttRefreshOrchestrator>();
        _ = orchestrator.Setup(o => o.Refresh()).Returns(Success());
        var presented = new List<string>();

        RefreshSheetCommand.Run(orchestrator.Object, presented.Add);

        Assert.Empty(presented);
    }

    /// <summary>A refusal is presented with the orchestrator's own message.</summary>
    [Fact]
    public void Run_presents_the_orchestrators_message_on_a_refusal()
    {
        var orchestrator = new Mock<IGanttRefreshOrchestrator>();
        _ = orchestrator
            .Setup(o => o.Refresh())
            .Returns(Refused(GanttRefreshRefusal.BlockingValidationErrors));
        var presented = new List<string>();

        RefreshSheetCommand.Run(orchestrator.Object, presented.Add);

        Assert.Equal("the chart could not be refreshed", Assert.Single(presented));
    }

    /// <summary>
    /// A presenter that throws degrades to no dialog rather than escaping into Excel.
    /// </summary>
    /// <remarks>
    /// The refresh has already happened by the time the presenter runs, so escaping
    /// would surface an error to the user for work that was done. This is the same
    /// containment the AddRow command applies.
    /// </remarks>
    [Fact]
    public void A_presenter_that_throws_does_not_escape_into_excel()
    {
        var orchestrator = new Mock<IGanttRefreshOrchestrator>();
        _ = orchestrator.Setup(o => o.Refresh()).Returns(Refused(GanttRefreshRefusal.TargetProtected));

        RefreshSheetCommand.Run(orchestrator.Object, _ => throw new InvalidOperationException("no UI"));
    }

    /// <summary>
    /// A partial refresh is presented, not swallowed.
    /// </summary>
    /// <remarks>
    /// This is the case where the sheet is in a state the user must know about, so
    /// suppressing the dialog because "the command usually stays quiet on success"
    /// would be exactly wrong. The command does not inspect the reason — it presents
    /// whatever the orchestrator reports — and this test pins that it does not
    /// special-case the partial one away.
    /// </remarks>
    [Fact]
    public void A_partial_refresh_is_presented_to_the_user()
    {
        var orchestrator = new Mock<IGanttRefreshOrchestrator>();
        _ = orchestrator
            .Setup(o => o.Refresh())
            .Returns(Refused(GanttRefreshRefusal.PartialReconciliation));
        var presented = new List<string>();

        RefreshSheetCommand.Run(orchestrator.Object, presented.Add);

        Assert.Single(presented);
    }

    /// <summary>Null arguments are refused at the boundary.</summary>
    [Fact]
    public void Run_refuses_null_arguments()
    {
        var orchestrator = new Mock<IGanttRefreshOrchestrator>();
        _ = orchestrator.Setup(o => o.Refresh()).Returns(Success());

        Assert.Throws<ArgumentNullException>(() => RefreshSheetCommand.Run(null!, _ => { }));
        Assert.Throws<ArgumentNullException>(() => RefreshSheetCommand.Run(orchestrator.Object, null!));
    }

    /// <summary>
    /// The command names no Excel interop type, so the AddIn's existing
    /// CS0433 duplicate-type hazard is not reintroduced.
    /// </summary>
    [Fact]
    public void The_command_source_names_no_interop_type()
    {
        string path = Path.Combine(
            FindRepositoryRoot(),
            "src/GanttCreator.AddIn/RefreshSheetCommand.cs");

        string source = File.ReadAllText(path);

        Assert.DoesNotContain("Microsoft.Office.Interop", source, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Interop.Excel", source, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// The workbook-state path repairs the outline symbols before it refreshes, and it is the
    /// ONLY path that does.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <c>EnableOutlining</c> is refused on a protected sheet and is not saved with the
    /// worksheet, so a reopened workbook comes up protected with the <c>+</c>/<c>-</c>
    /// symbols off, and only a protect run on an UNPROTECTED sheet restores them. The repair
    /// must precede the refresh it guards: a refresh can refuse early (an unreadable
    /// configuration, a blocked row) and would leave the user with greyed-out controls.
    /// </para>
    /// <para>
    /// <b>Why exactly one call site.</b> An earlier revision also repaired from
    /// <c>Activate</c>, to cover a workbook that was already open when the add-in attached.
    /// That call ran in the AddIn test host too, where <c>ExcelDnaUtil.Application</c>
    /// resolves the developer's live Excel, and the protect it performed polluted the
    /// <c>AutoOpen</c> log fixtures — 9 red tests. The reopen case is covered by this path
    /// plus the protect cycle a chart refresh performs; it must not come back.
    /// </para>
    /// <para>
    /// <b>Why a source assertion.</b> The repair is a static command over the live COM
    /// application, with no seam a unit test can substitute, so the wiring is asserted at
    /// the call site — the same technique
    /// <see cref="The_command_source_names_no_interop_type"/> already uses.
    /// </para>
    /// </remarks>
    [Fact]
    public void Only_the_workbook_state_path_repairs_the_outline_symbols()
    {
        string path = Path.Combine(
            FindRepositoryRoot(),
            "src/GanttCreator.AddIn/RibbonStateService.cs");

        string source = File.ReadAllText(path);
        const string repair = "RefreshSheetCommand.RepairOutlineSymbols();";

        int repairAt = source.IndexOf(repair, StringComparison.Ordinal);
        Assert.True(repairAt >= 0, "The workbook-state path no longer repairs the outline symbols.");
        int refreshAfter = source.IndexOf("Refresh();", repairAt, StringComparison.Ordinal);
        Assert.True(
            refreshAfter > repairAt,
            "The repair must precede the refresh, or an early refusal leaves the symbols off.");

        // Exactly one call site: a second one re-introduces the live-Excel write from Activate.
        Assert.Equal(
            -1,
            source.IndexOf(repair, repairAt + repair.Length, StringComparison.Ordinal));
    }

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (Directory.Exists(Path.Combine(directory.FullName, "src"))
                && File.Exists(Path.Combine(directory.FullName, "GanttCreator.slnx")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException(
            "Could not locate the repository root (a directory holding src/ and GanttCreator.slnx).");
    }
}
