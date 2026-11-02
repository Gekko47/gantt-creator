using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ExcelConfigSheetVisibilityRepairer"/>: the
/// guard-first ordering, every refusal, the case-insensitive contract-sheet match,
/// and the VeryHidden write.
/// </summary>
public sealed class ExcelConfigSheetVisibilityRepairerTests
{
    [Fact]
    public void Active_sheet_protection_refuses_before_any_workbook_read()
    {
        Fixture fixture = new(GanttWorkbookContract.ConfigSheetName);
        fixture.SetActiveSheet(ProtectionGuardOutcome.SheetProtected);

        ConfigSheetVisibilityRepairOutcome outcome = fixture.Repair();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ConfigSheetVisibilityRepairRefusalReason.TargetProtected, outcome.Refusal);
        fixture.VerifyNoWorkbookRead();
    }

    [Fact]
    public void An_absent_active_workbook_is_reported_as_no_active_workbook()
    {
        Fixture fixture = new(GanttWorkbookContract.ConfigSheetName);
        fixture.SetActiveSheet(ProtectionGuardOutcome.NoActiveWorkbook);

        ConfigSheetVisibilityRepairOutcome outcome = fixture.Repair();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ConfigSheetVisibilityRepairRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }

    [Fact]
    public void A_missing_application_is_refused_as_no_active_workbook()
    {
        Fixture fixture = new(GanttWorkbookContract.ConfigSheetName);

        ConfigSheetVisibilityRepairOutcome outcome = fixture.Repair(applicationPresent: false);

        Assert.False(outcome.Succeeded);
        Assert.Equal(ConfigSheetVisibilityRepairRefusalReason.NoActiveWorkbook, outcome.Refusal);
        fixture.VerifyNoWorkbookRead();
    }

    [Fact]
    public void A_workbook_without_an_active_sheet_owner_is_refused()
    {
        Fixture fixture = new(GanttWorkbookContract.ConfigSheetName);
        fixture.ForgetWorkbook();

        ConfigSheetVisibilityRepairOutcome outcome = fixture.Repair();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ConfigSheetVisibilityRepairRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }

    [Fact]
    public void A_workbook_without_the_configuration_sheet_is_refused()
    {
        Fixture fixture = new("Gantt", "Another Sheet");

        ConfigSheetVisibilityRepairOutcome outcome = fixture.Repair();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ConfigSheetVisibilityRepairRefusalReason.ConfigSheetMissing, outcome.Refusal);
    }

    [Fact]
    public void The_configuration_sheet_is_matched_case_insensitively_and_made_very_hidden()
    {
        // Derived, not re-typed: Excel preserves the spelling it was given, and a
        // hand-typed upper-case copy is one silent typo away from testing nothing.
        string upperInvariant = GanttWorkbookContract.ConfigSheetName.ToUpperInvariant();
        Fixture fixture = new("Gantt", upperInvariant, "Another Sheet");

        ConfigSheetVisibilityRepairOutcome outcome = fixture.Repair();

        Assert.True(outcome.Succeeded);
        Assert.Null(outcome.Refusal);
        Assert.Equal(Excel.XlSheetVisibility.xlSheetVeryHidden, fixture.VisibleOf(1));
        Assert.NotEqual(Excel.XlSheetVisibility.xlSheetVeryHidden, fixture.VisibleOf(0));
    }

    [Fact]
    public void A_protected_configuration_sheet_is_refused_as_target_protected()
    {
        Fixture fixture = new(GanttWorkbookContract.ConfigSheetName);
        fixture.ProtectTarget(ProtectionGuardOutcome.WorkbookStructureProtected);

        ConfigSheetVisibilityRepairOutcome outcome = fixture.Repair();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ConfigSheetVisibilityRepairRefusalReason.TargetProtected, outcome.Refusal);
        Assert.NotEqual(Excel.XlSheetVisibility.xlSheetVeryHidden, fixture.VisibleOf(0));
    }

    [Fact]
    public void A_target_guard_without_a_workbook_is_refused_as_no_active_workbook()
    {
        Fixture fixture = new(GanttWorkbookContract.ConfigSheetName);
        fixture.ProtectTarget(ProtectionGuardOutcome.NoActiveWorkbook);

        ConfigSheetVisibilityRepairOutcome outcome = fixture.Repair();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ConfigSheetVisibilityRepairRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }

    /// <summary>The repairer with the COM sheet-indexer seam substituted.</summary>
    private sealed class TestableRepairer(
        object? application,
        IWorksheetProtectionGuard guard,
        IReadOnlyList<Excel.Worksheet> sheets)
        : ExcelConfigSheetVisibilityRepairer(application, guard)
    {
        internal override Excel.Worksheet GetSheetAt(Excel.Sheets sheetCollection, int index) =>
            sheets[index - 1];
    }

    /// <summary>The mocked workbook graph and the guard.</summary>
    private sealed class Fixture
    {
        private readonly List<Mock<Excel.Worksheet>> _sheets = [];

        public Fixture(params string[] sheetNames)
        {
            foreach (string name in sheetNames)
            {
                Mock<Excel.Worksheet> sheet = new();
                _ = sheet.SetupProperty(s => s.Visible);
                _ = sheet.SetupGet(s => s.Name).Returns(name);
                _sheets.Add(sheet);
            }

            _ = Sheets.SetupGet(s => s.Count).Returns(_sheets.Count);
            _ = Workbook.SetupGet(w => w.Sheets).Returns(Sheets.Object);
            _ = Application.SetupGet(a => a.ActiveWorkbook).Returns(Workbook.Object);
            _ = Guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
            _ = Guard.Setup(g => g.QueryTarget(It.IsAny<object?>()))
                .Returns(ProtectionGuardOutcome.NotProtected);
        }

        private Mock<Excel.Application> Application { get; } = new();

        private Mock<Excel.Workbook> Workbook { get; } = new();

        private Mock<Excel.Sheets> Sheets { get; } = new();

        private Mock<IWorksheetProtectionGuard> Guard { get; } = new();

        /// <summary>Runs the repairer.</summary>
        /// <param name="applicationPresent">Whether the host supplied an application.</param>
        /// <returns>The outcome.</returns>
        public ConfigSheetVisibilityRepairOutcome Repair(bool applicationPresent = true) =>
            new TestableRepairer(
                applicationPresent ? Application.Object : null,
                Guard.Object,
                [.. _sheets.Select(sheet => sheet.Object)]).Repair();

        /// <summary>The visibility of a sheet, by its position in the workbook.</summary>
        /// <param name="index">The zero-based position.</param>
        /// <returns>The reported visibility.</returns>
        public Excel.XlSheetVisibility VisibleOf(int index) => _sheets[index].Object.Visible;

        /// <summary>Asserts the workbook was never read.</summary>
        public void VerifyNoWorkbookRead() =>
            Application.VerifyGet(a => a.ActiveWorkbook, Times.Never);

        /// <summary>Makes the active-sheet query report the supplied outcome.</summary>
        /// <param name="outcome">The guard outcome.</param>
        public void SetActiveSheet(ProtectionGuardOutcome outcome) =>
            _ = Guard.Setup(g => g.Query()).Returns(outcome);

        /// <summary>Makes the target-sheet query report the supplied outcome.</summary>
        /// <param name="outcome">The guard outcome.</param>
        public void ProtectTarget(ProtectionGuardOutcome outcome) =>
            _ = Guard.Setup(g => g.QueryTarget(It.IsAny<object?>())).Returns(outcome);

        /// <summary>Removes the active workbook.</summary>
        public void ForgetWorkbook() => _ = Application
            .SetupGet(a => a.ActiveWorkbook)
            .Returns((Excel.Workbook)null!);
    }
}
