using GanttCreator.Core;
using GanttCreator.Core.ConfigIntegrity;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

public sealed class RepairConfigCommandTests
{
    private static ConfigIntegrityCheckOutcome Check(params ConfigIntegrityFinding[] findings) =>
        ConfigIntegrityCheckOutcome.Ok(findings);

    [Fact]
    public void Run_reports_nothing_to_repair_for_a_valid_workbook()
    {
        var checker = new Mock<IConfigIntegrityChecker>();
        _ = checker.Setup(c => c.Check()).Returns(Check());
        var repairer = new Mock<IConfigRepairer>(MockBehavior.Strict);
        var messages = new List<string>();

        RepairConfigCommand.Run(checker.Object, repairer.Object, _ => true, messages.Add);

        Assert.Equal(RepairConfigCommand.NothingToRepairMessage, Assert.Single(messages));
        repairer.Verify(r => r.Repair(It.IsAny<ConfigIntegrityPlan>(), It.IsAny<bool>()), Times.Never);
    }

    [Fact]
    public void Run_rebuilds_the_table_before_the_configuration_check_and_schedules_one_refresh()
    {
        var checker = new Mock<IConfigIntegrityChecker>();
        _ = checker.Setup(c => c.Check()).Returns(Check());
        var repairer = new Mock<IConfigRepairer>(MockBehavior.Strict);
        var recovery = new Mock<IWorkbookRecovery>();
        var inspection = new WorkbookRecoveryInspection(
            TableReadable: false,
            TablePresent: false,
            ConfigurationPresent: true,
            NeedsRecovery: true,
            TargetSheetName: "Gantt Data",
            BodyRowCount: 2,
            HeaderPlan: null,
            StructurePlan: null,
            RequiresConfirmation: true,
            ConfirmationMessage: "Adopt 2 rows?",
            Refusal: null);
        _ = recovery.Setup(r => r.Inspect()).Returns(inspection);
        _ = recovery.Setup(r => r.Recover(inspection, true))
            .Returns(new WorkbookRecoveryOutcome(true, true, false, 2, null));
        var messages = new List<string>();
        var refreshes = 0;

        RepairConfigCommand.Run(
            checker.Object,
            repairer.Object,
            _ => true,
            messages.Add,
            recovery.Object,
            () => refreshes++);

        recovery.Verify(r => r.Recover(inspection, true), Times.Once);
        checker.Verify(c => c.Check(), Times.Once);
        Assert.Contains("Rebuilt tblGanttData", Assert.Single(messages), StringComparison.Ordinal);
        Assert.Contains("adopted 2 row", messages[0], StringComparison.Ordinal);
        Assert.Equal(1, refreshes);
    }

    [Fact]
    public void Run_declines_the_rebuild_confirmation_and_mutates_nothing()
    {
        var checker = new Mock<IConfigIntegrityChecker>(MockBehavior.Strict);
        var repairer = new Mock<IConfigRepairer>(MockBehavior.Strict);
        var recovery = new Mock<IWorkbookRecovery>();
        var inspection = new WorkbookRecoveryInspection(
            TableReadable: false,
            TablePresent: false,
            ConfigurationPresent: true,
            NeedsRecovery: true,
            TargetSheetName: "Gantt Data",
            BodyRowCount: 1,
            HeaderPlan: null,
            StructurePlan: null,
            RequiresConfirmation: true,
            ConfirmationMessage: "Adopt 1 row?",
            Refusal: null);
        _ = recovery.Setup(r => r.Inspect()).Returns(inspection);
        var messages = new List<string>();

        RepairConfigCommand.Run(
            checker.Object,
            repairer.Object,
            _ => false,
            messages.Add,
            recovery.Object);

        recovery.Verify(r => r.Recover(It.IsAny<WorkbookRecoveryInspection>(), It.IsAny<bool>()), Times.Never);
        checker.Verify(c => c.Check(), Times.Never);
        Assert.Equal(RepairConfigCommand.CancelledMessage, Assert.Single(messages));
    }

    [Fact]
    public void Run_schedules_a_refresh_after_a_configuration_repair()
    {
        var checker = new Mock<IConfigIntegrityChecker>();
        _ = checker.Setup(c => c.Check()).Returns(Check(
            new ConfigIntegrityFinding(ConfigIntegrityFindingKind.IdentityDamage, "identity")));
        var repairer = new Mock<IConfigRepairer>();
        _ = repairer.Setup(r => r.Repair(It.IsAny<ConfigIntegrityPlan>(), true))
            .Returns(ConfigRepairOutcome.Ok(1, 0, 0));
        var messages = new List<string>();
        var refreshes = 0;

        RepairConfigCommand.Run(
            checker.Object,
            repairer.Object,
            _ => true,
            messages.Add,
            onChanged: () => refreshes++);

        Assert.Equal(1, refreshes);
        Assert.Contains("1 repaired", Assert.Single(messages), StringComparison.Ordinal);
    }

    [Fact]
    public void Run_declines_confirmation_and_stops_before_repair()
    {
        var checker = new Mock<IConfigIntegrityChecker>();
        _ = checker.Setup(c => c.Check()).Returns(Check(
            new ConfigIntegrityFinding(ConfigIntegrityFindingKind.IdentityDamage, "identity")));
        var repairer = new Mock<IConfigRepairer>();
        _ = repairer.Setup(r => r.Repair(It.IsAny<ConfigIntegrityPlan>(), false))
            .Returns(ConfigRepairOutcome.Ok(0, 1, 0));
        var confirmations = new List<string>();
        var messages = new List<string>();

        RepairConfigCommand.Run(
            checker.Object,
            repairer.Object,
            message => { confirmations.Add(message); return false; },
            messages.Add);

        Assert.Single(confirmations);
        Assert.Contains("1", confirmations[0], StringComparison.Ordinal);
        repairer.Verify(r => r.Repair(It.IsAny<ConfigIntegrityPlan>(), false), Times.Once);
        Assert.Contains("0 repaired", Assert.Single(messages), StringComparison.Ordinal);
    }

    [Fact]
    public void Run_confirms_and_repairs_confirm_required_findings()
    {
        var checker = new Mock<IConfigIntegrityChecker>();
        _ = checker.Setup(c => c.Check()).Returns(Check(
            new ConfigIntegrityFinding(ConfigIntegrityFindingKind.IdentityDamage, "identity")));
        var repairer = new Mock<IConfigRepairer>();
        _ = repairer.Setup(r => r.Repair(It.IsAny<ConfigIntegrityPlan>(), true))
            .Returns(ConfigRepairOutcome.Ok(1, 0, 0));
        var messages = new List<string>();

        RepairConfigCommand.Run(checker.Object, repairer.Object, _ => true, messages.Add);

        repairer.Verify(r => r.Repair(It.IsAny<ConfigIntegrityPlan>(), true), Times.Once);
        Assert.Contains("1 repaired", Assert.Single(messages), StringComparison.Ordinal);
    }

    [Fact]
    public void Run_refuses_before_repair_when_a_plan_contains_an_unsafe_finding()
    {
        var checker = new Mock<IConfigIntegrityChecker>();
        _ = checker.Setup(c => c.Check()).Returns(Check(
            new ConfigIntegrityFinding(ConfigIntegrityFindingKind.SecondHelperSheet, "extra")));
        var repairer = new Mock<IConfigRepairer>(MockBehavior.Strict);
        var messages = new List<string>();

        RepairConfigCommand.Run(checker.Object, repairer.Object, _ => true, messages.Add);

        repairer.Verify(r => r.Repair(It.IsAny<ConfigIntegrityPlan>(), It.IsAny<bool>()), Times.Never);
        Assert.Contains("1 refused", Assert.Single(messages), StringComparison.Ordinal);
    }

    [Fact]
    public void Run_surfaces_a_no_workbook_check_refusal()
    {
        var checker = new Mock<IConfigIntegrityChecker>();
        _ = checker.Setup(c => c.Check()).Returns(ConfigIntegrityCheckOutcome.NoActiveWorkbook());
        var repairer = new Mock<IConfigRepairer>(MockBehavior.Strict);
        var messages = new List<string>();

        RepairConfigCommand.Run(checker.Object, repairer.Object, _ => true, messages.Add);

        Assert.Contains("open workbook", Assert.Single(messages), StringComparison.OrdinalIgnoreCase);
        repairer.Verify(r => r.Repair(It.IsAny<ConfigIntegrityPlan>(), It.IsAny<bool>()), Times.Never);
    }

    [Fact]
    public void Run_rejects_null_dependencies()
    {
        var checker = new Mock<IConfigIntegrityChecker>();
        var repairer = new Mock<IConfigRepairer>();
        Assert.Throws<ArgumentNullException>(() => RepairConfigCommand.Run(null!, repairer.Object, _ => true, _ => { }));
        Assert.Throws<ArgumentNullException>(() => RepairConfigCommand.Run(checker.Object, null!, _ => true, _ => { }));
        Assert.Throws<ArgumentNullException>(() => RepairConfigCommand.Run(checker.Object, repairer.Object, null!, _ => { }));
        Assert.Throws<ArgumentNullException>(() => RepairConfigCommand.Run(checker.Object, repairer.Object, _ => true, null!));
    }
}
