using GanttCreator.Core;
using GanttCreator.Office;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

public sealed class DeferredRepairAdapterTests
{
    [Fact]
    public void Visibility_repair_refuses_protection_before_workbook_access()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.SheetProtected);
        var application = new Mock<Microsoft.Office.Interop.Excel.Application>(MockBehavior.Strict);

        ConfigSheetVisibilityRepairOutcome outcome = new ExcelConfigSheetVisibilityRepairer(
            application.Object,
            guard.Object).Repair();

        Assert.Equal(ConfigSheetVisibilityRepairRefusalReason.TargetProtected, outcome.Refusal);
        application.VerifyGet(a => a.ActiveWorkbook, Times.Never);
    }

    [Fact]
    public void Visibility_repair_refuses_without_an_active_workbook()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NoActiveWorkbook);

        ConfigSheetVisibilityRepairOutcome outcome = new ExcelConfigSheetVisibilityRepairer(
            null,
            guard.Object).Repair();

        Assert.Equal(ConfigSheetVisibilityRepairRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }

    [Fact]
    public void Plot_anchor_repair_refuses_without_a_supported_gantt_table()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var application = new Mock<Microsoft.Office.Interop.Excel.Application>();
        var workbook = new Mock<Microsoft.Office.Interop.Excel.Workbook>();
        var sheets = new Mock<Microsoft.Office.Interop.Excel.Sheets>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);
        _ = sheets.SetupGet(s => s.Count).Returns(0);

        PlotAnchorRepairOutcome outcome = new ExcelPlotAnchorRepairer(application.Object, guard.Object).Repair();

        Assert.Equal(PlotAnchorRepairRefusalReason.TableMissing, outcome.Refusal);
    }

    [Fact]
    public void Identity_repair_refuses_the_resolved_target_when_it_is_protected()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard
            .Setup(g => g.QueryTarget(It.IsAny<object>()))
            .Returns(ProtectionGuardOutcome.SheetProtected);
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var sheets = new Mock<Excel.Sheets>();
        var worksheet = new Mock<Excel.Worksheet>();
        var objects = new Mock<Excel.ListObjects>();
        var table = new Mock<Excel.ListObject>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);
        _ = sheets.SetupGet(s => s.Count).Returns(1);
        _ = sheets.Setup(s => s[1]).Returns(worksheet.Object);
        _ = worksheet.SetupGet(w => w.ListObjects).Returns(objects.Object);
        _ = objects.SetupGet(o => o.Count).Returns(1);
        _ = objects.Setup(o => o[1]).Returns(table.Object);
        _ = table.SetupGet(t => t.Name).Returns(GanttTableSchema.TableName);

        GanttRowIdentityRepairOutcome outcome = new ExcelGanttRowIdentityRepairer(
            application.Object,
            guard.Object).Repair();

        Assert.Equal(GanttRowIdentityRepairRefusalReason.TargetProtected, outcome.Refusal);
        guard.Verify(g => g.QueryTarget(worksheet.Object), Times.Once);
        table.VerifyGet(t => t.DataBodyRange, Times.Never);
    }
}
