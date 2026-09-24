using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using GanttCreator.Core.ConfigIntegrity;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.Office.ContractTests;

public sealed class ConfigIntegrityCheckerTests
{
    [Fact]
    public void Check_refuses_without_an_active_workbook()
    {
        var outcome = new ExcelConfigIntegrityChecker(null, new Mock<IConfigCatalogueReader>().Object).Check();

        Assert.Equal(ConfigIntegrityCheckRefusalReason.NoActiveWorkbook, outcome.Refusal);
        Assert.Empty(outcome.Findings);
    }

    [Fact]
    public void Check_reports_a_missing_configuration_sheet_without_mutation()
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var sheets = new Mock<Excel.Sheets>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);
        _ = sheets.SetupGet(s => s.Count).Returns(0);

        ConfigIntegrityCheckOutcome outcome = new ExcelConfigIntegrityChecker(
            application.Object,
            new Mock<IConfigCatalogueReader>().Object).Check();

        Assert.Null(outcome.Refusal);
        Assert.Equal(ConfigIntegrityFindingKind.ConfigSheetMissing, Assert.Single(outcome.Findings).Kind);
    }

    [Fact]
    public void Check_refuses_an_unscoped_missing_catalogue_table()
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var sheets = new Mock<Excel.Sheets>();
        var config = new Mock<Excel.Worksheet>();
        var configObjects = new Mock<Excel.ListObjects>();
        var reader = new Mock<IConfigCatalogueReader>();
        var tableReader = new Mock<IGanttTableReader>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);
        _ = sheets.SetupGet(s => s.Count).Returns(1);
        _ = sheets.SetupGet(s => s[1]).Returns(config.Object);
        _ = config.SetupGet(s => s.Name).Returns(GanttWorkbookContract.ConfigSheetName);
        _ = config.SetupGet(s => s.Visible).Returns(Excel.XlSheetVisibility.xlSheetVeryHidden);
        _ = config.SetupGet(s => s.ListObjects).Returns(configObjects.Object);
        _ = configObjects.SetupGet(o => o.Count).Returns(0);
        _ = reader.Setup(r => r.Read()).Returns(ConfigReadOutcome.Refused(ConfigReadRefusalReason.TableMissing));
        _ = tableReader.Setup(r => r.Read()).Returns(GanttTableReadOutcome.Ok([]));

        ConfigIntegrityCheckOutcome outcome = new ExcelConfigIntegrityChecker(
            application.Object,
            reader.Object,
            tableReader.Object).Check();

        ConfigIntegrityFinding finding = Assert.Single(outcome.Findings);
        Assert.Equal(ConfigIntegrityFindingKind.CatalogueTableMissing, finding.Kind);
        Assert.Null(finding.TableName);
        Assert.Equal(ConfigRepairClassification.Refuse, ConfigIntegrityPlan.Classify(finding));
    }

    [Fact]
    public void Check_reports_wrong_visibility_and_catalogue_drift()
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var sheets = new Mock<Excel.Sheets>();
        var config = new Mock<Excel.Worksheet>();
        var configObjects = new Mock<Excel.ListObjects>();
        var reader = new Mock<IConfigCatalogueReader>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);
        _ = sheets.SetupGet(s => s.Count).Returns(1);
        _ = sheets.SetupGet(s => s[1]).Returns(config.Object);
        _ = config.SetupGet(s => s.Name).Returns(GanttWorkbookContract.ConfigSheetName);
        _ = config.SetupGet(s => s.Visible).Returns(Excel.XlSheetVisibility.xlSheetVisible);
        _ = config.SetupGet(s => s.ListObjects).Returns(configObjects.Object);
        _ = configObjects.SetupGet(o => o.Count).Returns(0);
        _ = reader.Setup(r => r.Read()).Returns(ConfigReadOutcome.Refused(ConfigReadRefusalReason.CatalogueHashMismatch));

        ConfigIntegrityCheckOutcome outcome = new ExcelConfigIntegrityChecker(application.Object, reader.Object).Check();

        Assert.Equal(
            [
                ConfigIntegrityFindingKind.WrongVisibility,
                ConfigIntegrityFindingKind.CatalogueHashMismatch,
            ],
            outcome.Findings.Select(finding => finding.Kind));
    }
}
