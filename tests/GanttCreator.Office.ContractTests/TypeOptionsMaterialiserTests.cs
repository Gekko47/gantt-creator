using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>Positive refusal and zero-mutation tests for TypeOptions materialisation.</summary>
public class TypeOptionsMaterialiserTests
{
    [Fact]
    public void Materialise_refuses_without_an_active_workbook_before_reading_catalogue()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NoActiveWorkbook);
        var catalogue = new Mock<IConfigCatalogueReader>(MockBehavior.Strict);

        TypeOptionsMaterialiseOutcome outcome = new ExcelTypeOptionsMaterialiser(
            null, catalogue.Object, guard.Object).Materialise();

        Assert.Equal(TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.NoActiveWorkbook), outcome);
        catalogue.Verify(r => r.Read(), Times.Never);
    }

    [Fact]
    public void Materialise_refuses_protected_workbook_before_catalogue_read()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.SheetProtected);
        var catalogue = new Mock<IConfigCatalogueReader>(MockBehavior.Strict);

        TypeOptionsMaterialiseOutcome outcome = new ExcelTypeOptionsMaterialiser(
            null, catalogue.Object, guard.Object).Materialise();

        Assert.Equal(TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.TargetProtected), outcome);
        catalogue.Verify(r => r.Read(), Times.Never);
    }

    [Fact]
    public void Materialise_maps_missing_configuration_to_a_typed_refusal()
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var catalogue = new Mock<IConfigCatalogueReader>();
        _ = catalogue.Setup(r => r.Read()).Returns(ConfigReadOutcome.Refused(ConfigReadRefusalReason.ConfigSheetMissing));

        TypeOptionsMaterialiseOutcome outcome = new ExcelTypeOptionsMaterialiser(
            application.Object, catalogue.Object, guard.Object).Materialise();

        Assert.Equal(TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.ConfigSheetMissing), outcome);
    }

    [Fact]
    public void Materialise_maps_catalogue_hash_mismatch_to_a_typed_refusal()
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var catalogue = new Mock<IConfigCatalogueReader>();
        _ = catalogue.Setup(r => r.Read()).Returns(ConfigReadOutcome.Refused(ConfigReadRefusalReason.CatalogueHashMismatch));

        TypeOptionsMaterialiseOutcome outcome = new ExcelTypeOptionsMaterialiser(
            application.Object, catalogue.Object, guard.Object).Materialise();

        Assert.Equal(TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.CatalogueHashMismatch), outcome);
    }
}
