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

    /// <summary>
    /// The name-target guard must accept the value Excel echoes back for a name
    /// this adapter wrote, and must still refuse a name pointing elsewhere.
    /// </summary>
    [Theory]
    // Excel strips the quotes from a sheet name that does not need them, so the
    // adapter's own quoted value is read back unquoted. Accepting that form is the
    // regression: refusing it made every Add Row report TypeOptionsUnavailable.
    [InlineData("=_GanttCreatorConfig!$B$2:$B$17", "='_GanttCreatorConfig'!$B$2:$B$17", true)]
    [InlineData("='_GanttCreatorConfig'!$B$2:$B$17", "='_GanttCreatorConfig'!$B$2:$B$17", true)]
    [InlineData("=_ganttcreatorconfig!$B$2:$B$17", "='_GanttCreatorConfig'!$B$2:$B$17", true)]
    // A sheet name containing a quote is stored escaped and read back escaped.
    [InlineData("='It''s Config'!$B$2", "='It''s Config'!$B$2", true)]
    // A real mismatch is still refused.
    [InlineData("=_GanttCreatorConfig!$B$3:$B$17", "='_GanttCreatorConfig'!$B$2:$B$17", false)]
    [InlineData("='Gantt Data'!$B$2:$B$17", "='_GanttCreatorConfig'!$B$2:$B$17", false)]
    [InlineData("=Sheet1!$B$2", "='_GanttCreatorConfig'!$B$2:$B$17", false)]
    [InlineData("no-separator", "='_GanttCreatorConfig'!$B$2:$B$17", false)]
    [InlineData("", "='_GanttCreatorConfig'!$B$2:$B$17", false)]
    public void RefersTo_comparison_accepts_excel_normalisation_and_refuses_a_different_target(
        string existing,
        string expected,
        bool expectedMatch)
    {
        Assert.Equal(expectedMatch, ExcelTypeOptionsMaterialiser.RefersToMatches(existing, expected));
    }

    [Fact]
    public void RefersTo_comparison_refuses_a_missing_stored_target()
    {
        Assert.False(ExcelTypeOptionsMaterialiser.RefersToMatches(null, "='_GanttCreatorConfig'!$B$2:$B$17"));
    }
}
