using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using GanttCreator.Core.ConfigIntegrity;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Depth tests for <see cref="ExcelConfigIntegrityChecker"/>: the duplicate-sheet
/// scan, the workbook-name checks, and the catalogue-refusal mapping.
/// </summary>
/// <remarks>
/// <see cref="ConfigIntegrityCheckerTests"/> covers the missing-configuration-sheet
/// and catalogue-drift paths. This file exercises the branches that need a
/// resolvable Gantt table and a live defined-name collection, which is the half
/// the checker exists for.
/// </remarks>
public sealed class ConfigIntegrityCheckerDepthTests
{
    /// <summary>The checker with the sheet indexer substituted, so the scans run.</summary>
    private sealed class TestableChecker(
        Excel.Application application,
        IConfigCatalogueReader catalogueReader,
        IGanttTableReader tableReader,
        IReadOnlyList<Excel._Worksheet?> sheetObjects)
        : ExcelConfigIntegrityChecker(application, catalogueReader, tableReader)
    {
        internal override Excel.Worksheet GetSheetAt(Excel.Sheets sheets, int index) =>
            (Excel.Worksheet)sheetObjects[index - 1]!;
    }

    private const string GanttSheetName = "Gantt";

    private static Mock<Excel.Name> DefinedName(string name, string refersTo)
    {
        Mock<Excel.Name> defined = new();
        _ = defined.SetupGet(n => n.Name).Returns(name);
        _ = defined.SetupGet(n => n.RefersTo).Returns(refersTo);
        return defined;
    }

    private static Mock<Excel.Names> NamesCollection(params Excel.Name[] entries)
    {
        Mock<Excel.Names> names = new();
        _ = names.SetupGet(n => n.Count).Returns(entries.Length);
        for (int index = 1; index <= entries.Length; index++)
        {
            Excel.Name entry = entries[index - 1];
            _ = names.Setup(n => n.Item(It.IsAny<object>())).Returns(entry);
        }

        return names;
    }

    private static Excel.Worksheet ConfigSheet(Excel.XlSheetVisibility visibility = Excel.XlSheetVisibility.xlSheetVeryHidden)
    {
        // An empty table collection: the Gantt-table scan walks EVERY sheet, so a
        // configuration sheet reporting no ListObjects at all makes that scan throw
        // rather than skip it.
        Mock<Excel.ListObjects> tables = new();
        _ = tables.SetupGet(o => o.Count).Returns(0);
        Mock<Excel.Worksheet> sheet = new();
        _ = sheet.SetupGet(s => s.Name).Returns(GanttWorkbookContract.ConfigSheetName);
        _ = sheet.SetupGet(s => s.Visible).Returns(visibility);
        _ = sheet.SetupGet(s => s.ListObjects).Returns(tables.Object);
        return sheet.Object;
    }

    private static Excel.Worksheet GanttSheet(Excel.Names? names = null)
    {
        Mock<Excel.ListObject> table = new();
        _ = table.SetupGet(t => t.Name).Returns(GanttTableSchema.TableName);
        Mock<Excel.ListObjects> tables = new();
        _ = tables.SetupGet(t => t.Count).Returns(1);
        _ = tables.Setup(t => t[It.IsAny<int>()]).Returns(table.Object);
        Mock<Excel.Worksheet> sheet = new();
        _ = sheet.SetupGet(s => s.Name).Returns(GanttSheetName);
        _ = sheet.SetupGet(s => s.ListObjects).Returns(tables.Object);

        // An empty collection rather than a null one: the checker iterates the
        // sheet's defined names, and a null collection makes that read throw
        // instead of reporting "the anchor is missing".
        _ = sheet.SetupGet(s => s.Names).Returns(names ?? NamesCollection().Object);
        return sheet.Object;
    }

    private static ConfigIntegrityCheckOutcome Check(
        IReadOnlyList<Excel._Worksheet?> sheets,
        Excel.Names? workbookNames = null,
        ConfigReadOutcome? catalogue = null)
    {
        Mock<Excel.Sheets> collection = new();
        _ = collection.SetupGet(s => s.Count).Returns(sheets.Count);
        Mock<Excel.Workbook> workbook = new();
        _ = workbook.SetupGet(w => w.Sheets).Returns(collection.Object);
        _ = workbook.SetupGet(w => w.Names).Returns(workbookNames ?? NamesCollection().Object);

        Mock<Excel.Application> application = new();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);

        Mock<IConfigCatalogueReader> reader = new();
        _ = reader.Setup(r => r.Read())
            .Returns(catalogue ?? ConfigReadOutcome.Refused(ConfigReadRefusalReason.TableMissing));
        Mock<IGanttTableReader> tableReader = new();
        _ = tableReader.Setup(r => r.Read())
            .Returns(GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.TableMissing));

        return new TestableChecker(application.Object, reader.Object, tableReader.Object, sheets).Check();
    }

    [Fact]
    public void A_second_worksheet_carrying_the_configuration_name_is_reported()
    {
        // Two distinct worksheets with the name: the second is a helper sheet the
        // contract forbids, so it must be reported rather than silently used.
        ConfigIntegrityCheckOutcome outcome = Check([ConfigSheet(), ConfigSheet(), GanttSheet(names: null)]);

        Assert.Contains(
            outcome.Findings,
            finding => finding.Kind == ConfigIntegrityFindingKind.SecondHelperSheet);
    }

    [Fact]
    public void The_same_worksheet_resolved_twice_is_not_a_second_helper_sheet()
    {
        // Identity, not name: the scan must compare instances, or resolving one
        // sheet at two indices would report a duplicate that does not exist.
        Excel.Worksheet config = ConfigSheet();
        ConfigIntegrityCheckOutcome outcome = Check([config, config]);

        Assert.DoesNotContain(
            outcome.Findings,
            finding => finding.Kind == ConfigIntegrityFindingKind.SecondHelperSheet);
    }

    [Fact]
    public void A_missing_plot_anchor_defined_name_is_reported_as_a_disagreement()
    {
        Excel.Worksheet gantt = GanttSheet(names: NamesCollection().Object);

        ConfigIntegrityCheckOutcome outcome = Check(
            [ConfigSheet(), gantt],
            workbookNames: NamesCollection().Object);

        Assert.Contains(
            outcome.Findings,
            finding => finding.Kind == ConfigIntegrityFindingKind.PlotAnchorDisagreement);
    }

    [Fact]
    public void An_anchor_pointing_at_another_sheet_is_reported_as_a_disagreement()
    {
        Excel.Worksheet gantt = GanttSheet(
            names: NamesCollection(
                DefinedName(GanttWorkbookContract.PlotAnchorDefinedName, "=SomeOtherSheet!$B$2").Object).Object);

        ConfigIntegrityCheckOutcome outcome = Check(
            [ConfigSheet(), gantt],
            workbookNames: NamesCollection(
                DefinedName(GanttWorkbookContract.TypeOptionsDefinedName, "={1,2}").Object).Object);

        Assert.Contains(
            outcome.Findings,
            finding => finding.Kind == ConfigIntegrityFindingKind.PlotAnchorDisagreement);
    }

    [Fact]
    public void A_matching_anchor_and_a_present_type_options_name_agree()
    {
        string expected = GanttSheetLayout.BuildPlotAnchorRefersTo(GanttSheetName);
        Excel.Worksheet gantt = GanttSheet(
            names: NamesCollection(
                DefinedName(GanttWorkbookContract.PlotAnchorDefinedName, expected).Object).Object);

        ConfigIntegrityCheckOutcome outcome = Check(
            [ConfigSheet(), gantt],
            workbookNames: NamesCollection(
                DefinedName(GanttWorkbookContract.TypeOptionsDefinedName, "={1,2}").Object).Object);

        Assert.DoesNotContain(
            outcome.Findings,
            finding => finding.Kind == ConfigIntegrityFindingKind.PlotAnchorDisagreement);
        Assert.DoesNotContain(
            outcome.Findings,
            finding => finding.Kind == ConfigIntegrityFindingKind.TypeOptionsMissing);
    }

    [Fact]
    public void A_missing_type_options_defined_name_is_reported()
    {
        string expected = GanttSheetLayout.BuildPlotAnchorRefersTo(GanttSheetName);
        Excel.Worksheet gantt = GanttSheet(
            names: NamesCollection(
                DefinedName(GanttWorkbookContract.PlotAnchorDefinedName, expected).Object).Object);

        ConfigIntegrityCheckOutcome outcome = Check(
            [ConfigSheet(), gantt],
            workbookNames: NamesCollection().Object);

        Assert.Contains(
            outcome.Findings,
            finding => finding.Kind == ConfigIntegrityFindingKind.TypeOptionsMissing);
    }

    /// <summary>
    /// Every catalogue refusal maps to its own finding kind. The mapping is the
    /// checker's whole testimony about why the catalogue is unusable, so a default
    /// arm swallowing a new reason would hide it.
    /// </summary>
    public static TheoryData<ConfigReadRefusalReason, ConfigIntegrityFindingKind> RefusalMappings
    {
        get
        {
            TheoryData<ConfigReadRefusalReason, ConfigIntegrityFindingKind> data = [];
            data.Add(ConfigReadRefusalReason.NoActiveWorkbook, ConfigIntegrityFindingKind.Unknown);
            data.Add(ConfigReadRefusalReason.ConfigSheetMissing, ConfigIntegrityFindingKind.ConfigSheetMissing);
            data.Add(ConfigReadRefusalReason.CatalogueHashMismatch, ConfigIntegrityFindingKind.CatalogueHashMismatch);
            data.Add(ConfigReadRefusalReason.TableMissing, ConfigIntegrityFindingKind.CatalogueTableMissing);
            data.Add(ConfigReadRefusalReason.HeaderMismatch, ConfigIntegrityFindingKind.CatalogueTableCorrupt);
            data.Add(ConfigReadRefusalReason.RowCountMismatch, ConfigIntegrityFindingKind.CatalogueTableCorrupt);
            data.Add(ConfigReadRefusalReason.CatalogueMismatch, ConfigIntegrityFindingKind.CatalogueTableCorrupt);
            data.Add(ConfigReadRefusalReason.ValueOutOfRange, ConfigIntegrityFindingKind.CatalogueTableCorrupt);
            data.Add(ConfigReadRefusalReason.BadColourFormat, ConfigIntegrityFindingKind.CatalogueTableCorrupt);
            return data;
        }
    }

    [Theory]
    [MemberData(nameof(RefusalMappings))]
    public void Each_catalogue_refusal_maps_to_its_own_finding_kind(
        ConfigReadRefusalReason refusal,
        ConfigIntegrityFindingKind expected)
    {
        ConfigIntegrityCheckOutcome outcome = Check(
            [ConfigSheet()],
            catalogue: ConfigReadOutcome.Refused(refusal));

        Assert.Contains(
            outcome.Findings,
            finding => finding.Kind == expected && finding.Detail.Contains(refusal.ToString(), StringComparison.Ordinal));
    }
}
