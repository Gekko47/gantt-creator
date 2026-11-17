using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

public class ConfigSheetInspectorTests
{
    private sealed class SheetGraph
    {
        public SheetGraph(string name, Excel.XlSheetVisibility visibility, params string[] tables)
        {
            Name = name;
            Visibility = visibility;
            TableMocks =
            [
                .. tables.Select(tableName =>
                {
                    Mock<Excel.ListObject> table = new();
                    _ = table.SetupGet(t => t.Name).Returns(tableName);
                    return table;
                }),
            ];

            _ = ListObjects.SetupGet(l => l.Count).Returns(tables.Length);
            _ = Worksheet.SetupGet(w => w.ListObjects).Returns(ListObjects.Object);
        }

        public string Name { get; }

        public Excel.XlSheetVisibility Visibility { get; }

        public Mock<Excel.Worksheet> Worksheet { get; } = new();

        public Mock<Excel.ListObjects> ListObjects { get; } = new();

        public IReadOnlyList<Mock<Excel.ListObject>> TableMocks { get; }
    }

    private sealed class TestableInspector(IReadOnlyList<object> sheets, IReadOnlyList<SheetGraph> graphs)
        : ExcelConfigSheetInspector
    {
        internal override int GetSheetCount(Excel.Sheets sheetsObject) => sheets.Count;

        internal override object GetSheetAt(Excel.Sheets sheetsObject, int index) => sheets[index - 1];

        internal override string GetSheetName(Excel.Worksheet worksheet) =>
            Graph(worksheet).Name;

        internal override Excel.XlSheetVisibility GetSheetVisibility(Excel.Worksheet worksheet) =>
            Graph(worksheet).Visibility;

        internal override Excel.ListObject GetTableAt(Excel.ListObjects objects, int index) =>
            graphs.First(g => ReferenceEquals(g.ListObjects.Object, objects)).TableMocks[index - 1].Object;

        private SheetGraph Graph(Excel.Worksheet worksheet) =>
            graphs.First(g => ReferenceEquals(g.Worksheet.Object, worksheet));
    }

    private static (TestableInspector Inspector, Mock<Excel.Sheets> Sheets) Build(params SheetGraph[] graphs)
    {
        Mock<Excel.Sheets> sheets = new();
        _ = sheets.SetupGet(s => s.Count).Returns(graphs.Length);
        for (int index = 1; index <= graphs.Length; index++)
        {
            _ = sheets.Setup(s => s[index]).Returns(graphs[index - 1].Worksheet.Object);
        }

        return (
            new TestableInspector([.. graphs.Select(g => (object)g.Worksheet.Object)], graphs),
            sheets);
    }

    [Fact]
    public void The_inspector_reports_a_stripped_very_hidden_canonical_sheet_as_present()
    {
        SheetGraph sheet = new(
            GanttWorkbookContract.ConfigSheetName,
            Excel.XlSheetVisibility.xlSheetVeryHidden);
        (TestableInspector inspector, Mock<Excel.Sheets> sheets) = Build(sheet);

        ConfigSheetLocation location = inspector.Locate(sheets.Object);

        Assert.Equal(ConfigSheetState.Present, location.State);
    }

    [Fact]
    public void The_inspector_reports_a_renamed_very_hidden_sheet_with_a_contract_table_as_renamed()
    {
        SheetGraph sheet = new(
            "ConfigBackup",
            Excel.XlSheetVisibility.xlSheetVeryHidden,
            GanttCatalogues.TypesTableName);
        (TestableInspector inspector, Mock<Excel.Sheets> sheets) = Build(sheet);

        ConfigSheetLocation location = inspector.Locate(sheets.Object);

        Assert.Equal(ConfigSheetState.Renamed, location.State);
        Assert.Equal("ConfigBackup", location.SheetName);
    }

    [Fact]
    public void The_inspector_reports_a_user_sheet_holding_the_name_as_a_collision()
    {
        SheetGraph userSheet = new("_GanttCreatorConfig", Excel.XlSheetVisibility.xlSheetVisible);
        (TestableInspector inspector, Mock<Excel.Sheets> sheets) = Build(userSheet);

        ConfigSheetLocation location = inspector.Locate(sheets.Object);

        Assert.Equal(ConfigSheetState.NameCollision, location.State);
    }

    [Fact]
    public void The_inspector_does_not_count_a_non_contract_table_as_a_contract_table()
    {
        SheetGraph sheet = new(
            "ConfigBackup",
            Excel.XlSheetVisibility.xlSheetVeryHidden,
            "tblUserData");
        (TestableInspector inspector, Mock<Excel.Sheets> sheets) = Build(sheet);

        ConfigSheetLocation location = inspector.Locate(sheets.Object);

        Assert.Equal(ConfigSheetState.Unrecognised, location.State);
    }

    [Fact]
    public void The_inspector_ignores_a_sheet_that_is_not_a_worksheet()
    {
        (TestableInspector inspector, Mock<Excel.Sheets> sheets) = Build();

        ConfigSheetLocation location = inspector.Locate(sheets.Object);

        Assert.Equal(ConfigSheetState.Missing, location.State);
    }
}
