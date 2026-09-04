namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Demonstrates the fake-adapter pattern used in Office.ContractTests.
/// The real Office adapters implement the same narrow ports so tests
/// can run without a live Office installation.
/// </summary>
public class FakeAdapterTests
{
    // A minimal port that represents the part of the Office object model
    // that the domain code actually depends on.
    private interface IWorksheet
    {
        string Name { get; }
        object? GetCellValue(int row, int col);
        void SetCellValue(int row, int col, object value);
    }

    // A concrete fake that exercises the port.
    private sealed class FakeWorksheet : IWorksheet
    {
        private readonly Dictionary<(int Row, int Col), object?> _cells = new();

        public string Name { get; init; } = "Sheet1";

        public object? GetCellValue(int row, int col)
            => _cells.TryGetValue((row, col), out var v) ? v : null;

        public void SetCellValue(int row, int col, object value)
            => _cells[(row, col)] = value;
    }

    [Fact]
    public void FakeWorksheet_SetAndGet_roundtrips()
    {
        var sheet = new FakeWorksheet { Name = "Data" };

        sheet.SetCellValue(1, 1, "As-Built Activity");
        sheet.SetCellValue(1, 2, new DateTime(2025, 3, 1));
        sheet.SetCellValue(1, 3, new DateTime(2025, 6, 15));

        Assert.Equal("As-Built Activity", sheet.GetCellValue(1, 1));
        Assert.Equal(new DateTime(2025, 3, 1), sheet.GetCellValue(1, 2));
        Assert.Equal(new DateTime(2025, 6, 15), sheet.GetCellValue(1, 3));
    }

    [Fact]
    public void FakeWorksheet_GetCellValue_returns_null_for_unset()
    {
        var sheet = new FakeWorksheet();
        Assert.Null(sheet.GetCellValue(5, 5));
    }

    [Fact]
    public void FakeWorksheet_Name_defaults_to_Sheet1()
    {
        var sheet = new FakeWorksheet();
        Assert.Equal("Sheet1", sheet.Name);
    }

    [Fact]
    public void FakeWorksheet_Name_can_be_custom()
    {
        var sheet = new FakeWorksheet { Name = "GanttData" };
        Assert.Equal("GanttData", sheet.Name);
    }
}
