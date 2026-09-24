using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Shared graph for the catalogue tests: the fake-backed
/// <c>TestableWriter</c>/<c>TestableReader</c> seam substitutions and the
/// builders that materialise the real writer payloads as in-memory
/// fixtures. Both test classes use this helper so the write→read
/// round-trip runs through the real production code with only COM seams
/// substituted (no live Office).
/// </summary>
internal static class ConfigGraph
{
    private sealed class TestableWriter : ExcelConfigCatalogueWriter
    {
        public TestableWriter(object? application, ConfigSheetFake fake, IWorksheetProtectionGuard? protectionGuard = null)
            : base(application, protectionGuard)
        {
            Fake = fake;
        }

        public ConfigSheetFake Fake { get; }

        internal override object GetSheetAt(Excel.Sheets sheets, int index) =>
            index == 1 ? Fake.Worksheet : throw new InvalidOperationException("Only index 1 is configured.");

        internal override Excel.ListObjects GetListObjects(Excel.Worksheet worksheet) =>
            Fake.ListObjects;

        internal override Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index) =>
            Fake.TableAt(index);

        internal override Excel.ListObject AddTable(Excel.ListObjects listObjects, object source) =>
            Fake.Add((Excel.Range)source);

        internal override Excel.Range GetCellRange(Excel.Worksheet worksheet, int row, int column) =>
            Fake.CellAt(ToAddress(row, column));

        internal override Excel.Range GetResizedRange(Excel.Range range, int rows, int columns) =>
            ConfigSheetFake.CreateExtent(rows, columns);

        internal override Excel.Range? GetTableBody(Excel.ListObject table) =>
            Fake.BodyOf(table);

        internal override object? GetBodyValues(Excel.Range body) =>
            ConfigSheetFake.ValuesOf(body);
    }

    private sealed class TestableReader : ExcelConfigCatalogueReader
    {
        public TestableReader(object? application, ConfigSheetFake fake)
            : base(application)
        {
            Fake = fake;
        }

        public ConfigSheetFake Fake { get; }

        internal override object GetSheetAt(Excel.Sheets sheets, int index) =>
            index == 1 ? Fake.Worksheet : throw new InvalidOperationException("Only index 1 is configured.");

        internal override Excel.ListObjects GetListObjects(Excel.Worksheet worksheet) =>
            Fake.ListObjects;

        internal override Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index) =>
            Fake.TableAt(index);

        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index) =>
            Fake.ColumnAt(columns, index);

        internal override Excel.Range? GetTableBody(Excel.ListObject table) =>
            Fake.BodyOf(table);

        internal override object? GetBodyValues(Excel.Range body) =>
            ConfigSheetFake.ValuesOf(body);
    }

    /// <summary>
    /// Builds a writer whose workbook graph contains exactly the config
    /// sheet fake (at <c>Sheets[1]</c>), with protection flags as given.
    /// </summary>
    /// <param name="fake">The config sheet fake.</param>
    /// <param name="structureProtected">The mocked <c>Workbook.ProtectStructure</c>.</param>
    /// <returns>The writer under test.</returns>
    public static ExcelConfigCatalogueWriter BuildWriter(
        ConfigSheetFake fake,
        bool structureProtected = false,
        IWorksheetProtectionGuard? protectionGuard = null)
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var sheets = new Mock<Excel.Sheets>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);
        _ = workbook.SetupGet(w => w.ActiveSheet).Returns(fake.Worksheet);
        _ = workbook.SetupGet(w => w.ProtectStructure).Returns(structureProtected);
        _ = sheets.SetupGet(s => s.Count).Returns(1);
        _ = application.SetupGet(a => a.DisplayAlerts).Returns(true);
        return new TestableWriter(application.Object, fake, protectionGuard);
    }

    /// <summary>
    /// Builds a reader whose workbook graph contains exactly the config
    /// sheet fake (at <c>Sheets[1]</c>).
    /// </summary>
    /// <param name="fake">The config sheet fake.</param>
    /// <returns>The reader under test.</returns>
    public static ExcelConfigCatalogueReader BuildReader(ConfigSheetFake fake)
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var sheets = new Mock<Excel.Sheets>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);
        _ = sheets.SetupGet(s => s.Count).Returns(1);
        return new TestableReader(application.Object, fake);
    }

    /// <summary>
    /// Converts a one-based (row, column) pair to an A1-style address.
    /// </summary>
    /// <param name="row">The one-based row.</param>
    /// <param name="column">The one-based column.</param>
    /// <returns>The address.</returns>
    public static string ToAddress(int row, int column)
    {
        var builder = new System.Text.StringBuilder();
        var remaining = column;
        while (remaining > 0)
        {
            var digit = (remaining - 1) % 26;
            _ = builder.Insert(0, (char)('A' + digit));
            remaining = (remaining - 1) / 26;
        }

        _ = builder.Append(row.ToString(System.Globalization.CultureInfo.InvariantCulture));
        return builder.ToString();
    }
}
