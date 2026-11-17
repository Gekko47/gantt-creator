using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for the visible-column width default Initialise applies (R5.21,
/// owner ruling 2026-11-17): every user-visible column reaches 60 pixels (45pt) through
/// the same bounded characters-to-points conversion the gutter uses, and an engine column
/// is never touched.
/// </summary>
public sealed class VisibleColumnWidthTests
{
    private const double _liveCharacters = 8.43;
    private const double _livePoints = 48.0;

    private sealed class TestableInitialiser : ExcelWorkbookInitialiser
    {
        public TestableInitialiser()
            : base(null)
        {
        }

        public List<double> WidthWrites { get; } = [];

        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index) =>
            columns[index];

        internal override Excel.Range GetEntireColumn(Excel.Range range) => range;

        internal override double? ReadColumnWidthCharacters(Excel.Range column) => _liveCharacters;

        internal override double? ReadColumnWidthPoints(Excel.Range column) => _livePoints;

        internal override void SetColumnWidthCharacters(Excel.Range column, double width) =>
            WidthWrites.Add(width);
    }

    private static Mock<Excel.ListColumn> Column()
    {
        Mock<Excel.Range> range = new();
        Mock<Excel.ListColumn> column = new();
        column.SetupGet(c => c.Range).Returns(range.Object);
        return column;
    }

    [Fact]
    public void Every_visible_column_is_written_to_the_default_width_and_engine_columns_are_not()
    {
        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        Mock<Excel.ListColumns> columns = new();
        columns.SetupGet(c => c.Count).Returns(schema.Count);
        for (var index = 1; index <= schema.Count; index++)
        {
            Excel.ListColumn column = Column().Object;
            columns.Setup(c => c[index]).Returns(column);
        }

        TestableInitialiser initialiser = new();
        initialiser.ApplyVisibleColumnWidths(columns.Object, schema);

        int visible = schema.Count(column => !column.IsHidden);
        double expected = _liveCharacters * GanttTableAppearance.DefaultVisibleColumnWidthPt / _livePoints;

        // Two bounded corrections per column: the conversion is affine, so one pass
        // cannot land on a point target. Engine columns are skipped before any read.
        Assert.Equal(visible * 2, initialiser.WidthWrites.Count);
        Assert.All(initialiser.WidthWrites, written => Assert.Equal(expected, written, 6));
    }

    [Fact]
    public void A_column_already_on_the_default_width_is_not_written()
    {
        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        Mock<Excel.ListColumns> columns = new();
        columns.SetupGet(c => c.Count).Returns(schema.Count);
        for (var index = 1; index <= schema.Count; index++)
        {
            Excel.ListColumn column = Column().Object;
            columns.Setup(c => c[index]).Returns(column);
        }

        ConvergedInitialiser initialiser = new();
        initialiser.ApplyVisibleColumnWidths(columns.Object, schema);

        Assert.Empty(initialiser.WidthWrites);
    }

    private sealed class ConvergedInitialiser : ExcelWorkbookInitialiser
    {
        public ConvergedInitialiser()
            : base(null)
        {
        }

        public List<double> WidthWrites { get; } = [];

        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index) =>
            columns[index];

        internal override Excel.Range GetEntireColumn(Excel.Range range) => range;

        internal override double? ReadColumnWidthCharacters(Excel.Range column) => _liveCharacters;

        internal override double? ReadColumnWidthPoints(Excel.Range column) =>
            GanttTableAppearance.DefaultVisibleColumnWidthPt;

        internal override void SetColumnWidthCharacters(Excel.Range column, double width) =>
            WidthWrites.Add(width);
    }
}
