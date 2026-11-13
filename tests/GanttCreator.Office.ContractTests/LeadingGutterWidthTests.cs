using GanttCreator.Core;
using GanttCreator.Office;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R5.18k: the leading gutter column's width, and the measurement that feeds the scene
/// from the same column (ADR-0059).
/// </summary>
/// <remarks>
/// The two halves are tested together on purpose: the write makes the column a known
/// width and the read is what the scene trusts, so a bug that made them disagree would
/// otherwise sit between two green tests.
/// </remarks>
public sealed class LeadingGutterWidthTests
{
    /// <summary>The framing rows' padding height, which the gutter matches.</summary>
    private static double TargetPt => GanttCatalogues.MetricDefault("ChartPaddingRowHeightPt");

    [Fact]
    public void A_width_already_at_the_target_is_not_rewritten()
    {
        // The no-write path. A no-op pass must not mark the workbook dirty, and Excel's
        // two units never round-trip exactly, so "at the target" has to mean a band
        // rather than equality.
        Assert.Null(ExcelWorkbookInitialiser.CorrectedColumnWidthCharacters(8.43, TargetPt, TargetPt));
        Assert.Null(ExcelWorkbookInitialiser.CorrectedColumnWidthCharacters(8.43, TargetPt + 0.4, TargetPt));
    }

    [Fact]
    public void A_width_off_the_target_is_corrected_by_the_measured_ratio()
    {
        // 8.43 characters measure 48 points at the default font, so the gutter needs
        // 8.43 * 5.75 / 48 characters. Asserted by converting the answer BACK through the
        // same measured ratio: restating the formula would pass even if the units were
        // swapped.
        const double characters = 8.43;
        const double points = 48;

        double corrected = ExcelWorkbookInitialiser
            .CorrectedColumnWidthCharacters(characters, points, TargetPt)!.Value;

        Assert.Equal(TargetPt, corrected * points / characters, 6);
    }

    [Fact]
    public void An_unusable_measurement_is_not_written()
    {
        // A correction computed from an unusable pair would write a width nobody
        // measured; the column is left alone and the scene then composes without a margin.
        Assert.Null(ExcelWorkbookInitialiser.CorrectedColumnWidthCharacters(0, 48, TargetPt));
        Assert.Null(ExcelWorkbookInitialiser.CorrectedColumnWidthCharacters(8.43, 0, TargetPt));
        Assert.Null(ExcelWorkbookInitialiser.CorrectedColumnWidthCharacters(double.NaN, 48, TargetPt));
        Assert.Null(ExcelWorkbookInitialiser.CorrectedColumnWidthCharacters(8.43, double.PositiveInfinity, TargetPt));
        Assert.Null(ExcelWorkbookInitialiser.CorrectedColumnWidthCharacters(8.43, 48, 0));
    }

    [Fact]
    public void The_measurement_reads_the_gutter_column_of_the_tables_own_worksheet()
    {
        // The worksheet is reached through the table's Parent rather than by a second
        // sheet walk, so this cannot measure a different sheet's column A than the one
        // the table sits on.
        var column = new Mock<Excel.Range>();
        _ = column.SetupGet(range => range.Width).Returns(5.75d);

        var worksheet = new Mock<Excel.Worksheet>();
        _ = worksheet
            .SetupGet(sheet => sheet.Columns[GanttSheetLayout.LeadingGutterColumnIndex])
            .Returns(column.Object);

        var table = new Mock<Excel.ListObject>();
        _ = table.SetupGet(source => source.Parent).Returns(worksheet.Object);

        double? measured = new ExcelPanelGridMeasurement(null).ReadLeadingGutterWidth(table.Object);

        Assert.Equal(5.75, measured);
        worksheet.VerifyAll();
    }

    [Fact]
    public void An_unreadable_parent_measures_no_gutter_rather_than_refusing_the_chart()
    {
        // The gutter is decoration: losing it costs a margin, while refusing would cost
        // the whole chart. Every measurement that POSITIONS content still refuses.
        var table = new Mock<Excel.ListObject>();
        _ = table.SetupGet(source => source.Parent).Returns(new object());

        Assert.Null(new ExcelPanelGridMeasurement(null).ReadLeadingGutterWidth(table.Object));
    }
}
