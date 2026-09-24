namespace GanttCreator.Core.Tests;

public class GanttRowDtoTests
{
    [Fact]
    public void Dto_carries_all_neutral_fields()
    {
        var dto = new GanttRowDto(
            3,
            "G-0123456789abcdef0123456789abcdef",
            "G-abcdef0123456789abcdef0123456789",
            1,
            "As-Planned Activity",
            "Build frame",
            new DateOnly(2026, 9, 1),
            new DateOnly(2026, 9, 5),
            null,
            null,
            "Auto",
            null,
            null,
            true,
            null);

        Assert.Equal(3, dto.RowNumber);
        Assert.Equal("G-0123456789abcdef0123456789abcdef", dto.Id);
        Assert.Equal(new DateOnly(2026, 9, 1), dto.Start);
        Assert.Equal(new DateOnly(2026, 9, 5), dto.Finish);
        Assert.True(dto.Visible);
    }

    [Fact]
    public void Dto_preserves_unreadable_fields_as_null()
    {
        var dto = new GanttRowDto(
            1, null, null, null, null, null, null, null,
            null, null, null, null, null, null, null);

        Assert.Null(dto.Id);
        Assert.Null(dto.Start);
        Assert.Null(dto.StackIndex);
        Assert.Null(dto.Visible);
    }

    [Theory]
    [InlineData(2000, GanttExcelErrorCode.Null)]
    [InlineData(2007, GanttExcelErrorCode.DivisionByZero)]
    [InlineData(2015, GanttExcelErrorCode.Value)]
    [InlineData(2023, GanttExcelErrorCode.Reference)]
    [InlineData(2029, GanttExcelErrorCode.Name)]
    [InlineData(2036, GanttExcelErrorCode.Number)]
    [InlineData(2042, GanttExcelErrorCode.NotAvailable)]
    [InlineData(2043, GanttExcelErrorCode.GettingData)]
    [InlineData(2045, GanttExcelErrorCode.Spill)]
    [InlineData(2046, GanttExcelErrorCode.Connect)]
    [InlineData(2047, GanttExcelErrorCode.Blocked)]
    [InlineData(2048, GanttExcelErrorCode.ExcelUnknown)]
    [InlineData(2049, GanttExcelErrorCode.Field)]
    [InlineData(2050, GanttExcelErrorCode.Calculation)]
    public void Excel_error_mapper_preserves_known_codes(int code, GanttExcelErrorCode expected)
    {
        int value2 = -2146826288 + (code - 2000);

        Assert.True(GanttExcelErrorMapper.TryMap(value2, out GanttExcelErrorCode actual));
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(int.MaxValue)]
    [InlineData(int.MinValue)]
    public void Excel_error_mapper_preserves_future_or_unknown_integer_as_unknown(int value2)
    {
        Assert.True(GanttExcelErrorMapper.TryMap(value2, out GanttExcelErrorCode code));
        Assert.Equal(GanttExcelErrorCode.Unknown, code);
    }

    [Fact]
    public void Typed_cells_preserve_empty_value_error_and_unsupported_states()
    {
        GanttRowDto dto = new(
            1,
            GanttCells.ExcelError<string>(GanttExcelErrorCode.NotAvailable),
            GanttCells.Value("lane"),
            GanttCells.Unsupported<int?>(),
            GanttCells.Value("As-Planned Activity"),
            GanttCells.Empty<string>(),
            GanttCells.Value<DateOnly?>(new DateOnly(2026, 9, 1)),
            GanttCells.ExcelError<DateOnly?>(GanttExcelErrorCode.Value),
            GanttCells.Empty<string>(),
            GanttCells.Empty<string>(),
            GanttCells.Empty<string>(),
            GanttCells.Empty<string>(),
            GanttCells.Empty<string>(),
            GanttCells.Unsupported<bool?>(),
            GanttCells.Empty<string>());

        Assert.Equal(GanttCellState.ExcelError, dto.IdCell.State);
        Assert.Equal(GanttExcelErrorCode.NotAvailable, dto.IdCell.ErrorCode);
        Assert.Equal(GanttCellState.Value, dto.LaneIdCell.State);
        Assert.Equal(GanttCellState.Unsupported, dto.StackIndexCell.State);
        Assert.Equal(GanttCellState.Empty, dto.DescriptionCell.State);
        Assert.Equal(GanttCellState.ExcelError, dto.FinishCell.State);
        Assert.Equal(GanttExcelErrorCode.Value, dto.FinishCell.ErrorCode);
        Assert.Equal(GanttCellState.Unsupported, dto.VisibleCell.State);
        Assert.Null(dto.Id);
        Assert.Equal("lane", dto.LaneId);
    }
}
