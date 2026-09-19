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
}
