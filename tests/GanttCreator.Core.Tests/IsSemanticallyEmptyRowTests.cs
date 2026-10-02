namespace GanttCreator.Core.Tests;

/// <summary>
/// Tests for <see cref="IsSemanticallyEmptyRow"/>, the single blank-entry-row
/// rule R4.7A D3 requires.
/// </summary>
public sealed class IsSemanticallyEmptyRowTests
{
    private static GanttRowDto Row(
        int rowNumber = 2,
        string? id = null,
        string? laneId = null,
        int? stackIndex = null,
        string? typeText = null,
        string? description = null,
        DateOnly? start = null,
        DateOnly? finish = null,
        string? parentId = null,
        int? siblingOrder = null,
        string? styleKey = null,
        string? labelPositionText = null,
        string? fillColourText = null,
        string? strokeColourText = null,
        bool? visible = null,
        string? sortOrder = null
    ) =>
        new(
            rowNumber,
            Cell(id),
            Cell(laneId),
            Cell(stackIndex),
            Cell(typeText),
            Cell(description),
            Cell(start),
            Cell(finish),
            GanttCells.Empty<string>(),
            Cell(parentId),
            Cell<int>(siblingOrder),
            Cell(styleKey),
            Cell(labelPositionText),
            Cell(fillColourText),
            Cell(strokeColourText),
            Cell(visible),
            Cell(sortOrder));

    private static GanttCell<T?> Cell<T>(T? value) where T : struct =>
        value is null ? GanttCells.Empty<T?>() : GanttCells.Value(value);

    private static GanttCell<string> Cell(string? value) =>
        value is null ? GanttCells.Empty<string>() : GanttCells.Value(value);

    /// <summary>
    /// A row with nothing in it is the blank entry row. This is the invariant
    /// every other test in this file exists to protect.
    /// </summary>
    [Fact]
    public void A_row_with_no_values_is_blank()
    {
        Assert.True(IsSemanticallyEmptyRow.Test(Row()));
        Assert.True(IsSemanticallyEmptyRow.TestStrict(Row()));
        Assert.False(IsSemanticallyEmptyRow.IsEntity(Row()));
    }

    /// <summary>
    /// A null row is blank rather than a crash, so a caller filtering a partially
    /// read batch does not have to pre-filter nulls.
    /// </summary>
    [Fact]
    public void A_null_row_is_blank()
    {
        Assert.True(IsSemanticallyEmptyRow.Test(null));
        Assert.True(IsSemanticallyEmptyRow.TestStrict(null));
        Assert.False(IsSemanticallyEmptyRow.IsEntity(null));
    }

    /// <summary>
    /// Each user-authored cell independently makes the row an entity. Engine cells
    /// are excluded, and that asymmetry is deliberate -- see the separate test
    /// below that pins it.
    /// </summary>
    [Theory]
    [InlineData("As-Planned Activity", null)]
    [InlineData(null, "Build frame")]
    public void A_type_or_description_makes_the_row_an_entity(string? typeText, string? description)
    {
        GanttRowDto row = Row(typeText: typeText, description: description);

        Assert.False(IsSemanticallyEmptyRow.Test(row));
        Assert.True(IsSemanticallyEmptyRow.IsEntity(row));
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void A_date_makes_the_row_an_entity(bool withStart, bool withFinish)
    {
        GanttRowDto row = Row(
            start: withStart ? new DateOnly(2026, 9, 1) : null,
            finish: withFinish ? new DateOnly(2026, 9, 5) : null);

        Assert.False(IsSemanticallyEmptyRow.Test(row));
        Assert.True(IsSemanticallyEmptyRow.IsEntity(row));
    }

    /// <summary>
    /// Engine-maintained cells do NOT make a row an entity. This is the rule that
    /// makes "exactly one blank entry row" satisfiable: the add-in pre-fills
    /// LaneId/StackIndex on a fresh row, and treating those as content would mean
    /// no new row is ever blank.
    /// </summary>
    [Fact]
    public void Engine_maintained_values_alone_do_not_make_a_row_an_entity()
    {
        GanttRowDto row = Row(
            id: GanttRowId.New().Value,
            laneId: GanttRowId.New().Value,
            stackIndex: 0,
            parentId: GanttRowId.New().Value,
            styleKey: "AsPlannedActivity",
            labelPositionText: "Auto",
            fillColourText: "#FFFFFF",
            strokeColourText: "#000000",
            visible: true,
            sortOrder: "0");

        Assert.True(IsSemanticallyEmptyRow.Test(row));
        Assert.False(IsSemanticallyEmptyRow.IsEntity(row));
    }

    /// <summary>
    /// A whitespace-only user value counts as blank. Excel writes an empty
    /// string to a cleared cell, and a user pressing space then deleting the
    /// content should not accidentally create an entity with a description of "".
    /// </summary>
    [Fact]
    public void Whitespace_only_user_values_are_still_blank()
    {
        GanttRowDto row = Row(typeText: "   ", description: "  ");

        Assert.True(IsSemanticallyEmptyRow.Test(row));
        Assert.True(IsSemanticallyEmptyRow.TestStrict(row));
    }

    /// <summary>
    /// An Excel error in a user cell is NOT blank. `#N/A` in the Type column means
    /// the user typed something there; the row must be reported rather than
    /// silently skipped, which is the difference between Test and TestStrict.
    /// </summary>
    [Fact]
    public void An_excel_error_in_a_user_cell_is_not_blank()
    {
        GanttRowDto row = Row() with { TypeCell = GanttCells.ExcelError<string>(GanttExcelErrorCode.NotAvailable) };

        // The normalized view cannot represent the error, so the lenient rule
        // cannot see it; the strict rule reads the raw state and does.
        Assert.True(IsSemanticallyEmptyRow.Test(row));
        Assert.False(IsSemanticallyEmptyRow.TestStrict(row));
    }

    /// <summary>
    /// An unsupported payload in a user cell is also not blank, for the same
    /// reason: the user put content in a cell the reader could not type.
    /// </summary>
    [Fact]
    public void An_unsupported_payload_in_a_user_cell_is_not_blank()
    {
        GanttRowDto row = Row() with { StartCell = GanttCells.Unsupported<DateOnly?>() };

        Assert.True(IsSemanticallyEmptyRow.Test(row));
        Assert.False(IsSemanticallyEmptyRow.TestStrict(row));
    }
}
