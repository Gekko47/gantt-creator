namespace GanttCreator.Core.Tests;

/// <summary>
/// Contract tests for <see cref="GanttSheetLayout"/>, the single authority for the
/// live sheet's vertical layout and plot anchor (R4.7I slice 1 D1).
/// </summary>
public sealed class GanttSheetLayoutTests
{
    /// <summary>
    /// One row is reserved above the table, and the header sits directly beneath it.
    /// </summary>
    /// <remarks>
    /// The owner's original request was two reserved rows — one for the year band and
    /// one for the period band. ADR-0030 D4/D5 implements <b>one</b>, because entity
    /// guide §4 already requires the header's height to align with the period header's
    /// bottom, so the header row <em>is</em> the period band's row. This test pins
    /// the one-row decision so a later "add the second band row" change has to amend
    /// ADR-0030 D5 and the guide rather than drift in quietly.
    /// </remarks>
    [Fact]
    public void Exactly_one_row_is_reserved_above_the_header()
    {
        Assert.Equal(1, GanttSheetLayout.ReservedRowCount);
        Assert.Equal(1, GanttSheetLayout.ReservedRowIndex);
        Assert.Equal(2, GanttSheetLayout.HeaderRowIndex);
    }

    /// <summary>
    /// The first body row follows the header, and the plot's top is that row's top.
    /// </summary>
    /// <remarks>
    /// The whole point of ADR-0030. The reported defect was a lane drawn ~43pt
    /// (2.4 rows) below its own row because the plot's top came from the size
    /// preset's page coordinates instead of the worksheet.
    /// </remarks>
    [Fact]
    public void The_first_body_row_is_the_row_after_the_header()
    {
        Assert.Equal(
            GanttSheetLayout.HeaderRowIndex + 1,
            GanttSheetLayout.FirstBodyRowIndex);
    }

    /// <summary>
    /// The row indices are contiguous with no gap, so no band can be stranded between
    /// the header and the plot.
    /// </summary>
    [Fact]
    public void The_reserved_header_and_body_rows_are_contiguous()
    {
        Assert.Equal(
            GanttSheetLayout.ReservedRowIndex + 1,
            GanttSheetLayout.HeaderRowIndex);
        Assert.Equal(
            GanttSheetLayout.HeaderRowIndex + 1,
            GanttSheetLayout.FirstBodyRowIndex);
    }

    /// <summary>
    /// The anchor is the column one past the table's last column — the same rule the
    /// pre-ADR-0030 code used, unchanged.
    /// </summary>
    [Fact]
    public void The_anchor_column_is_one_past_the_last_schema_column()
    {
        Assert.Equal(
            GanttTableSchema.Default.Columns.Count + 1,
            GanttSheetLayout.PlotAnchorColumnIndex);
    }

    /// <summary>
    /// The anchor names the <b>header</b> row, not row 1.
    /// </summary>
    /// <remarks>
    /// This is the assertion the three duplicated <c>$1</c> literals could not make.
    /// A reserved row moves the header to row 2, and the anchor is contractually "the
    /// cell one column right of the table's last column, <em>on the header row</em>"
    /// (<c>02-ARCHITECTURE.md</c>). Writing the row from the layout authority rather
    /// than from a literal is what keeps the initialiser and the integrity checker in
    /// agreement.
    /// </remarks>
    [Fact]
    public void The_anchor_sits_on_the_header_row()
    {
        string refersTo = GanttSheetLayout.BuildPlotAnchorRefersTo("Gantt Data");

        Assert.Equal(
            $"='Gantt Data'!${GanttSheetLayout.ToA1Column(GanttSheetLayout.PlotAnchorColumnIndex)}${GanttSheetLayout.HeaderRowIndex}",
            refersTo);
        Assert.DoesNotContain("$1'", refersTo, StringComparison.Ordinal);
    }

    /// <summary>
    /// A sheet name containing an apostrophe is escaped, because Excel requires each
    /// apostrophe inside a quoted sheet reference to be doubled.
    /// </summary>
    [Fact]
    public void A_sheet_name_containing_an_apostrophe_is_escaped()
    {
        // The anchor column is derived from the schema rather than restated, so this
        // test states the escaping rule and not the schema's current width. Hardcoding
        // a column letter here is exactly the drift the layout authority removes: the
        // real regression this repo records is a literal going stale when a schema ADR
        // added a column.
        Assert.Equal(
            $"='Bob''s Schedule'!${GanttSheetLayout.ToA1Column(GanttSheetLayout.PlotAnchorColumnIndex)}${GanttSheetLayout.HeaderRowIndex}",
            GanttSheetLayout.BuildPlotAnchorRefersTo("Bob's Schedule"));
    }

    /// <summary>
    /// A column index past <c>Z</c> produces two letters rather than a punctuation
    /// character.
    /// </summary>
    /// <remarks>
    /// <b>Positive test for the reason <see cref="GanttSheetLayout.ToA1Column"/> is
    /// not a single cast.</b> The naive <c>(char)('A' + index - 1)</c> form is
    /// correct to column 26 and emits <c>[</c> at 27, producing a reference that
    /// names no cell. The schema has 14 columns today, so without this test the bug
    /// would stay dormant until a schema ADR added enough columns to reach
    /// <c>AA</c>.
    /// </remarks>
    [Theory]
    [InlineData(1, "A")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(28, "AB")]
    [InlineData(52, "AZ")]
    [InlineData(53, "BA")]
    [InlineData(702, "ZZ")]
    [InlineData(703, "AAA")]
    public void Column_indices_convert_to_the_expected_A1_letters(int columnIndex, string expected)
    {
        Assert.Equal(expected, GanttSheetLayout.ToA1Column(columnIndex));
    }

    /// <summary>
    /// A column index of zero or less is refused rather than converted.
    /// </summary>
    /// <remarks>
    /// <b>Positive test for the validator.</b> A zero column would otherwise emit an
    /// empty string and produce a reference like <c>='Sheet'!$2</c>, which is not a
    /// cell address at all.
    /// </remarks>
    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void A_non_positive_column_index_is_refused(int columnIndex)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => GanttSheetLayout.ToA1Column(columnIndex));
    }

    /// <summary>
    /// A null sheet name is refused as a null argument.
    /// </summary>
    /// <remarks>
    /// <b>Positive test for the validator.</b> <c>ThrowIfNullOrWhiteSpace</c> reports
    /// null as <see cref="ArgumentNullException"/>, so it is asserted separately from
    /// the blank case rather than as one family: xUnit's <c>Assert.Throws</c> requires
    /// an exact type match, and collapsing the two would have hidden which contract
    /// actually holds.
    /// </remarks>
    [Fact]
    public void A_null_sheet_name_is_refused_as_a_null_argument()
    {
        Assert.Throws<ArgumentNullException>(
            () => GanttSheetLayout.BuildPlotAnchorRefersTo(null!));
    }

    /// <summary>
    /// A blank or whitespace-only sheet name is refused rather than written as an
    /// anchor that can never resolve.
    /// </summary>
    /// <remarks>
    /// <b>Positive test for the validator.</b> The three previous call sites each
    /// formatted the string unconditionally, so a blank name produced
    /// <c>=''!$O$2</c> — a syntactically valid string that names nothing, which is
    /// the failure this guard exists to prevent.
    /// </remarks>
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("\t")]
    public void A_blank_sheet_name_is_refused(string sheetName)
    {
        Assert.Throws<ArgumentException>(() => GanttSheetLayout.BuildPlotAnchorRefersTo(sheetName));
    }

    /// <summary>
    /// The anchor is a complete, absolute A1 reference.
    /// </summary>
    [Fact]
    public void The_anchor_is_an_absolute_quoted_reference()
    {
        string refersTo = GanttSheetLayout.BuildPlotAnchorRefersTo("Gantt Data");

        Assert.StartsWith("='Gantt Data'!$", refersTo, StringComparison.Ordinal);
        Assert.EndsWith($"${GanttSheetLayout.HeaderRowIndex}", refersTo, StringComparison.Ordinal);
    }

    /// <summary>
    /// The layout authority is Office-free, so it stays testable without a live host.
    /// </summary>
    [Fact]
    public void The_layout_authority_has_no_office_dependency()
    {
        string source = File.ReadAllText(LocateSourceFile());

        Assert.DoesNotContain("Microsoft.Office", source, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Interop.Excel", source, StringComparison.OrdinalIgnoreCase);
    }

    private static string LocateSourceFile()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            string candidate = Path.Combine(
                directory.FullName,
                "src",
                "GanttCreator.Core",
                "GanttSheetLayout.cs");
            if (File.Exists(candidate))
            {
                return candidate;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException(
            "Could not locate src/GanttCreator.Core/GanttSheetLayout.cs from " + AppContext.BaseDirectory);
    }
}