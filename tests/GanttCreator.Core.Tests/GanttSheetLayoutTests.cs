namespace GanttCreator.Core.Tests;

/// <summary>
/// Contract tests for <see cref="GanttSheetLayout"/>, the single authority for the
/// live sheet's vertical layout and plot anchor (R4.7I slice 1 D1).
/// </summary>
public sealed class GanttSheetLayoutTests
{
    /// <summary>
    /// <summary>
    /// Two rows are reserved above the table: a top padding row, then the title/year
    /// row, with the header directly beneath.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The owner's original request was two reserved rows — one for the year band and
    /// one for the period band. ADR-0030 D4/D5 implements <b>one band row</b>, because
    /// entity guide §4 already requires the header's height to align with the period
    /// header's bottom, so the header row <em>is</em> the period band's row. That
    /// decision still stands and is why the header is directly beneath the title row.
    /// </para>
    /// <para>
    /// ADR-0031 D1 adds a <b>second</b> row above the title row, for a different
    /// reason: it is the chart's top margin, not a band. The owner asked for the
    /// chart's top and bottom padding to be real worksheet rows so its edges line up
    /// with the sheet's grid, rather than a sub-row sliver of chart chrome. This test
    /// pins the two-row arrangement so a later change to either the band decision or
    /// the padding rows has to amend the ADR and the guide rather than drift quietly.
    /// </para>
    /// </remarks>
    [Fact]
    public void Exactly_two_rows_are_reserved_above_the_header()
    {
        Assert.Equal(2, GanttSheetLayout.ReservedRowCount);
        Assert.Equal(1, GanttSheetLayout.TopPaddingRowIndex);
        Assert.Equal(2, GanttSheetLayout.ReservedRowIndex);
        Assert.Equal(3, GanttSheetLayout.HeaderRowIndex);
    }

    /// <summary>
    /// The reserved bottom padding row is the row immediately below the measured
    /// body, and nothing else (ADR-0035 D2).
    /// </summary>
    /// <remarks>
    /// This is the assertion whose absence caused the live defect. The two adapters
    /// used to take <see cref="GanttSheetLayout.BottomPaddingRowIndex"/> on trust --
    /// one writing a row height to it, one reading it as the chart's bottom margin
    /// -- and nothing checked that the row existed, was the right one, or was empty.
    /// </remarks>
    [Fact]
    public void The_reserved_bottom_padding_row_is_the_row_directly_below_the_measured_body()
    {
        // Body rows 4..10, so the padding row is 11. This agrees with the derived
        // index for a seven-row body, which is the case where the old arithmetic
        // happened to be right.
        BottomPaddingResolution resolved = GanttSheetLayout.ResolveBottomPaddingRow(
            firstBodyRow: 4,
            lastBodyRow: 10,
            observedPaddingRow: 11,
            paddingRowIsEmpty: true);

        Assert.True(resolved.Succeeded);
        Assert.Equal(11, resolved.Row);
        Assert.Null(resolved.Refusal);
    }

    /// <summary>
    /// The measured body wins over the derived one: a body that has MOVED resolves
    /// against where it actually is.
    /// </summary>
    /// <remarks>
    /// This is the case the derived index could never handle. If the table sits at
    /// rows 20..25, the old arithmetic named row 11 -- nine rows above the table,
    /// inside the header band area -- and wrote the chart's 6pt margin height there.
    /// Resolving from the measured span is what makes a moved table correct.
    /// </remarks>
    [Fact]
    public void A_moved_table_resolves_against_its_measured_span_not_the_derived_index()
    {
        BottomPaddingResolution resolved = GanttSheetLayout.ResolveBottomPaddingRow(
            firstBodyRow: 20,
            lastBodyRow: 25,
            observedPaddingRow: 26,
            paddingRowIsEmpty: true);

        Assert.True(resolved.Succeeded);
        Assert.Equal(26, resolved.Row);

        // The derived index would have said 11 for a six-row body -- provably not
        // the same answer, so this test discriminates rather than restating.
        Assert.NotEqual(GanttSheetLayout.BottomPaddingRowIndex(6), resolved.Row);
    }

    /// <summary>
    /// A padding row that is not contiguous with the body is refused: the add-in
    /// would be writing to a row it does not own.
    /// </summary>
    /// <remarks>
    /// <b>Positive test for the validator.</b> The gap case is what a user gets by
    /// deleting a worksheet row between the table and the margin, and the
    /// inside-the-body case is what the old arithmetic produced. Both must be
    /// reported, because both are workbook states Repair can act on.
    /// </remarks>
    [Theory]
    [InlineData(12)] // a gap: two rows below the body
    [InlineData(10)] // inside the body: the last body row itself
    [InlineData(4)] // the first body row
    public void A_padding_row_that_is_not_the_row_below_the_body_is_refused(int observed)
    {
        BottomPaddingResolution resolved = GanttSheetLayout.ResolveBottomPaddingRow(
            firstBodyRow: 4,
            lastBodyRow: 10,
            observedPaddingRow: observed,
            paddingRowIsEmpty: true);

        Assert.False(resolved.Succeeded);
        Assert.Equal(BottomPaddingRefusalReason.PaddingRowNotContiguous, resolved.Refusal);
        Assert.Null(resolved.Row);
    }

    /// <summary>
    /// An occupied or unverified padding row is refused, because it is the user's
    /// row rather than the chart's margin.
    /// </summary>
    /// <remarks>
    /// <b>Positive test for the validator</b>, and the one that closes the reported
    /// defect. <c>null</c> is refused as well as <c>false</c>: an unestablished
    /// emptiness must not be read as an empty row, because that substitution is
    /// exactly how the add-in came to resize a row the user had typed in.
    /// </remarks>
    [Theory]
    [InlineData(false)]
    [InlineData(null)]
    public void An_occupied_or_unverified_padding_row_is_refused(bool? isEmpty)
    {
        BottomPaddingResolution resolved = GanttSheetLayout.ResolveBottomPaddingRow(
            firstBodyRow: 4,
            lastBodyRow: 10,
            observedPaddingRow: 11,
            paddingRowIsEmpty: isEmpty);

        Assert.False(resolved.Succeeded);
        Assert.Equal(BottomPaddingRefusalReason.PaddingRowOccupied, resolved.Refusal);
        Assert.Null(resolved.Row);
    }

    /// <summary>
    /// A body span that is not a body is refused before any row is named.
    /// </summary>
    /// <remarks>
    /// <b>Positive test for the validator.</b> Grouped as one theory because both
    /// inputs describe the same failure -- there is no body to sit below -- and
    /// neither can produce a plausible row index. Asserting them separately would
    /// only restate the same two guards.
    /// </remarks>
    [Theory]
    [InlineData(0, 10)]
    [InlineData(-1, 10)]
    [InlineData(4, 0)]
    [InlineData(10, 4)]
    public void A_span_that_is_not_a_body_is_refused(int firstBodyRow, int lastBodyRow)
    {
        BottomPaddingResolution resolved = GanttSheetLayout.ResolveBottomPaddingRow(
            firstBodyRow,
            lastBodyRow,
            observedPaddingRow: 11,
            paddingRowIsEmpty: true);

        Assert.False(resolved.Succeeded);
        Assert.NotNull(resolved.Refusal);
        Assert.Null(resolved.Row);
    }

    /// <summary>
    /// A padding row that could not be identified at all is refused.
    /// </summary>
    /// <remarks>
    /// <b>Positive test for the validator.</b> The host failing to report a row is
    /// absence, and absence is not contiguity.
    /// </remarks>
    [Fact]
    public void An_unidentified_padding_row_is_refused()
    {
        BottomPaddingResolution resolved = GanttSheetLayout.ResolveBottomPaddingRow(
            firstBodyRow: 4,
            lastBodyRow: 10,
            observedPaddingRow: null,
            paddingRowIsEmpty: true);

        Assert.False(resolved.Succeeded);
        Assert.Equal(BottomPaddingRefusalReason.PaddingRowNotContiguous, resolved.Refusal);
    }

    /// <summary>
    /// The title occupies exactly ONE cell, directly above the <c>Description</c>
    /// column, on the title row.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This asserted D2:H2 - a five-column span - on the reasoning that Excel resolves
    /// a range assignment to the range's top-left cell. <b>That premise is false</b>:
    /// assigning <c>Value2</c> to a multi-cell range writes into every cell, so the
    /// live sheet rendered the title once per visible column. The owner asked for a
    /// single title cell above <c>Description</c>.
    /// </para>
    /// <para>
    /// The column is derived from the SCHEMA's <c>Description</c> column and the A1
    /// reference is derived from that, so a schema that adds or reorders columns moves
    /// this test rather than stranding the title over a hidden column.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_title_is_one_cell_directly_above_the_description_column()
    {
        // ONE column, so start and end agree and the span is 1. A span above 1 is
        // exactly what made Excel write the value into every cell.
        Assert.Equal(1, GanttSheetLayout.TitleColumnSpan);
        Assert.Equal(GanttSheetLayout.TitleColumnIndex, GanttSheetLayout.TitleStartColumnIndex);
        Assert.Equal(GanttSheetLayout.TitleColumnIndex, GanttSheetLayout.TitleEndColumnIndex);

        // The derived reference is the single cell above Description.
        Assert.Equal(
            "E2",
            GanttSheetLayout.ToA1Cell(GanttSheetLayout.TitleColumnIndex, GanttSheetLayout.TitleRowIndex));

        // It starts AFTER the engine columns: a title in column A would sit over
        // hidden bookkeeping the user cannot see.
        Assert.True(GanttSheetLayout.TitleColumnIndex > 1);

        // And it really is the Description column's own index, not a coincidence.
        Assert.Equal(
            GanttTableSchema.Default.Columns
                .Select((column, index) => (column, index))
                .Single(entry => string.Equals(entry.column.Name, "Description", StringComparison.Ordinal))
                .index + 1,
            GanttSheetLayout.TitleColumnIndex);
    }

    /// <summary>
    /// The bottom padding row is the row directly below the last activity row, and
    /// it is derived from the body length rather than fixed.
    /// </summary>
    /// <remarks>
    /// A fixed index would be correct for exactly one table length. This also proves
    /// the derived value is not simply the header row: a body that starts and ends
    /// where it should is what makes the chart's bottom margin land below the chart
    /// rather than inside the table.
    /// </remarks>
    [Fact]
    public void The_bottom_padding_row_is_the_row_after_the_last_activity_row()
    {
        // Body starts at row 4, so seven activity rows occupy 4..10 and the bottom
        // padding row is 11: the row directly below the last activity row.
        Assert.Equal(11, GanttSheetLayout.BottomPaddingRowIndex(7));
        Assert.Equal(5, GanttSheetLayout.BottomPaddingRowIndex(1));

        // A table with no body has no last activity row, so there is no row to name.
        // Returning a plausible index would place the chart's bottom margin inside
        // the table, so this is refused rather than guessed.
        Assert.Throws<ArgumentOutOfRangeException>(() => GanttSheetLayout.BottomPaddingRowIndex(0));
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
            GanttSheetLayout.TopPaddingRowIndex + 1,
            GanttSheetLayout.ReservedRowIndex);
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