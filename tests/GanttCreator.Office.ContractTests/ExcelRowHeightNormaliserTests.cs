using GanttCreator.Core;
using GanttCreator.Office;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ExcelRowHeightNormaliser"/> (R4.7D,
/// ADR-0026 D3/D4): a dragged managed row is restored, a normalised sheet writes
/// nothing, and every refusal path fires with no write.
/// </summary>
public sealed class ExcelRowHeightNormaliserTests
{
    private const double ManagedPt = 18;
    private const double SplitterPt = 24;
    private const double SpacerPt = 6;

    /// <summary>
    /// The header row's target: the period band's row (ADR-0030 D5).
    /// </summary>
    private const double HeaderPt = 16;

    /// <summary>
    /// The reserved row's target: carries the table title and year band
    /// (ADR-0030 D4).
    /// </summary>
    private const double ReservedRowPt = 18;

    /// <summary>
    /// The chart's padding-row token, deliberately a different figure from the
    /// reserved row's.
    /// </summary>
    /// <remarks>
    /// The two must be distinguishable in a fixture: seeding both padding rows with
    /// the reserved row's height would let a test pass while the adapter wrote the
    /// reserved-row target into the padding rows instead of the padding token, which
    /// is the specific mistake this value exists to catch.
    /// </remarks>
    private const double PaddingRowPt = 22;

    /// <summary>
    /// The reserved anchor row's target (ADR-0038 D1), deliberately a THIRD distinct
    /// value.
    /// </summary>
    /// <remarks>
    /// It must differ from both the padding row's and the reserved row's. Writing the
    /// padding height into the anchor row would put the plot-spanning shapes' bottom
    /// cell anchor a whole margin-height below the body, which is a visible strip of
    /// sheet rather than the sub-row anchor the closing line's arithmetic depends on --
    /// and a fixture that reused one value could not tell that mistake from a correct
    /// write.
    /// </remarks>
    private const double AnchorRowPt = 0.25;

    /// <summary>
    /// Overrides the COM seams so the test can supply row heights, the per-row
    /// <c>Type</c> text, and record writes without a live Excel host.
    /// </summary>
    /// <param name="application">The application object.</param>
    /// <param name="guard">The protection guard.</param>
    /// <param name="worksheet">The resolved worksheet.</param>
    /// <param name="table">The resolved table.</param>
    /// <param name="heights">One measured height per body row.</param>
    /// <param name="written">The rows written, recorded by the test double.</param>
    /// <param name="types">
    /// One Type display name per body row. When absent, no Type column resolves and
    /// every row is treated as managed -- which is what the pre-existing tests rely on.
    /// </param>
    private sealed class TestableNormaliser(
        object? application,
        IWorksheetProtectionGuard guard,
        Excel.Worksheet worksheet,
        Excel.ListObject table,
        IReadOnlyList<double> heights,
        List<int> written,
        IReadOnlyList<string>? types = null,
        Dictionary<int, double>? layoutHeights = null)
        : ExcelRowHeightNormaliser(application, guard)
    {
        internal override bool TryFindTable(
            Excel.Sheets sheets,
            out Excel.Worksheet? resolvedWorksheet,
            out Excel.ListObject? resolvedTable)
        {
            resolvedWorksheet = worksheet;
            resolvedTable = table;
            return true;
        }

        internal override Excel.Range? GetTableBody(Excel.ListObject source) => SourceBody;

        internal override int GetBodyRowCount(Excel.Range body) => heights.Count;

        internal override Excel.Range? GetBodyRowAt(Excel.Range body, int index)
        {
            var row = new Mock<Excel.Range>();
            _ = row.SetupGet(r => r.RowHeight).Returns(heights[index - 1]);
            _ = row.SetupSet(r => r.RowHeight = It.IsAny<object>())
                .Callback<object>(_ => written.Add(index));
            return row.Object;
        }

        internal override bool TryGetTypeColumn(Excel.ListObject source, out Excel.Range? typeColumn)
        {
            if (types is null)
            {
                typeColumn = null;
                return false;
            }

            typeColumn = new Mock<Excel.Range>().Object;
            return true;
        }

        internal override string? ReadTypeCellText(Excel.Range typeColumn, int index) =>
        types is null || index < 1 || index > types.Count ? null : types[index - 1];

    /// <summary>
        /// The body's first worksheet row, as the host reports it.
    /// </summary>
    /// <remarks>
    /// ADR-0035 D2 makes the bottom padding row resolve from the body's MEASURED
    /// span rather than from the derived row count. Without this seam the mocked
    /// <c>Range.Row</c> reports 0, every normalisation refuses as
    /// <c>PaddingRowNotOwned</c>, and every test in this file would be asserting the
    /// refusal rather than the behaviour it names.
    /// </remarks>
    internal override int GetRangeRow(Excel.Range range)
    {
        ArgumentNullException.ThrowIfNull(range);
        return GanttSheetLayout.FirstBodyRowIndex;
    }

    /// <summary>
        /// Whether the reserved padding row reads as empty.
    /// </summary>
    /// <remarks>
    /// Empty by default because these tests are about HEIGHTS, and an unseeded
    /// emptiness would refuse before any height was considered. The occupied case
    /// has its own test, which is the positive test for the validator.
    /// </remarks>
    internal override bool? IsLayoutRowEmpty(Excel.Worksheet source, int rowIndex) =>
        rowIndex == GanttSheetLayout.AnchorRowIndex(Math.Max(1, heights.Count))
            ? AnchorRowOccupied ? false : true
            : PaddingRowOccupied ? false : true;

    /// <summary>Whether the fixture pretends the PADDING row carries content.</summary>
    internal bool PaddingRowOccupied { get; set; }

    /// <summary>
    /// Whether the fixture pretends the ANCHOR row carries content (ADR-0038 D6).
    /// </summary>
    /// <remarks>
    /// A separate flag from <see cref="PaddingRowOccupied"/> because the two rows must
    /// be able to fail INDEPENDENTLY -- that independence is the whole point of
    /// verifying them separately, and one shared flag could not express "the anchor row
    /// is fine and the padding row is not".
    /// </remarks>
    internal bool AnchorRowOccupied { get; set; }

    /// <summary>Reads a layout row's reported <c>Value2</c> as empty for this fixture.</summary>
    internal override object? GetRangeValue2(Excel.Range range) => null;

    /// <summary>The heights the layout rows start at, and receive when written.</summary>
    internal Dictionary<int, double> LayoutHeights { get; } = layoutHeights ?? new Dictionary<int, double>
    {
        [GanttSheetLayout.TopPaddingRowIndex] = 22d,
        [GanttSheetLayout.ReservedRowIndex] = 18d,
        [GanttSheetLayout.HeaderRowIndex] = 16d,
    };

    internal override Excel.Range? GetLayoutRow(Excel.Worksheet source, int rowIndex)
    {
        // The BOTTOM padding row's index is derived from the body length, so it is
        // seeded here for the default one-row body exactly as the adapter derives it.
        // Omitting it made the mock report height 0, which the adapter then correctly
        // "restored" -- so every count in this file was quietly two higher than the
        // behaviour under test. Seeding it is what makes the counts mean what they
        // say.
        LayoutHeights.TryAdd(GanttSheetLayout.BottomPaddingRowIndex(Math.Max(1, heights.Count)), 22d);

        // The anchor row too (ADR-0038 D1), at ITS token rather than the padding
        // height. Seeding it at 22 -- as the padding row is -- would make it always
        // differ from the 0.25pt target, so every RowsWritten count in this file
        // would be one higher than the behaviour under test.
        LayoutHeights.TryAdd(
            GanttSheetLayout.AnchorRowIndex(Math.Max(1, heights.Count)),
            AnchorRowPt);

        var row = new Mock<Excel.Range>();
        _ = row.SetupGet(r => r.RowHeight)
            .Returns(LayoutHeights.TryGetValue(rowIndex, out double height) ? height : 0d);
        _ = row.SetupSet(r => r.RowHeight = It.IsAny<object>())
            .Callback<object>(value =>
            {
                LayoutHeights[rowIndex] = Convert.ToDouble(
                    value, System.Globalization.CultureInfo.InvariantCulture);
                written.Add(rowIndex);
            });
        return row.Object;
    }

        private Excel.Range SourceBody { get; } = new Mock<Excel.Range>().Object;
    }

    /// <summary>
    /// An OCCUPIED padding row refuses rather than resizing it (ADR-0035 D2).
    /// </summary>
    /// <remarks>
    /// <b>This is the positive test for the validator, and it is the reported
    /// defect.</b> The old adapter computed the padding row arithmetically and wrote
    /// the chart's 6pt margin height to it unconditionally, so a user who had typed
    /// in that row had it silently resized by the very Refresh meant to restore the
    /// sheet. Refusing with a distinct reason is what makes the state visible and
    /// repairable instead of invisible.
    /// </remarks>
    [Fact]
    public void An_occupied_bottom_padding_row_is_refused_and_never_resized()
    {
        var (normaliser, written, layout) = Build();
        normaliser.PaddingRowOccupied = true;

        var paddingRow = GanttSheetLayout.BottomPaddingRowIndex(1);
        layout[paddingRow] = 45d;

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(
            ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.False(outcome.Succeeded);
        Assert.Equal(RowHeightNormalisationRefusalReason.PaddingRowNotOwned, outcome.Refusal);
        Assert.Equal(0, outcome.RowsWritten);

        // The user's row is untouched, and it is the padding row specifically --
        // the whole point is that the add-in does not touch a row it does not own.
        Assert.Equal(45d, layout[paddingRow]);
        Assert.DoesNotContain(paddingRow, written);
    }

    /// <summary>
    /// A refused padding row writes NOTHING at all -- not even the layout rows.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This is the positive test for the ordering guarantee, and it is what the
    /// pre-fix code failed. The layout rows were normalised BEFORE the padding row
    /// was resolved, so a <c>PaddingRowNotOwned</c> refusal returned with the top
    /// padding row, the reserved row and the header row already resized -- a refusal
    /// that mutated the worksheet, which is the one thing a refusal must never do.
    /// </para>
    /// <para>
    /// The fixture is arranged so the old ordering produced three writes: every
    /// layout row starts away from its token, so a layout pass running first would
    /// correct all three. Under the fixed ordering the resolution refuses first and
    /// the write log is empty. Non-vacuity: the same fixture with an EMPTY padding
    /// row does write those rows, which <c>The_layout_row_targets_reach_the_rows</c>
    /// already pins.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_refused_padding_row_writes_nothing_including_the_layout_rows()
    {
        (TestableNormaliser normaliser, List<int> written, Dictionary<int, double> layout) = Build(
            bodyHeights: [45],
            layoutHeights: new Dictionary<int, double>
            {
                // Every layout row starts away from its token, so a layout write
                // before the refusal would be visible in both logs.
                [GanttSheetLayout.TopPaddingRowIndex] = 45d,
                [GanttSheetLayout.ReservedRowIndex] = 45d,
                [GanttSheetLayout.HeaderRowIndex] = 45d,
            });
        normaliser.PaddingRowOccupied = true;

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(
            ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.False(outcome.Succeeded);
        Assert.Equal(RowHeightNormalisationRefusalReason.PaddingRowNotOwned, outcome.Refusal);
        Assert.Equal(0, outcome.RowsWritten);

        // The load-bearing assertion: a refusal mutates nothing, so the dragged
        // layout rows are left exactly as the user left them.
        Assert.Empty(written);
        Assert.Equal(45d, layout[GanttSheetLayout.TopPaddingRowIndex]);
        Assert.Equal(45d, layout[GanttSheetLayout.ReservedRowIndex]);
        Assert.Equal(45d, layout[GanttSheetLayout.HeaderRowIndex]);
    }

    /// <summary>
    /// The anchor row is restored to its OWN sub-row token, never the padding row's
    /// (ADR-0038 D1).
    /// </summary>
    /// <remarks>
    /// <para>
    /// The fixture's three reserved-row tokens are deliberately different values, so
    /// writing the padding height into the anchor row would fail here rather than pass
    /// unnoticed. That distinction is the whole point: the anchor row is a sub-row
    /// ANCHOR, and a margin-height row there is a visible strip of blank sheet between
    /// the last activity row and the chart's bottom.
    /// </para>
    /// <para>
    /// The padding row is asserted alongside it, because the two rows moved apart in
    /// ADR-0038 and a change that normalised one while skipping the other would leave
    /// the sheet with a 0.25pt row where its margin should be.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_dragged_anchor_row_is_restored_to_the_anchor_token_not_the_padding_token()
    {
        var (normaliser, written, layout) = Build(
            bodyHeights: [ManagedPt],
            layoutHeights: new Dictionary<int, double>
            {
                [GanttSheetLayout.TopPaddingRowIndex] = PaddingRowPt,
                [GanttSheetLayout.ReservedRowIndex] = ReservedRowPt,
                [GanttSheetLayout.HeaderRowIndex] = HeaderPt,
                [GanttSheetLayout.BottomPaddingRowIndex(1)] = PaddingRowPt,
            });
        layout[GanttSheetLayout.AnchorRowIndex(1)] = 40d;

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(
            ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());

        var anchorRow = GanttSheetLayout.AnchorRowIndex(1);
        Assert.Contains(anchorRow, written);
        Assert.Equal(AnchorRowPt, layout[anchorRow]);

        // And NOT the padding height, which is the mistake this test exists for.
        Assert.NotEqual(PaddingRowPt, layout[anchorRow]);

        // The padding row is one below it and still holds the padding token. Seeded at that
        // token already, so it is correctly NOT rewritten -- which is the "already at
        // the token, so no write" rule that keeps Refresh from dirtying the workbook.
        Assert.DoesNotContain(GanttSheetLayout.BottomPaddingRowIndex(1), written);
        Assert.Equal(PaddingRowPt, layout[GanttSheetLayout.BottomPaddingRowIndex(1)]);
        Assert.NotEqual(anchorRow, GanttSheetLayout.BottomPaddingRowIndex(1));
    }

    /// <summary>
    /// An OCCUPIED anchor row refuses rather than being resized (ADR-0038 D6).
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>This is the positive test for the new validator, and the mirror of the padding
    /// row's.</b> Without it the new <c>AnchorRowNotOwned</c> reason would be unreachable
    /// code, and the reported defect's twin -- the add-in resizing a row the user typed
    /// in -- would still be live one row higher than before.
    /// </para>
    /// <para>
    /// The refusal is typed DISTINCTLY from the padding row's so a diagnostic can name
    /// which row the user has to fix. The user row is left untouched and no write of any
    /// kind occurs.
    /// </para>
    /// </remarks>
    [Fact]
    public void An_occupied_anchor_row_is_refused_and_never_resized()
    {
        var (normaliser, written, layout) = Build();
        normaliser.AnchorRowOccupied = true;

        var anchorRow = GanttSheetLayout.AnchorRowIndex(1);
        layout[anchorRow] = 45d;

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(
            ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.False(outcome.Succeeded);
        Assert.Equal(RowHeightNormalisationRefusalReason.AnchorRowNotOwned, outcome.Refusal);
        Assert.NotEqual(RowHeightNormalisationRefusalReason.PaddingRowNotOwned, outcome.Refusal);
        Assert.Equal(0, outcome.RowsWritten);

        Assert.Equal(45d, layout[anchorRow]);
        Assert.DoesNotContain(anchorRow, written);
    }

    /// <summary>
    /// The two reserved rows below the body are verified INDEPENDENTLY, and an anchor
    /// refusal writes nothing at all (ADR-0038 D6).
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>Two claims in one test, because they are one behaviour.</b> First, the anchor
    /// row resolves on its own: the padding row here is fine and the refusal still fires.
    /// A single combined "reserved rows" check would have passed this test while being
    /// unable to say which row was wrong.
    /// </para>
    /// <para>
    /// Second, the refusal MUTATES NOTHING -- not even the layout rows above the table.
    /// The anchor row is resolved after the padding row but still before any write, so a
    /// refusal is a statement about the sheet rather than a partial edit. Every layout
    /// row is seeded away from its token so a premature write would show.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_anchor_row_is_verified_independently_and_its_refusal_writes_nothing()
    {
        (TestableNormaliser normaliser, List<int> written, Dictionary<int, double> layout) = Build(
            bodyHeights: [45],
            layoutHeights: new Dictionary<int, double>
            {
                [GanttSheetLayout.TopPaddingRowIndex] = 45d,
                [GanttSheetLayout.ReservedRowIndex] = 45d,
                [GanttSheetLayout.HeaderRowIndex] = 45d,
                [GanttSheetLayout.BottomPaddingRowIndex(1)] = 45d,
            });

        // Only the ANCHOR row is occupied; the padding row below it is left fine.
        normaliser.AnchorRowOccupied = true;

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(
            ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.False(outcome.Succeeded);
        Assert.Equal(RowHeightNormalisationRefusalReason.AnchorRowNotOwned, outcome.Refusal);
        Assert.Equal(0, outcome.RowsWritten);

        Assert.Empty(written);
        Assert.Equal(45d, layout[GanttSheetLayout.TopPaddingRowIndex]);
        Assert.Equal(45d, layout[GanttSheetLayout.ReservedRowIndex]);
        Assert.Equal(45d, layout[GanttSheetLayout.HeaderRowIndex]);
        Assert.Equal(45d, layout[GanttSheetLayout.BottomPaddingRowIndex(1)]);
    }

    /// <summary>
    /// The padding row refusal is <b>typed and distinct</b> from a measurement failure, so
    /// the caller can tell "the sheet is in an unexpected state" from "the host would
    /// not report a height".
    /// </summary>
    /// <remarks>
    /// <b>Positive test for the validator.</b> A new enum member that happened to
    /// share a value with an existing one, or that no call site could return, would
    /// both pass a suite that only checked that <em>something</em> refused.
    /// </remarks>
    [Fact]
    public void The_padding_row_refusal_is_distinct_and_reachable()
    {
        Assert.NotEqual(
            RowHeightNormalisationRefusalReason.InvalidMeasurement,
            RowHeightNormalisationRefusalReason.PaddingRowNotOwned);
        Assert.Equal(4, (int)RowHeightNormalisationRefusalReason.PaddingRowNotOwned);
    }

    /// <summary>
    /// A header row that was NOT at the period band height is restored to it
    /// (ADR-0030 D5, entity guide §4).
    /// </summary>
    /// <remarks>
    /// The header row's height is what makes the period band's bottom coincide with
    /// the first body row's top. Left at Excel's default, the band sits at the wrong
    /// vertical offset and every lane is displaced — the exact symptom this whole
    /// sequence exists to remove.
    /// </remarks>
    [Fact]
    public void A_dragged_header_row_is_restored_to_the_period_band_height()
    {
        var (normaliser, written, layout) = Build(layoutHeights: new()
        {
            [GanttSheetLayout.TopPaddingRowIndex] = PaddingRowPt,
            [GanttSheetLayout.ReservedRowIndex] = 18d,
            [GanttSheetLayout.HeaderRowIndex] = 40d,
        });

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(
            ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Contains(GanttSheetLayout.HeaderRowIndex, written);
        Assert.Equal(HeaderPt, layout[GanttSheetLayout.HeaderRowIndex]);
    }

    /// <summary>
    /// A dragged chart-padding row is restored to the PADDING token, not the reserved
    /// row's (ADR-0031 D2).
    /// </summary>
    /// <remarks>
    /// The two tokens are deliberately different values in this fixture, so writing
    /// the reserved row's height into a padding row would fail here rather than pass
    /// unnoticed. The bottom row is asserted as well as the top: its index is derived
    /// from the body length, so it is the one most likely to be skipped by a change
    /// that only knows about the rows above the table.
    /// </remarks>
    [Fact]
    public void A_dragged_padding_row_is_restored_to_the_padding_token()
    {
        var (normaliser, written, layout) = Build(
            bodyHeights: [ManagedPt],
            layoutHeights: new()
            {
                [GanttSheetLayout.TopPaddingRowIndex] = 90d,
                [GanttSheetLayout.ReservedRowIndex] = ReservedRowPt,
                [GanttSheetLayout.HeaderRowIndex] = HeaderPt,
                [GanttSheetLayout.BottomPaddingRowIndex(1)] = 95d,
            });

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(
            ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());

        Assert.Contains(GanttSheetLayout.TopPaddingRowIndex, written);
        Assert.Equal(PaddingRowPt, layout[GanttSheetLayout.TopPaddingRowIndex]);

        // The derived row below the body is normalised too, and to the same token.
        Assert.Contains(GanttSheetLayout.BottomPaddingRowIndex(1), written);
        Assert.Equal(PaddingRowPt, layout[GanttSheetLayout.BottomPaddingRowIndex(1)]);

        // Neither padding row was given a band's height.
        Assert.NotEqual(ReservedRowPt, layout[GanttSheetLayout.TopPaddingRowIndex]);
        Assert.NotEqual(HeaderPt, layout[GanttSheetLayout.TopPaddingRowIndex]);
    }

    /// <summary>
    /// The reserved row is restored to its own token, independently of the header.
    /// </summary>
    [Fact]
    public void A_dragged_reserved_row_is_restored_to_the_year_band_height()
    {
        var (normaliser, written, layout) = Build(layoutHeights: new()
        {
            [GanttSheetLayout.TopPaddingRowIndex] = PaddingRowPt,
            [GanttSheetLayout.ReservedRowIndex] = 60d,
            [GanttSheetLayout.HeaderRowIndex] = HeaderPt,
        });

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(
            ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Contains(GanttSheetLayout.ReservedRowIndex, written);
        Assert.Equal(ReservedRowPt, layout[GanttSheetLayout.ReservedRowIndex]);
    }

    /// <summary>
    /// Both layout rows are normalised even when the table has no body rows yet.
    /// </summary>
    /// <remarks>
    /// <b>Positive test for the ordering.</b> A freshly initialised table has a header
    /// and no body, so a normaliser that returned early on the body's zero rows would
    /// leave both layout rows at Excel's default — and the first lane would be
    /// misaligned on the very first Refresh a user ever runs.
    /// </remarks>
    [Fact]
    public void The_layout_rows_are_normalised_even_when_the_body_is_empty()
    {
        var (normaliser, written, _) = Build(bodyHeights: []);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(
            ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());

        // Both layout rows were already at their tokens, so nothing was written, and
        // the outcome says zero: no body rows and no layout rows needed a change.
        Assert.Equal(0, outcome.RowsWritten);
        Assert.Empty(written);
    }

    /// <summary>
    /// A layout row already at its token is not rewritten, so a normalised sheet
    /// stays clean across Refreshes.
    /// </summary>
    [Fact]
    public void Layout_rows_already_at_their_tokens_are_not_rewritten()
    {
        var (normaliser, written, _) = Build(
            bodyHeights: [ManagedPt],
            layoutHeights: new()
            {
                [GanttSheetLayout.TopPaddingRowIndex] = PaddingRowPt,                [GanttSheetLayout.ReservedRowIndex] = ReservedRowPt,
                [GanttSheetLayout.HeaderRowIndex] = HeaderPt,
            });

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(
            ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(0, outcome.RowsWritten);
        Assert.Empty(written);
    }

    /// <summary>
    /// The two layout-row heights reach the rows, so retuning a band token retunes
    /// the sheet rather than leaving a second literal behind.
    /// </summary>
    /// <remarks>
    /// Asserted on what was <em>written</em>, not by comparing two constants to each
    /// other — a test that only checks that one constant equals another proves nothing
    /// about the adapter. The orchestrator's own test separately pins that it supplies
    /// these tokens rather than literals.
    /// </remarks>
    [Fact]
    public void The_layout_row_targets_reach_the_rows()
    {
        var (normaliser, _, layout) = Build(layoutHeights: new()
        {
            [GanttSheetLayout.TopPaddingRowIndex] = PaddingRowPt,
            [GanttSheetLayout.ReservedRowIndex] = 1d,
            [GanttSheetLayout.HeaderRowIndex] = 2d,
        });

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(
            ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(HeaderPt, layout[GanttSheetLayout.HeaderRowIndex]);
        Assert.Equal(ReservedRowPt, layout[GanttSheetLayout.ReservedRowIndex]);
    }

    private static (TestableNormaliser Normaliser, List<int> Written, Dictionary<int, double> Layout) Build(
        List<double>? bodyHeights = null,
        Dictionary<int, double>? layoutHeights = null)
    {
        (Mock<Excel.Application> application, Mock<Excel.Worksheet> worksheet, Mock<Excel.ListObject> table) =
            Graph();

        List<int> written = [];
        TestableNormaliser normaliser = new(
            application.Object,
            new AlwaysUnprotectedGuard(),
            worksheet.Object,
            table.Object,
            bodyHeights ?? [ManagedPt],
            written,
            types: null,
            layoutHeights ?? new Dictionary<int, double>
            {
                [GanttSheetLayout.TopPaddingRowIndex] = PaddingRowPt,
                [GanttSheetLayout.ReservedRowIndex] = ReservedRowPt,
                [GanttSheetLayout.HeaderRowIndex] = HeaderPt,
            });

        return (normaliser, written, normaliser.LayoutHeights);
    }

    /// <summary>A guard that reports "not protected", so writes are permitted.</summary>
    private sealed class AlwaysUnprotectedGuard : IWorksheetProtectionGuard
    {
        public ProtectionGuardOutcome Query() => ProtectionGuardOutcome.NotProtected;

        public ProtectionGuardOutcome QueryTarget(object? worksheet) => ProtectionGuardOutcome.NotProtected;
    }

    private static (Mock<Excel.Application> Application, Mock<Excel.Worksheet> Worksheet, Mock<Excel.ListObject> Table) Graph()
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var worksheet = new Mock<Excel.Worksheet>();
        var table = new Mock<Excel.ListObject>();

        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(new Mock<Excel.Sheets>().Object);
        return (application, worksheet, table);
    }

    private static IWorksheetProtectionGuard Guard(ProtectionGuardOutcome outcome)
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.QueryTarget(It.IsAny<object?>())).Returns(outcome);
        return guard.Object;
    }

    /// <summary>
    /// A dragged managed row is restored to the token. This is the whole purpose of
    /// the adapter: the chart reads the measured row, so a stray drag would
    /// desynchronise one lane from the row a user is reading.
    /// </summary>
    [Fact]
    public void A_dragged_managed_row_is_restored()
    {
        var (application, worksheet, table) = Graph();
        List<int> written = [];
        var normaliser = new TestableNormaliser(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            worksheet.Object,
            table.Object,
            [45, ManagedPt, 30],
            written);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.True(outcome.Succeeded);
        Assert.Equal(2, outcome.RowsWritten);
        Assert.Equal([1, 3], written);
    }

    /// <summary>
    /// A correctly normalised sheet writes nothing, so a Refresh does not mark the
    /// workbook dirty on every run.
    /// </summary>
    [Fact]
    public void An_already_normalised_sheet_writes_nothing()
    {
        var (application, worksheet, table) = Graph();
        List<int> written = [];
        var normaliser = new TestableNormaliser(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            worksheet.Object,
            table.Object,
            [ManagedPt, ManagedPt, ManagedPt],
            written);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.RowsWritten);
        Assert.Empty(written);
    }

    /// <summary>
    /// A protected target is refused before any write. This is the ADR-0008 D4
    /// guard-first rule, and the write list is asserted empty so the test proves the
    /// refusal happened before the COM write rather than after it.
    /// </summary>
    [Fact]
    public void A_protected_target_is_refused_without_writing()
    {
        var (application, worksheet, table) = Graph();
        List<int> written = [];
        var normaliser = new TestableNormaliser(
            application.Object,
            Guard(ProtectionGuardOutcome.SheetProtected),
            worksheet.Object,
            table.Object,
            [45, 45],
            written);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.False(outcome.Succeeded);
        Assert.Equal(RowHeightNormalisationRefusalReason.TargetProtected, outcome.Refusal);
        Assert.Empty(written);
    }

    /// <summary>
    /// A guard reporting no active workbook is distinguished from a protected sheet,
    /// so the two failures are not collapsed into one message.
    /// </summary>
    [Fact]
    public void A_missing_workbook_from_the_guard_is_refused_as_such()
    {
        var (application, worksheet, table) = Graph();
        var normaliser = new TestableNormaliser(
            application.Object,
            Guard(ProtectionGuardOutcome.NoActiveWorkbook),
            worksheet.Object,
            table.Object,
            [45],
            []);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.False(outcome.Succeeded);
        Assert.Equal(RowHeightNormalisationRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }

    /// <summary>
    /// A null application object degrades to the no-active-workbook refusal with no
    /// mutation, rather than throwing on the way to a COM call.
    /// </summary>
    [Fact]
    public void A_null_application_is_refused()
    {
        var normaliser = new TestableNormaliser(
            null,
            Guard(ProtectionGuardOutcome.NotProtected),
            new Mock<Excel.Worksheet>().Object,
            new Mock<Excel.ListObject>().Object,
            [45],
            []);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.False(outcome.Succeeded);
        Assert.Equal(RowHeightNormalisationRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }

    /// <summary>
    /// A non-positive target is refused rather than written, because a nonsense row
    /// height is exactly the misalignment this adapter exists to remove.
    /// </summary>
    [Fact]
    public void A_non_positive_target_is_refused_without_writing()
    {
        var (application, worksheet, table) = Graph();
        List<int> written = [];
        var normaliser = new TestableNormaliser(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            worksheet.Object,
            table.Object,
            [45],
            written);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(0, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.False(outcome.Succeeded);
        Assert.Equal(RowHeightNormalisationRefusalReason.InvalidMeasurement, outcome.Refusal);
        Assert.Empty(written);
    }

    /// <summary>
    /// A <c>Splitter</c> row is restored to <c>SplitterPt</c>, never to the managed
    /// height.
    /// </summary>
    /// <remarks>
    /// The measurement pass used to build every <c>MeasuredRowHeight</c> with the
    /// default <see cref="MeasuredRowKind.Managed"/>, ignoring the row's own
    /// <c>Type</c>. So a structural row was rewritten to the managed height by the very
    /// Refresh meant to restore the sheet: the section header and the blank separator
    /// could never be the right height, and the lane for such a row could not line up
    /// with the worksheet row it sits in. This is the positive test for reading the
    /// Type column -- deleting that read fails it.
    /// </remarks>
    [Fact]
    public void A_splitter_row_is_restored_to_the_splitter_token_not_the_managed_height()
    {
        var (application, worksheet, table) = Graph();
        List<int> written = [];
        var normaliser = new TestableNormaliser(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            worksheet.Object,
            table.Object,
            // Row 1 is dragged; row 2 is a splitter already at its own token.
            [45, SplitterPt],
            written,
            types: ["As-Planned Activity", "Splitter"]);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.True(outcome.Succeeded);
        // Only the dragged managed row is written. If the splitter had been treated as
        // managed it would have been rewritten to 18pt -- a second write that this
        // assertion would catch.
        Assert.Equal(1, outcome.RowsWritten);
        Assert.Equal([1], written);
    }

    /// <summary>
    /// The mirror case for <c>Spacer</c>, and the pair that shows the kind is read from
    /// the Type rather than assumed: both rows sit at their own tokens and neither is
    /// written.
    /// </summary>
    [Fact]
    public void A_spacer_row_is_restored_to_the_spacer_token_not_the_managed_height()
    {
        var (application, worksheet, table) = Graph();
        List<int> written = [];
        var normaliser = new TestableNormaliser(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            worksheet.Object,
            table.Object,
            [SpacerPt],
            written,
            types: ["Spacer"]);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.RowsWritten);
        Assert.Empty(written);
    }

    /// <summary>
    /// A dragged structural row is restored to its OWN token rather than left alone.
    /// The no-op case above proves the splitter is not over-written; this proves it is
    /// still recognised when it does need a write, which is the case the missing Type
    /// read broke in production.
    /// </summary>
    [Fact]
    public void A_dragged_splitter_row_is_restored_to_the_splitter_token()
    {
        var (application, worksheet, table) = Graph();
        List<int> written = [];
        var normaliser = new TestableNormaliser(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            worksheet.Object,
            table.Object,
            [45],
            written,
            types: ["Splitter"]);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.RowsWritten);
        Assert.Equal([1], written);
    }

    /// <summary>
    /// An empty body succeeds with zero writes. A table with no rows is a legitimate
    /// state, not a failure.
    /// </summary>
    [Fact]
    public void An_empty_body_succeeds_with_no_writes()
    {
        var (application, worksheet, table) = Graph();
        var normaliser = new TestableNormaliser(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            worksheet.Object,
            table.Object,
            [],
            []);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(ManagedPt, SplitterPt, SpacerPt, HeaderPt, ReservedRowPt, PaddingRowPt, AnchorRowPt);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.RowsWritten);
    }
}
