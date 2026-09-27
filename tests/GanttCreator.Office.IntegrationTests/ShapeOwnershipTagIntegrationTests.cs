using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R4.1's outstanding Step-0 obligation, discharged live: does Excel truncate
/// the ADR-0019 ownership tag written to a shape's alternative text?
/// </summary>
/// <remarks>
/// <para>
/// ADR-0019 D2 made the tag bounded precisely because a truncation would break
/// the R4.7 reconciliation key and the R4.8 ownership filter at the same time,
/// and because the pre-agreed remedy for a truncating host is a narrower hash.
/// This test therefore writes a real tag through the real
/// <see cref="ExcelShapeWriter"/>, reads it back from the live shape, and
/// compares it byte-for-byte. A pass closes the obligation; a failure narrows
/// the hash and is recorded here.
/// </para>
/// <para>
/// The probe deliberately uses a <em>shared-owner</em> identifier, the
/// ADR-0017 case that motivated bounding the tag in the first place, so the
/// test would catch a host that truncates by length rather than by content.
/// </para>
/// </remarks>
public class ShapeOwnershipTagIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task The_ownership_tag_round_trips_through_a_live_shape_unaltered()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            // ComScope releases the proxies this test body acquires, so the test
            // does not add to the run's COM-proxy leak signal (the gate's ratchet
            // counts exactly this).
            using var scope = new OfficeFixture.ComScope();
            Excel.Workbook workbook = scope.Track(fixture.CreateWorkbook());
            var initialiser = new ExcelWorkbookInitialiser(fixture.Excel);
            WorkbookInitialiseOutcome initialised = initialiser.Initialise();
            Assert.True(initialised.Succeeded, $"Initialise refused: {initialised.Refusal}");

            var writer = new ExcelShapeWriter(fixture.Excel);

            // A shared-owner identifier: the unbounded case ADR-0017 introduced
            // and the reason the tag hashes rather than embeds the identifier.
// A shared-owner identifier that is realistic but still exercises the
            // unbounded-shape case ADR-0017 introduced. It is kept under Excel's
            // 255-character shape-name limit, which is a SEPARATE limit from the
            // tag's: the name carries the identifier, the alternative text carries
            // only the bounded hash. Both limits are asserted below.
            string sharedOwnerId = string.Join("|", Enumerable.Range(1, 20).Select(i => "row-" + i));
            string primitiveId = sharedOwnerId + ":bar";
            string expectedTag = ShapeOwnershipTag.ForPrimitiveId(primitiveId);

            _output.WriteLine("identifier length : " + primitiveId.Length);
            _output.WriteLine("tag length         : " + expectedTag.Length);
            _output.WriteLine("expected tag       : " + expectedTag);

            // The tag is bounded by construction: 22 prefix characters plus a
            // 16-character hash, regardless of how long the identifier is. This
            // assertion is what makes the probe meaningful - a tag that grew with
            // the identifier could be truncated by a host without this test
            // noticing, because the test's own expectation would grow too.
            string farLongerId = string.Join("|", Enumerable.Range(1, 2000).Select(i => "r" + i));
            Assert.Equal(22 + 16, expectedTag.Length);
            Assert.Equal(22 + 16, ShapeOwnershipTag.ForPrimitiveId(farLongerId).Length);
            Assert.True(farLongerId.Length > primitiveId.Length * 10);

            ShapeWriteOutcome outcome = writer.Create(
                new OfficeShapeRequest(
                    primitiveId,
                    OfficeShapeKind.Rectangle,
                    new OfficeShapeGeometry(Bounds: new RectD(10, 10, 100, 30)),
                    ZLayer.ActivityBody));

            Assert.True(outcome.Succeeded, $"Create refused: {outcome.Refusal}");

            // Read the alternative text back off the live shape, not from the
            // value the adapter intended to write.
            var sheet = (Excel.Worksheet)scope.Track(workbook.Sheets[GanttWorkbookContract.GanttSheetLabel]);
            Excel.Shape shape = scope.Track(sheet.Shapes.Item(primitiveId));
            string? actualTag = shape.AlternativeText;

            _output.WriteLine("actual tag         : " + actualTag);

            Assert.Equal(expectedTag, actualTag);
            Assert.Equal(primitiveId, shape.Name);

            // The shape must also be recognised as owned by the live read path,
            // or the filter R4.8 depends on would never match it.
            Assert.True(ShapeOwnershipTag.IsOwnedTag(actualTag));

            // The two members together are the R4.7 reconciliation key.
            Assert.Equal([primitiveId], writer.ListOwned());

            // Delete the shape but leave the workbook open for the fixture to
            // close. Closing it here would disconnect the very proxies the
            // ComScope releases at the end of this body, and the resulting
            // RPC_E_DISCONNECTED would fail the test for a teardown reason that
            // has nothing to do with the ownership tag.
            shape.Delete();
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }
}
