using System.Runtime.InteropServices;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// The unlock command acts only on one shape this add-in owns, and the Ribbon asks the same port
/// whether to enable it — so every refusal has to be typed rather than thrown (ADR-0055 clause 6).
/// </summary>
public class ExcelSelectedShapeReaderTests
{
    private const string Owned = "G-0123456789abcdef0123456789abcdef:bar";

    [Fact]
    public void No_application_is_its_own_refusal()
    {
        Assert.Equal(
            SelectedShapeRefusalReason.NoApplication,
            new ExcelSelectedShapeReader(application: null).Read().Refusal);
    }

    [Fact]
    public void A_cell_range_selection_is_not_a_shape()
    {
        var range = new Mock<Excel.Range>();
        var reader = new TestableReader(range.Object);

        SelectedShapeOutcome outcome = reader.Read();

        Assert.False(outcome.Succeeded);
        Assert.Equal(SelectedShapeRefusalReason.SelectionIsNotASingleShape, outcome.Refusal);
    }

    [Fact]
    public void An_owned_shape_resolves_to_its_primitive_id()
    {
        Excel.Shape shape = NewShape(Owned, ShapeOwnershipTag.ForPrimitiveId(Owned));
        var reader = new TestableReader(shape);

        SelectedShapeOutcome outcome = reader.Read();

        Assert.True(outcome.Succeeded);
        Assert.Equal(Owned, outcome.PrimitiveId);
    }

    /// <summary>
    /// The tag, not the name, proves ownership: a user can rename their own shape to anything.
    /// </summary>
    [Fact]
    public void A_shape_with_our_name_but_another_tag_is_refused()
    {
        Excel.Shape shape = NewShape(Owned, "a note the user typed");
        var reader = new TestableReader(shape);

        SelectedShapeOutcome outcome = reader.Read();

        Assert.False(outcome.Succeeded);
        Assert.Equal(SelectedShapeRefusalReason.ShapeNotOwned, outcome.Refusal);
    }

    [Fact]
    public void A_shape_with_no_tag_is_refused()
    {
        Excel.Shape shape = NewShape(Owned, alternativeText: null);
        var reader = new TestableReader(shape);

        Assert.Equal(SelectedShapeRefusalReason.ShapeNotOwned, reader.Read().Refusal);
    }

    /// <summary>
    /// A grouped shape, a chart selection, or a protected sheet can refuse the read; the Ribbon must
    /// see a refusal rather than a host error dialog.
    /// </summary>
    [Fact]
    public void A_host_that_refuses_the_read_is_reported_as_HostRejected()
    {
        var shape = new Mock<Excel.Shape>();
        // CA2201: COMException is the reserved type the host actually raises, and the adapter catches
        // it specifically, so the test must raise that type rather than a stand-in.
#pragma warning disable CA2201
        shape.SetupGet(s => s.Name).Throws(new COMException("refused"));
#pragma warning restore CA2201
        var reader = new TestableReader(shape.Object);

        Assert.Equal(SelectedShapeRefusalReason.HostRejected, reader.Read().Refusal);
    }

    private static Excel.Shape NewShape(string name, string? alternativeText)
    {
        var shape = new Mock<Excel.Shape>();
        _ = shape.SetupGet(s => s.Name).Returns(name);
        _ = shape.SetupGet(s => s.AlternativeText).Returns(alternativeText!);
        return shape.Object;
    }

    /// <summary>
    /// Substitutes the one COM property the reader cannot express in a Moq expression tree. The
    /// application is a mock rather than null because a null application short-circuits to
    /// <see cref="SelectedShapeRefusalReason.NoApplication"/> before the selection is ever read.
    /// </summary>
    private sealed class TestableReader(object? selection)
        : ExcelSelectedShapeReader(new Mock<Excel.Application>().Object)
    {
        internal override object? GetSelection(Excel.Application application) => selection;
    }
}
