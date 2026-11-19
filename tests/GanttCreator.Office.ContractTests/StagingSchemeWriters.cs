using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Excel = Microsoft.Office.Interop.Excel;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// One modelled worksheet shape collection, shared by two ownership schemes so the
/// live chart and a staged composition can be driven on the same worksheet without a
/// host (R6.2 D4, reused by R6.3's refresh-during-staging scenario).
/// </summary>
internal sealed class SharedShapes
{
    /// <summary>Gets the modelled shape collection, in creation order.</summary>
    public List<Excel.Shape> Shapes { get; } = [];
}

/// <summary>
/// A real <see cref="ExcelShapeWriter"/> whose only host interaction is one shared,
/// modelled shape collection, constructed with an explicit ownership scheme.
/// </summary>
/// <remarks>
/// The scheme is the production routing under test: the live writer stamps identity
/// names and the live tag, the staging writer stamps the staging prefix and tag, and
/// each one's <c>ListOwned</c> recomputation is blind to the other family.
/// </remarks>
internal sealed class SchemeWriter(SharedShapes sheet, ShapeOwnershipScheme scheme)
    : ExcelShapeWriter(StagingSchemeFixtures.ActiveApplication().Object, StagingSchemeFixtures.ClearGuard().Object, trace: null, scheme)
{
    internal override Excel._Worksheet? FindGanttWorksheet(Excel.Sheets sheets) =>
        new Mock<Excel._Worksheet>().Object;

    internal override Excel.Shapes? GetShapes(Excel._Worksheet sheet) =>
        new Mock<Excel.Shapes>().Object;

    internal override int GetShapeCount(Excel.Shapes shapes) => sheet.Shapes.Count;

    internal override Excel.Shape? GetShapeAt(Excel.Shapes shapes, int index) =>
        index >= 1 && index <= sheet.Shapes.Count ? sheet.Shapes[index - 1] : null;

    internal override Excel.Shape? FindShapeByName(Excel.Shapes shapes, string name) =>
        sheet.Shapes.FirstOrDefault(shape => string.Equals(shape.Name, name, StringComparison.Ordinal));

    internal override void ApplyStyle(Excel.Shape shape, OfficeShapeRequest request)
    {
        // The style write is exercised by the depth tests; here the modelled shape has
        // no fill or line, and the ownership stamp is what these tests observe.
    }

    internal override Excel.Shape? AddShape(Excel.Shapes shapes, OfficeShapeRequest request)
    {
        Mock<Excel.Shape> mock = new();
        _ = mock.SetupProperty(shape => shape.Name, string.Empty);
        _ = mock.SetupProperty(shape => shape.AlternativeText, string.Empty);
        _ = mock.SetupProperty(shape => shape.Placement, Excel.XlPlacement.xlMoveAndSize);
        _ = mock.SetupProperty(shape => shape.Locked, false);

        Excel.Shape created = mock.Object;
        _ = mock.Setup(shape => shape.Delete()).Callback(() => sheet.Shapes.Remove(created));

        sheet.Shapes.Add(created);
        return created;
    }
}

/// <summary>Shared construction helpers for the scheme-writer tests.</summary>
internal static class StagingSchemeFixtures
{
    /// <summary>A minimal rectangle request for the named primitive.</summary>
    /// <param name="id">The primitive identifier.</param>
    /// <returns>The shape request.</returns>
    public static OfficeShapeRequest RectangleRequest(string id) =>
        new(
            id,
            OfficeShapeKind.Rectangle,
            new OfficeShapeGeometry(Bounds: new RectD(10, 20, 100, 30)),
            ZLayer.ActivityBody);

    /// <summary>A protection guard that reports the worksheet unprotected.</summary>
    /// <returns>The guard mock.</returns>
    public static Mock<IWorksheetProtectionGuard> ClearGuard()
    {
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        _ = guard.Setup(g => g.QueryTarget(It.IsAny<object>())).Returns(ProtectionGuardOutcome.NotProtected);
        return guard;
    }

    /// <summary>An application whose active workbook exposes a sheet collection.</summary>
    /// <returns>The application mock.</returns>
    public static Mock<Excel.Application> ActiveApplication()
    {
        Mock<Excel.Workbook> workbook = new();
        _ = workbook.Setup(w => w.Sheets).Returns(new Mock<Excel.Sheets>().Object);
        Mock<Excel.Application> application = new();
        _ = application.Setup(a => a.ActiveWorkbook).Returns(workbook.Object);
        return application;
    }
}
