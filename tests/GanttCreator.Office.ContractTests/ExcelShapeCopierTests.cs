using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R6.6: the copy adapter resolves the Gantt worksheet, proves the staging identity by
/// recomputation, and copies only the owned group - never a same-named user shape.
/// </summary>
/// <remarks>
/// The modelled host is a real <see cref="ExcelShapeCopier"/> whose only host interactions
/// are a dictionary of shape mocks and a recorded copy, the same subclass-seam pattern the
/// grouping adapter tests use. Every refusal arm has a positive test (AGENTS.md validator
/// rule).
/// </remarks>
public sealed class ExcelShapeCopierTests
{
    [Fact]
    public void Copying_copies_the_owned_staged_group()
    {
        ModelledCopier copier = NewCopier();
        copier.AddShape(ShapeStagingTag.GroupName, ShapeStagingTag.ForStagedName(ShapeStagingTag.GroupName));

        ShapeWriteOutcome outcome = copier.CopyStaged(ShapeStagingTag.GroupName);

        Assert.True(outcome.Succeeded, "Copy refused: " + outcome.Refusal);
        Assert.Equal([ShapeStagingTag.GroupName], copier.Copied);
    }

    [Fact]
    public void An_absent_group_refuses_not_found_before_any_copy()
    {
        ModelledCopier copier = NewCopier();

        ShapeWriteOutcome outcome = copier.CopyStaged(ShapeStagingTag.GroupName);

        Assert.Equal(ShapeWriteRefusal.NotFound, outcome.Refusal);
        Assert.Empty(copier.Copied);
    }

    [Fact]
    public void An_unowned_shape_that_shares_the_name_is_refused_and_never_copied()
    {
        // Ownership is the recomputed tag, not the name: a user's drawing called the
        // group's name must survive and must never reach the clipboard.
        ModelledCopier copier = NewCopier();
        copier.AddShape(ShapeStagingTag.GroupName, "A user's description");

        ShapeWriteOutcome outcome = copier.CopyStaged(ShapeStagingTag.GroupName);

        Assert.Equal(ShapeWriteRefusal.NotFound, outcome.Refusal);
        Assert.Empty(copier.Copied);
    }

    [Fact]
    public void A_protected_target_refuses_before_any_copy()
    {
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        _ = guard.Setup(g => g.QueryTarget(It.IsAny<object>()))
            .Returns(ProtectionGuardOutcome.SheetProtected);
        ModelledCopier copier = new(StagingSchemeFixtures.ActiveApplication().Object, guard.Object);
        copier.AddShape(ShapeStagingTag.GroupName, ShapeStagingTag.ForStagedName(ShapeStagingTag.GroupName));

        ShapeWriteOutcome outcome = copier.CopyStaged(ShapeStagingTag.GroupName);

        Assert.Equal(ShapeWriteRefusal.TargetProtected, outcome.Refusal);
        Assert.Empty(copier.Copied);
    }

    [Fact]
    public void The_entry_guard_is_consulted_before_any_workbook_access()
    {
        // ADR-0008 D4: the guard is the first host check, and the copier's entry point
        // carries its own query.
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.SheetProtected);
        ModelledCopier copier = new(StagingSchemeFixtures.ActiveApplication().Object, guard.Object);

        Assert.Equal(
            ShapeWriteRefusal.TargetProtected,
            copier.CopyStaged(ShapeStagingTag.GroupName).Refusal);
        Assert.Empty(copier.Copied);
    }

    [Fact]
    public void A_missing_workbook_refuses()
    {
        ModelledCopier copier = new(null, StagingSchemeFixtures.ClearGuard().Object);

        Assert.Equal(
            ShapeWriteRefusal.NoActiveWorkbook,
            copier.CopyStaged(ShapeStagingTag.GroupName).Refusal);
    }

    [Fact]
    public void A_host_copy_refusal_is_host_rejected()
    {
        ModelledCopier copier = NewCopier(throwOnCopy: true);
        copier.AddShape(ShapeStagingTag.GroupName, ShapeStagingTag.ForStagedName(ShapeStagingTag.GroupName));

        ShapeWriteOutcome outcome = copier.CopyStaged(ShapeStagingTag.GroupName);

        Assert.Equal(ShapeWriteRefusal.HostRejected, outcome.Refusal);
        Assert.Equal([ShapeStagingTag.GroupName], copier.Copied);
    }

    [Fact]
    public void A_blank_name_is_refused()
    {
        ModelledCopier copier = NewCopier();

        Assert.Equal(ShapeWriteRefusal.BlankIdentifier, copier.CopyStaged(" ").Refusal);
        Assert.Empty(copier.Copied);
    }

    private static ModelledCopier NewCopier(bool throwOnCopy = false) =>
        new(StagingSchemeFixtures.ActiveApplication().Object, StagingSchemeFixtures.ClearGuard().Object)
        {
            ThrowOnCopy = throwOnCopy,
        };

    private sealed class ModelledCopier(object? application, IWorksheetProtectionGuard? guard)
        : ExcelShapeCopier(application, guard)
    {
        private readonly Dictionary<string, Mock<Excel.Shape>> _shapes = new(StringComparer.Ordinal);

        public bool ThrowOnCopy { get; init; }

        public List<string> Copied { get; } = [];

        public void AddShape(string name, string? alternativeText = null)
        {
            Mock<Excel.Shape> mock = new();
            _ = mock.SetupProperty(shape => shape.Name, name);
            _ = mock.SetupProperty(shape => shape.AlternativeText, alternativeText ?? string.Empty);
            _shapes[name] = mock;
        }

        internal override Excel._Worksheet? FindGanttWorksheet(Excel.Sheets sheets) =>
            new Mock<Excel._Worksheet>().Object;

        internal override Excel.Shapes? GetShapes(Excel._Worksheet sheet) =>
            new Mock<Excel.Shapes>().Object;

        internal override int GetShapeCount(Excel.Shapes shapes) => _shapes.Count;

        internal override Excel.Shape? GetShapeAt(Excel.Shapes shapes, int index) =>
            index >= 1 && index <= _shapes.Count
                ? _shapes.Values.ElementAt(index - 1).Object
                : null;

        internal override Excel.Shape? FindShapeByName(Excel.Shapes shapes, string name) =>
            _shapes.TryGetValue(name, out Mock<Excel.Shape>? shape) ? shape.Object : null;

        internal override void CopyShape(Excel.Shape shape)
        {
            Copied.Add(shape.Name);
            if (ThrowOnCopy)
            {
                throw HostFailure();
            }
        }

        /// <summary>
        /// The exception the host actually raises, built the way the interop boundary builds
        /// it rather than constructed directly (CA2201: COMException is reserved).
        /// </summary>
        /// <returns>The modelled host exception.</returns>
        private static COMException HostFailure() =>
            (COMException)Marshal.GetExceptionForHR(unchecked((int)0x800A03EC))!;
    }
}
