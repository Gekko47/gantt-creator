using System.Runtime.InteropServices;
using GanttCreator.Core;
using Microsoft.Office.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R6.4: the grouping adapter selects exactly the staged names (never a sheet scan), proves
/// the staging identity on the group it stamps, and restores the children when the host
/// cannot finish the operation.
/// </summary>
/// <remarks>
/// The modelled host is a real <see cref="ExcelShapeGrouper"/> whose only host interaction
/// is a dictionary of shape mocks, the same subclass-seam pattern the writer tests use.
/// Every refusal arm has a positive test that constructs the bad input (AGENTS.md validator
/// rule), including the single-shape range Excel returns as the shape itself.
/// </remarks>
public sealed class ExcelShapeGrouperTests
{
    private static readonly string[] _staged =
    [
        ShapeStagingTag.NamePrefix + "chart:frame:left",
        ShapeStagingTag.NamePrefix + "chart:bar:0",
    ];

    [Fact]
    public void Grouping_selects_exactly_the_staged_names_and_stamps_the_group()
    {
        ModelledGrouper grouper = NewGrouper();
        foreach (string name in _staged)
        {
            grouper.AddShape(name);
        }

        // The unowned sentinel sits on the same sheet and overlaps the staged envelope;
        // selection is by exact name, so it can never be part of the range.
        grouper.AddShape("User note");

        ShapeGroupReadback readback = grouper.GroupStaged(_staged);

        Assert.True(readback.Succeeded, "Grouping refused: " + readback.Refusal);
        Assert.Equal(_staged, grouper.RequestedNames);
        Assert.DoesNotContain("User note", grouper.RequestedNames);

        // D1: the group carries the deterministic staging identity and the temporary
        // placement class, exactly as every staged child does.
        Assert.Equal(ShapeStagingTag.GroupName, readback.GroupName);
        Assert.Equal(ShapeStagingTag.ForStagedName(ShapeStagingTag.GroupName), grouper.GroupShape.AlternativeText);
        Assert.True(ShapeStagingTag.IsStagingTag(grouper.GroupShape.AlternativeText));
        Assert.Equal(Excel.XlPlacement.xlMove, grouper.GroupShape.Placement);
        Assert.True(grouper.GroupShape.Locked);

        // The read-back is the host's own order and envelope, not a restatement.
        Assert.Equal(_staged, readback.ChildNames);
        Assert.Equal<RectD?>(grouper.GroupBounds, readback.Bounds);
    }

    [Fact]
    public void An_absent_staged_name_refuses_before_any_grouping()
    {
        ModelledGrouper grouper = NewGrouper();
        grouper.AddShape(_staged[0]);

        ShapeGroupReadback readback = grouper.GroupStaged([_staged[0], _staged[1]]);

        Assert.Equal(ShapeWriteRefusal.NotFound, readback.Refusal);
        Assert.Equal(0, grouper.GroupCalls);
        Assert.Equal(0, grouper.StampCalls);
    }

    [Fact]
    public void A_single_shape_range_is_not_reported_as_a_group()
    {
        // Excel returns the shape itself for a one-shape range. Reporting that as a group
        // would hand the stage an empty GroupItems collection with no explanation.
        ModelledGrouper grouper = NewGrouper(groupToSingleShape: true);
        grouper.AddShape(_staged[0]);

        ShapeGroupReadback readback = grouper.GroupStaged([_staged[0]]);

        Assert.Equal(ShapeWriteRefusal.HostRejected, readback.Refusal);
        Assert.Equal(1, grouper.GroupCalls);
        Assert.Equal(0, grouper.StampCalls);
        Assert.Empty(grouper.Ungrouped);
    }

    [Fact]
    public void A_host_grouping_refusal_is_host_rejected()
    {
        ModelledGrouper grouper = NewGrouper(refuseGrouping: true);
        foreach (string name in _staged)
        {
            grouper.AddShape(name);
        }

        ShapeGroupReadback readback = grouper.GroupStaged(_staged);

        Assert.Equal(ShapeWriteRefusal.HostRejected, readback.Refusal);
        Assert.Equal(0, grouper.StampCalls);
    }

    [Fact]
    public void Non_finite_bounds_restore_the_children_and_refuse()
    {
        ModelledGrouper grouper = NewGrouper(nonFiniteBounds: true);
        foreach (string name in _staged)
        {
            grouper.AddShape(name);
        }

        ShapeGroupReadback readback = grouper.GroupStaged(_staged);

        Assert.Equal(ShapeWriteRefusal.HostRejected, readback.Refusal);
        Assert.Null(readback.Bounds);
        Assert.Equal([ShapeStagingTag.GroupName], grouper.Ungrouped);
    }

    [Fact]
    public void A_stamp_failure_restores_the_children_and_refuses()
    {
        // An unstamped group is worse than no group: it is invisible to the staging
        // ownership scan, so it must be ungrouped before the refusal is reported.
        ModelledGrouper grouper = NewGrouper(throwOnStamp: true);
        foreach (string name in _staged)
        {
            grouper.AddShape(name);
        }

        ShapeGroupReadback readback = grouper.GroupStaged(_staged);

        Assert.Equal(ShapeWriteRefusal.HostRejected, readback.Refusal);
        Assert.Equal(1, grouper.StampCalls);

        // The stamp never landed, so the restore ungroups the shape by the name the host
        // assigned ("Group 1"), not by the staging name it would have carried.
        Assert.Equal(["Group 1"], grouper.Ungrouped);
    }

    [Fact]
    public void A_protected_target_refuses_before_any_grouping()
    {
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.QueryTarget(It.IsAny<object>())).Returns(ProtectionGuardOutcome.SheetProtected);
        ModelledGrouper grouper = NewGrouper(guard: guard.Object);
        foreach (string name in _staged)
        {
            grouper.AddShape(name);
        }

        ShapeGroupReadback readback = grouper.GroupStaged(_staged);

        Assert.Equal(ShapeWriteRefusal.TargetProtected, readback.Refusal);
        Assert.Equal(0, grouper.GroupCalls);
    }

    [Fact]
    public void A_missing_workbook_or_worksheet_refuses()
    {
        // The guard is cleared so the resolution path is reached: the real guard fails
        // closed for a null application, and that arm has its own positive test below.
        ModelledGrouper noApplication = new(null, StagingSchemeFixtures.ClearGuard().Object);
        Assert.Equal(
            ShapeWriteRefusal.NoActiveWorkbook,
            noApplication.GroupStaged(_staged).Refusal);

        ModelledGrouper noWorksheet = NewGrouper(worksheetPresent: false);
        Assert.Equal(
            ShapeWriteRefusal.WorksheetMissing,
            noWorksheet.GroupStaged(_staged).Refusal);
    }

    [Fact]
    public void The_entry_guard_is_consulted_before_any_workbook_access()
    {
        // ADR-0008 D4: the guard is the first host check in every mutating adapter, and
        // both mutating entry points carry their own query.
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.SheetProtected);
        ModelledGrouper grouper = NewGrouper(guard: guard.Object);
        foreach (string name in _staged)
        {
            grouper.AddShape(name);
        }

        ShapeGroupReadback readback = grouper.GroupStaged(_staged);
        Assert.Equal(ShapeWriteRefusal.TargetProtected, readback.Refusal);
        Assert.Equal(0, grouper.GroupCalls);

        Assert.Equal(
            ShapeWriteRefusal.TargetProtected,
            grouper.UngroupStaged(ShapeStagingTag.GroupName).Refusal);
        Assert.Empty(grouper.Ungrouped);
    }

    [Fact]
    public void Blank_and_empty_inputs_are_refused()
    {
        ModelledGrouper grouper = NewGrouper();

        Assert.Equal(ShapeWriteRefusal.NullRequest, grouper.GroupStaged(null!).Refusal);
        Assert.Equal(ShapeWriteRefusal.BlankIdentifier, grouper.GroupStaged([]).Refusal);
        Assert.Equal(ShapeWriteRefusal.BlankIdentifier, grouper.GroupStaged([" "]).Refusal);
        Assert.Equal(0, grouper.GroupCalls);
    }

    [Fact]
    public void Ungroup_restores_a_staging_owned_group()
    {
        ModelledGrouper grouper = NewGrouper();
        grouper.AddShape(ShapeStagingTag.GroupName, ShapeStagingTag.ForStagedName(ShapeStagingTag.GroupName));

        ShapeWriteOutcome outcome = grouper.UngroupStaged(ShapeStagingTag.GroupName);

        Assert.True(outcome.Succeeded, "Ungroup refused: " + outcome.Refusal);
        Assert.Equal([ShapeStagingTag.GroupName], grouper.Ungrouped);
    }

    [Fact]
    public void Ungroup_refuses_an_absent_or_unowned_group()
    {
        ModelledGrouper grouper = NewGrouper();

        Assert.Equal(
            ShapeWriteRefusal.NotFound,
            grouper.UngroupStaged(ShapeStagingTag.GroupName).Refusal);

        // A user's shape that happens to carry the group's name is refused, never touched:
        // ownership is the recomputed tag, not the name.
        grouper.AddShape(ShapeStagingTag.GroupName, "A user's description");
        Assert.Equal(
            ShapeWriteRefusal.NotFound,
            grouper.UngroupStaged(ShapeStagingTag.GroupName).Refusal);
        Assert.Empty(grouper.Ungrouped);

        Assert.Equal(ShapeWriteRefusal.BlankIdentifier, grouper.UngroupStaged(" ").Refusal);
    }

    [Fact]
    public void An_ungroup_host_refusal_is_host_rejected()
    {
        ModelledGrouper grouper = NewGrouper(throwOnUngroup: true);
        grouper.AddShape(ShapeStagingTag.GroupName, ShapeStagingTag.ForStagedName(ShapeStagingTag.GroupName));

        ShapeWriteOutcome outcome = grouper.UngroupStaged(ShapeStagingTag.GroupName);

        Assert.Equal(ShapeWriteRefusal.HostRejected, outcome.Refusal);
        Assert.Equal([ShapeStagingTag.GroupName], grouper.Ungrouped);
    }

    private static ModelledGrouper NewGrouper(
        bool worksheetPresent = true,
        bool refuseGrouping = false,
        bool groupToSingleShape = false,
        bool nonFiniteBounds = false,
        bool throwOnStamp = false,
        bool throwOnUngroup = false,
        IWorksheetProtectionGuard? guard = null) =>
        new(StagingSchemeFixtures.ActiveApplication().Object, guard ?? StagingSchemeFixtures.ClearGuard().Object)
        {
            WorksheetPresent = worksheetPresent,
            RefuseGrouping = refuseGrouping,
            GroupToSingleShape = groupToSingleShape,
            NonFiniteBounds = nonFiniteBounds,
            ThrowOnStamp = throwOnStamp,
            ThrowOnUngroup = throwOnUngroup,
        };

    /// <summary>
    /// A real grouper whose only host interaction is a dictionary of shape mocks.
    /// </summary>
    private sealed class ModelledGrouper(object? application, IWorksheetProtectionGuard? guard)
        : ExcelShapeGrouper(application, guard)
    {
        private readonly Dictionary<string, Mock<Excel.Shape>> _shapes = new(StringComparer.Ordinal);
        private Mock<Excel.Shape>? _group;

        public bool WorksheetPresent { get; init; } = true;

        public bool RefuseGrouping { get; init; }

        public bool GroupToSingleShape { get; init; }

        public bool NonFiniteBounds { get; init; }

        public bool ThrowOnStamp { get; init; }

        public bool ThrowOnUngroup { get; init; }

        public RectD GroupBounds { get; init; } = new(500, 60, 320, 280);

        public IReadOnlyList<string>? ChildrenOverride { get; init; }

        public List<string> RequestedNames { get; } = [];

        public List<string> Ungrouped { get; } = [];

        public int GroupCalls { get; private set; }

        public int StampCalls { get; private set; }

        public Excel.Shape GroupShape => _group!.Object;

        public void AddShape(string name, string? alternativeText = null)
        {
            Mock<Excel.Shape> mock = new();
            mock.SetupProperty(shape => shape.Name, name);
            mock.SetupProperty(shape => shape.AlternativeText, alternativeText ?? string.Empty);
            mock.SetupProperty(shape => shape.Placement, Excel.XlPlacement.xlMoveAndSize);
            mock.SetupProperty(shape => shape.Locked, false);
            _shapes[name] = mock;
        }

        internal override Excel._Worksheet? FindGanttWorksheet(Excel.Sheets sheets) =>
            WorksheetPresent ? new Mock<Excel._Worksheet>().Object : null;

        internal override Excel.Shapes? GetShapes(Excel._Worksheet sheet) =>
            new Mock<Excel.Shapes>().Object;

        internal override int GetShapeCount(Excel.Shapes shapes) => _shapes.Count;

        internal override Excel.Shape? GetShapeAt(Excel.Shapes shapes, int index) =>
            index >= 1 && index <= _shapes.Count
                ? _shapes.Values.ElementAt(index - 1).Object
                : null;

        internal override Excel.Shape? FindShapeByName(Excel.Shapes shapes, string name) =>
            _shapes.TryGetValue(name, out Mock<Excel.Shape>? shape) ? shape.Object : null;

        internal override Excel.ShapeRange GetShapeRange(
            Excel.Shapes shapes,
            IReadOnlyList<string> shapeNames)
        {
            RequestedNames.Clear();
            RequestedNames.AddRange(shapeNames);
            return new Mock<Excel.ShapeRange>().Object;
        }

        internal override Excel.Shape? GroupRange(Excel.ShapeRange range)
        {
            GroupCalls++;
            if (RefuseGrouping)
            {
                return null;
            }

            Mock<Excel.Shape> group = new();
            group.SetupProperty(shape => shape.Name, "Group 1");
            group.SetupProperty(shape => shape.AlternativeText, string.Empty);
            group.SetupProperty(shape => shape.Placement, Excel.XlPlacement.xlMoveAndSize);
            group.SetupProperty(shape => shape.Locked, false);
            _group = group;
            return group.Object;
        }

        internal override MsoShapeType ReadShapeType(Excel.Shape shape) =>
            _group is not null && ReferenceEquals(shape, _group.Object) && !GroupToSingleShape
                ? MsoShapeType.msoGroup
                : MsoShapeType.msoAutoShape;

        internal override Excel.GroupShapes? GetGroupItems(Excel.Shape group) =>
            new Mock<Excel.GroupShapes>().Object;

        internal override int GetGroupItemCount(Excel.GroupShapes items) => Children.Count;

        internal override string GetGroupItemName(Excel.GroupShapes items, int index) =>
            Children[index - 1];

        internal override double ReadLeft(Excel.Shape shape) =>
            NonFiniteBounds ? double.NaN : GroupBounds.X;

        internal override double ReadTop(Excel.Shape shape) => GroupBounds.Y;

        internal override double ReadWidth(Excel.Shape shape) => GroupBounds.Width;

        internal override double ReadHeight(Excel.Shape shape) => GroupBounds.Height;

        internal override void ApplyGroupStamp(Excel.Shape group)
        {
            StampCalls++;
            if (ThrowOnStamp)
            {
                throw HostFailure();
            }

            base.ApplyGroupStamp(group);
        }

        internal override void UngroupShape(Excel.Shape group)
        {
            Ungrouped.Add(group.Name);
            if (ThrowOnUngroup)
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

        private IReadOnlyList<string> Children => ChildrenOverride ?? RequestedNames;
    }
}
