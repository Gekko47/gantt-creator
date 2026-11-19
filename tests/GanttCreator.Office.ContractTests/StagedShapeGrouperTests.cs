using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R6.4: the grouping stage verifies what the host produced before reporting success -
/// membership exactly the staged set (D1), child order equal to the scene order (D2), and
/// the group envelope equal to the R6.3 staged bounds - and restores the children on any
/// mismatch.
/// </summary>
/// <remarks>
/// Every refusal arm has a positive test that constructs the bad input (AGENTS.md
/// validator rule), and each verification refusal asserts the rollback ungroup actually
/// ran, because a group left behind would block every later composition through the
/// composer's <c>StagingAlreadyPopulated</c> pre-flight.
/// </remarks>
public sealed class StagedShapeGrouperTests
{
    private static readonly RectD _bounds = new(500, 60, 320, 280);

    private static readonly string[] _ids = ["chart:frame:left", "chart:bar:0", "chart:panel-text:0"];

    [Fact]
    public void The_staged_set_groups_in_scene_order_and_verifies()
    {
        FakeGroupingPort port = new() { Bounds = _bounds };
        StagingCompositionOutcome staged = Staged(_ids, _bounds);

        StagedGroupOutcome outcome = StagedShapeGrouper.Group(staged, port);

        Assert.True(outcome.Succeeded, "Grouping refused: " + outcome.Refusal);

        // D2: the expected names are handed over in the scene's creation order.
        string[] expected = [.. _ids.Select(id => ShapeStagingTag.NamePrefix + id)];
        Assert.Equal(expected, port.RequestedNames);

        // D1: the group carries the deterministic staging identity.
        Assert.Equal(ShapeStagingTag.GroupName, outcome.GroupName);
        Assert.Equal(expected, outcome.ChildOrder);
        Assert.Equal(_bounds, outcome.Bounds);
        Assert.Empty(port.UngroupedNames);
    }

    [Fact]
    public void The_group_bounds_pin_reuses_the_staged_envelope_check()
    {
        // A host read-back is single-precision; a rounding inside the documented host
        // tolerance matches, exactly as the R6.3 live read-back pins it.
        double tolerance = HostOwnedGeometry.GeometryMatchTolerancePt;
        RectD rounded = new(
            _bounds.X + (tolerance / 4),
            _bounds.Y - (tolerance / 4),
            _bounds.Width - (tolerance / 2),
            _bounds.Height + (tolerance / 2));
        FakeGroupingPort port = new() { Bounds = rounded };

        StagedGroupOutcome outcome = StagedShapeGrouper.Group(Staged(_ids, _bounds), port);

        Assert.True(outcome.Succeeded, "Grouping refused: " + outcome.Refusal);
    }

    [Fact]
    public void A_group_envelope_beyond_the_host_tolerance_refuses_and_restores()
    {
        // The positive bounds-mismatch test: twice the host tolerance is a real
        // disagreement, and the group is ungrouped rather than accepted.
        double beyond = HostOwnedGeometry.GeometryMatchTolerancePt * 2;
        FakeGroupingPort port = new()
        {
            Bounds = new RectD(_bounds.X - beyond, _bounds.Y, _bounds.Width + beyond, _bounds.Height),
        };

        StagedGroupOutcome outcome = StagedShapeGrouper.Group(Staged(_ids, _bounds), port);

        Assert.Equal(StagedGroupRefusal.BoundsMismatch, outcome.Refusal);
        Assert.Equal([ShapeStagingTag.GroupName], port.UngroupedNames);
    }

    [Fact]
    public void A_missing_group_envelope_refuses_and_restores()
    {
        FakeGroupingPort port = new() { OmitBounds = true };

        StagedGroupOutcome outcome = StagedShapeGrouper.Group(Staged(_ids, _bounds), port);

        Assert.Equal(StagedGroupRefusal.BoundsMismatch, outcome.Refusal);
        Assert.Equal([ShapeStagingTag.GroupName], port.UngroupedNames);
    }

    [Fact]
    public void A_sentinel_child_that_the_host_grouped_refuses_and_restores()
    {
        // D1 positive: the host grouped one shape more than the staged set. Membership is
        // equality, so the extra child is a refusal - never a group quietly accepted with
        // a user's drawing inside it.
        string[] expected = [.. _ids.Select(id => ShapeStagingTag.NamePrefix + id)];
        FakeGroupingPort port = new()
        {
            Bounds = _bounds,
            ChildrenOverride = [.. expected, "User note"],
        };

        StagedGroupOutcome outcome = StagedShapeGrouper.Group(Staged(_ids, _bounds), port);

        Assert.Equal(StagedGroupRefusal.MembershipMismatch, outcome.Refusal);
        Assert.Contains("User note", outcome.ChildOrder);
        Assert.Equal([ShapeStagingTag.GroupName], port.UngroupedNames);
    }

    [Fact]
    public void A_missing_child_refuses_membership_and_restores()
    {
        string[] expected = [.. _ids.Select(id => ShapeStagingTag.NamePrefix + id)];
        FakeGroupingPort port = new()
        {
            Bounds = _bounds,
            ChildrenOverride = [.. expected.Take(expected.Length - 1)],
        };

        StagedGroupOutcome outcome = StagedShapeGrouper.Group(Staged(_ids, _bounds), port);

        Assert.Equal(StagedGroupRefusal.MembershipMismatch, outcome.Refusal);
        Assert.Equal([ShapeStagingTag.GroupName], port.UngroupedNames);
    }

    [Fact]
    public void Reordered_children_refuse_and_restore()
    {
        // D2 positive: the host re-ordered the children. The stage refuses rather than
        // silently accepting a different order - the work item's stop condition, made
        // observable in the contract harness.
        string[] expected = [.. _ids.Select(id => ShapeStagingTag.NamePrefix + id)];
        FakeGroupingPort port = new()
        {
            Bounds = _bounds,
            ChildrenOverride = [.. expected.Reverse()],
        };

        StagedGroupOutcome outcome = StagedShapeGrouper.Group(Staged(_ids, _bounds), port);

        Assert.Equal(StagedGroupRefusal.ChildOrderMismatch, outcome.Refusal);
        Assert.Equal([ShapeStagingTag.GroupName], port.UngroupedNames);
    }

    [Fact]
    public void A_wrong_group_identity_refuses_and_restores()
    {
        // D1 positive: the stamp did not land, so the group is not a staging shape and
        // must not be reported as one.
        FakeGroupingPort port = new()
        {
            Bounds = _bounds,
            GroupName = ShapeStagingTag.NamePrefix + "wrong",
        };

        StagedGroupOutcome outcome = StagedShapeGrouper.Group(Staged(_ids, _bounds), port);

        Assert.Equal(StagedGroupRefusal.GroupIdentityMismatch, outcome.Refusal);
        Assert.Equal([ShapeStagingTag.NamePrefix + "wrong"], port.UngroupedNames);
    }

    [Fact]
    public void A_host_refusal_is_reported_without_a_restore()
    {
        FakeGroupingPort port = new() { RefuseGrouping = ShapeWriteRefusal.HostRejected };

        StagedGroupOutcome outcome = StagedShapeGrouper.Group(Staged(_ids, _bounds), port);

        Assert.Equal(StagedGroupRefusal.HostRejected, outcome.Refusal);
        Assert.Equal(ShapeWriteRefusal.HostRejected, outcome.GroupRefusal);
        Assert.Empty(port.UngroupedNames);
    }

    [Fact]
    public void Null_and_unusable_inputs_are_refused()
    {
        FakeGroupingPort port = new() { Bounds = _bounds };

        Assert.Equal(StagedGroupRefusal.NullOutcome, StagedShapeGrouper.Group(null, port).Refusal);
        Assert.Equal(
            StagedGroupRefusal.NullPort,
            StagedShapeGrouper.Group(Staged(_ids, _bounds), null).Refusal);

        // A refused composition has no verified set to group.
        StagingCompositionOutcome refused = new(
            [],
            null,
            StagingCompositionRefusal.ExportRefused,
            null,
            null);
        Assert.Equal(
            StagedGroupRefusal.StagedSetRefused,
            StagedShapeGrouper.Group(refused, port).Refusal);

        // A supported empty composition is the caller's to interpret; Excel cannot group
        // an empty set, so the stage refuses it rather than touching the sheet.
        Assert.Equal(
            StagedGroupRefusal.NothingToGroup,
            StagedShapeGrouper.Group(Staged([], null), port).Refusal);
        Assert.Equal(
            StagedGroupRefusal.NothingToGroup,
            StagedShapeGrouper.Group(Staged(_ids, null), port).Refusal);

        Assert.Empty(port.RequestedNames);
    }

    [Fact]
    public void An_ungroup_failure_does_not_replace_the_refusal()
    {
        // The cleanup rule: the verification refusal is the report, and a throwing restore
        // must not escape as an unhandled exception.
        string[] expected = [.. _ids.Select(id => ShapeStagingTag.NamePrefix + id)];
        FakeGroupingPort port = new()
        {
            Bounds = _bounds,
            ChildrenOverride = [.. expected.Reverse()],
            ThrowOnUngroup = true,
        };

        StagedGroupOutcome outcome = StagedShapeGrouper.Group(Staged(_ids, _bounds), port);

        Assert.Equal(StagedGroupRefusal.ChildOrderMismatch, outcome.Refusal);
        Assert.Equal([ShapeStagingTag.GroupName], port.UngroupedNames);
    }

    private static StagingCompositionOutcome Staged(string[] ids, RectD? bounds) =>
        new(ids, bounds, null, null, null);

    private sealed class FakeGroupingPort : IShapeGroupingPort
    {
        public IReadOnlyList<string> RequestedNames { get; private set; } = [];

        public List<string> UngroupedNames { get; } = [];

        public ShapeWriteRefusal? RefuseGrouping { get; init; }

        public string GroupName { get; init; } = ShapeStagingTag.GroupName;

        public RectD? Bounds { get; init; }

        public bool OmitBounds { get; init; }

        public IReadOnlyList<string>? ChildrenOverride { get; init; }

        public bool ThrowOnUngroup { get; init; }

        public ShapeGroupReadback GroupStaged(IReadOnlyList<string> shapeNames)
        {
            RequestedNames = [.. shapeNames];

            if (RefuseGrouping is { } refusal)
            {
                return new ShapeGroupReadback(null, [], null, refusal);
            }

            IReadOnlyList<string> children = ChildrenOverride ?? shapeNames;
            return new ShapeGroupReadback(GroupName, [.. children], OmitBounds ? null : Bounds, null);
        }

        public ShapeWriteOutcome UngroupStaged(string groupName)
        {
            UngroupedNames.Add(groupName);
            return ThrowOnUngroup
                ? throw new InvalidOperationException("The modelled ungroup failed.")
                : ShapeWriteOutcome.Ok();
        }
    }
}
