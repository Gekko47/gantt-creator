using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R4.7's reconciliation contract: the plan is a pure function of (desired,
/// owned), an already-converged chart re-plans to an identical plan, a blocked
/// refresh touches nothing, and a refusal stops the reconciliation.
/// </summary>
/// <remarks>
/// <para>
/// These are contract tests, so there is no Excel: <see cref="ShapeReconciler"/>
/// holds no worksheet reference and calls nothing but the port, which is what
/// makes the idempotence proof exact rather than observational. The live
/// double-refresh gate is in the integration project, where it can also check
/// selection integrity.
/// </para>
/// <para>
/// Every validator in <see cref="ShapeReconciler"/> ships with a positive test
/// here that constructs the bad input and asserts the error path fires.
/// </para>
/// </remarks>
public class ShapeReconcilerTests
{
    private static OfficeShapeRequest Request(string id, double left = 10d) =>
        new(
            id,
            OfficeShapeKind.Rectangle,
            new OfficeShapeGeometry(Bounds: new RectD(left, 0, 100, 20)),
            ZLayer.ActivityBody);

    /// <summary>
    /// R4.7 D5's central pin: reconciling a converged chart twice produces the
    /// same plan and the same owned set both times.
    /// </summary>
    /// <remarks>
    /// This is the property the row actually needs, and it is stronger than
    /// "the second refresh wrote nothing": it holds regardless of whether the
    /// plan is empty, and it fails if the planner became order-dependent on the
    /// host's z-order.
    /// </remarks>
    [Fact]
    public void Reconciling_a_converged_chart_twice_plans_identically_and_leaves_the_same_owned_set()
    {
        OfficeShapeRequest[] desired =
        [
            Request("row-1:bar"),
            Request("row-2:bar", 40d),
            Request("row-3:bar", 80d),
        ];
        var port = new FakeShapeWritePort().Owned("row-1:bar", "row-2:bar", "row-3:bar");

        ShapeReconcileOutcome first = ShapeReconciler.Reconcile(port, new ShapeReconcileRequest(desired));
        IReadOnlyList<string> afterFirst = port.ListOwned();
        ShapeReconcileOutcome second = ShapeReconciler.Reconcile(port, new ShapeReconcileRequest(desired));

        Assert.True(first.Succeeded);
        Assert.True(second.Succeeded);
        Assert.Equal(first.Plan.Operations, second.Plan.Operations);
        Assert.Equal(first.Plan.BackToFront, second.Plan.BackToFront);
        Assert.Equal(afterFirst, port.ListOwned());

        // Owned set and z-order are the two things a user can actually see, so
        // they are asserted directly rather than through the plan.
        Assert.Equal(["row-1:bar", "row-2:bar", "row-3:bar"], afterFirst);
        Assert.Equal(2, port.ZOrderLists.Count);
        Assert.Equal(first.Plan.BackToFront, port.ZOrderLists[0]);
        Assert.Equal(first.Plan.BackToFront, port.ZOrderLists[1]);
    }

    /// <summary>
    /// The plan for an already-converged chart is pure: it is identical whether
    /// it is computed before or after execution, and it does not depend on the
    /// order the owned list arrives in.
    /// </summary>
    /// <remarks>
    /// The second half is the part a naive implementation gets wrong. A planner
    /// that iterated <c>owned</c> to build its delete list would emit deletes in
    /// host order, and the host order is mutated by both R4.5 and R4.7 — so the
    /// fixed-point pin would pass on one machine and fail on another.
    /// </remarks>
    [Fact]
    public void The_plan_does_not_depend_on_the_order_the_owned_set_arrives_in()
    {
        OfficeShapeRequest[] desired = [Request("row-2:bar")];

        ShapeReconcilePlan ascending = ShapeReconciler.Plan(desired, ["aaa:orphan", "zzz:orphan"]);
        ShapeReconcilePlan descending = ShapeReconciler.Plan(desired, ["zzz:orphan", "aaa:orphan"]);

        Assert.Equal(ascending.Operations, descending.Operations);
        Assert.Equal(
            ["aaa:orphan", "zzz:orphan"],
            ascending.Of(ShapeReconcileAction.Delete).Select(operation => operation.PrimitiveId));
    }

    /// <summary>
    /// The exact operation set for a changed scene: a scene-new identifier is
    /// created, an orphaned-owned identifier is deleted, and the survivors are
    /// updated.
    /// </summary>
    [Fact]
    public void A_changed_scene_creates_its_new_identifiers_deletes_its_orphans_and_updates_the_survivors()
    {
        OfficeShapeRequest[] desired = [Request("row-1:bar"), Request("row-3:bar", 80d)];
        var port = new FakeShapeWritePort().Owned("row-1:bar", "row-2:bar");

        ShapeReconcileOutcome outcome = ShapeReconciler.Reconcile(port, new ShapeReconcileRequest(desired));

        Assert.True(outcome.Succeeded);

        // The set, not just the count: a planner that created the right number of
        // shapes for the wrong identifiers would pass a count assertion.
        Assert.Equal(["Update:row-1:bar", "Create:row-3:bar", "Delete:row-2:bar"], port.ShapeCalls);
        Assert.Equal(["row-1:bar", "row-3:bar"], port.ListOwned());
    }

    /// <summary>Updates run before creates, which run before deletes.</summary>
    /// <remarks>
    /// The order is not cosmetic: <see cref="IShapeWritePort.Create"/> refuses
    /// <see cref="ShapeWriteRefusal.AlreadyExists"/>, so a delete that vacated a
    /// name before a create needed it would be the only order that could support
    /// a future delete-and-recreate. Pinning the order now means changing it
    /// later is a deliberate act.
    /// </remarks>
    [Fact]
    public void Operations_are_ordered_updates_then_creates_then_deletes()
    {
        OfficeShapeRequest[] desired = [Request("new:one"), Request("kept:one")];

        ShapeReconcilePlan plan = ShapeReconciler.Plan(desired, ["kept:one", "gone:one"]);

        Assert.Equal(
            [ShapeReconcileAction.Update, ShapeReconcileAction.Create, ShapeReconcileAction.Delete],
            plan.Operations.Select(operation => operation.Action));
    }

    /// <summary>
    /// A delete carries no request, because it removes a shape the scene no
    /// longer declares — but it must still name the identifier it deletes.
    /// </summary>
    /// <remarks>
    /// Without this the executor would have to recover the identifier from the
    /// operation's position, duplicating the planner's ordering rule in a second
    /// place where the two could silently disagree.
    /// </remarks>
    [Fact]
    public void A_delete_operation_names_the_identifier_it_deletes()
    {
        ShapeReconcilePlan plan = ShapeReconciler.Plan([Request("kept:one")], ["gone:one"]);

        ShapeReconcileOperation delete = Assert.Single(plan.Of(ShapeReconcileAction.Delete));
        Assert.Equal("gone:one", delete.PrimitiveId);
        Assert.Null(delete.Request);
    }

    /// <summary>
    /// D3's positive pin: a blocked refresh issues no shape write and does not
    /// even read the owned set.
    /// </summary>
    /// <remarks>
    /// <c>ListOwnedCallCount</c> is asserted, not just the shape calls, because a
    /// version that read first and decided afterwards would mutate nothing yet
    /// still be a different observable behaviour — and "did nothing at all" is
    /// the property the architecture states.
    /// </remarks>
    [Fact]
    public void A_blocked_refresh_touches_no_shape_and_does_not_read_the_owned_set()
    {
        var port = new FakeShapeWritePort().Owned("row-1:bar");

        ShapeReconcileOutcome outcome = ShapeReconciler.Reconcile(
            port,
            new ShapeReconcileRequest(
                [Request("row-1:bar"), Request("row-2:bar")],
                HasBlockingErrors: true));

        Assert.Equal(ShapeReconcileRefusal.BlockingErrors, outcome.Refusal);
        Assert.False(outcome.MutatedAnyShape);
        Assert.Equal(0, outcome.CompletedCount);
        Assert.Empty(port.ShapeCalls);
        Assert.Equal(0, port.ListOwnedCallCount);
        Assert.Empty(port.ZOrderLists);
    }

    /// <summary>
    /// A refused operation stops the reconciliation and names the identifier,
    /// rather than continuing and leaving a half-applied chart.
    /// </summary>
    [Fact]
    public void A_refused_operation_stops_the_reconciliation_and_reports_the_identifier()
    {
        var port = new FakeShapeWritePort().Owned("row-1:bar", "row-2:bar");
        _ = port.Refusals.TryAdd("Update:row-2:bar", ShapeWriteRefusal.HostRejected);

        ShapeReconcileOutcome outcome = ShapeReconciler.Reconcile(
            port,
            new ShapeReconcileRequest(
                [Request("row-1:bar"), Request("row-2:bar"), Request("row-3:bar")]));

        Assert.Equal(ShapeReconcileRefusal.OperationRefused, outcome.Refusal);
        Assert.Equal("row-2:bar", outcome.Failure?.PrimitiveId);
        Assert.Equal(ShapeWriteRefusal.HostRejected, outcome.Failure?.Refusal);
        Assert.Equal(1, outcome.CompletedCount);

        // The create that followed the refusal must not have run, and the
        // half-applied chart must not have been stacked as if it were complete.
        Assert.DoesNotContain("Create:row-3:bar", port.ShapeCalls);
        Assert.Empty(port.ZOrderLists);
    }

    /// <summary>
    /// A refused z-order is reported separately from a shape-write refusal: the
    /// shapes are all correct and only their stacking is wrong.
    /// </summary>
    [Fact]
    public void A_refused_z_order_is_reported_distinctly_after_every_shape_write_succeeded()
    {
        var port = new FakeShapeWritePort().Owned("row-1:bar");
        _ = port.Refusals.TryAdd("ZOrder", ShapeWriteRefusal.NotFound);

        ShapeReconcileOutcome outcome =
            ShapeReconciler.Reconcile(port, new ShapeReconcileRequest([Request("row-1:bar")]));

        Assert.Equal(ShapeReconcileRefusal.ZOrderRefused, outcome.Refusal);
        Assert.Null(outcome.Failure);
        Assert.Equal(1, outcome.CompletedCount);
    }

    /// <summary>
    /// Positive test for the duplicate-identifier validator: an ambiguous
    /// desired set throws rather than silently picking one request per name.
    /// </summary>
    /// <remarks>
    /// Two requests claiming one identifier means the plan is ambiguous, and
    /// de-duplicating by iteration order would make the rendered chart depend on
    /// the order the scene happened to enumerate.
    /// </remarks>
    [Fact]
    public void A_desired_set_that_declares_one_identifier_twice_is_refused()
    {
        OfficeShapeRequest[] desired = [Request("row-1:bar", 10d), Request("row-1:bar", 99d)];

        ArgumentException error = Assert.Throws<ArgumentException>(
            () => ShapeReconciler.Plan(desired, []));

        Assert.Contains("row-1:bar", error.Message, StringComparison.Ordinal);
    }

    /// <summary>Positive test for the null-argument validators.</summary>
    [Fact]
    public void A_null_desired_owned_or_port_argument_is_refused()
    {
        Assert.Throws<ArgumentNullException>(() => ShapeReconciler.Plan(null!, []));
        Assert.Throws<ArgumentNullException>(() => ShapeReconciler.Plan([], null!));
        Assert.Throws<ArgumentNullException>(
            () => ShapeReconciler.Reconcile(null!, new ShapeReconcileRequest([])));
        Assert.Throws<ArgumentNullException>(
            () => ShapeReconciler.Reconcile(new FakeShapeWritePort(), null!));
    }

    /// <summary>
    /// An empty scene reconciles to an empty plan and deletes every owned shape,
    /// which is the "user deleted all their rows" case.
    /// </summary>
    [Fact]
    public void An_empty_scene_deletes_every_owned_shape_and_plans_nothing_else()
    {
        var port = new FakeShapeWritePort().Owned("row-1:bar", "row-2:bar");

        ShapeReconcileOutcome outcome = ShapeReconciler.Reconcile(port, new ShapeReconcileRequest([]));

        Assert.True(outcome.Succeeded);
        Assert.Equal(["Delete:row-1:bar", "Delete:row-2:bar"], port.ShapeCalls);
        Assert.Empty(port.ListOwned());
    }

    /// <summary>
    /// The first refresh on an empty sheet creates everything and deletes
    /// nothing, with the z-order list carrying the scene's own order.
    /// </summary>
    [Fact]
    public void The_first_refresh_creates_every_request_in_the_scenes_back_to_front_order()
    {
        OfficeShapeRequest[] desired =
        [
            Request("chart:background", 0d),
            Request("row-1:bar", 10d),
            Request("row-1:label", 60d),
        ];

        var port = new FakeShapeWritePort();
        ShapeReconcileOutcome outcome =
            ShapeReconciler.Reconcile(port, new ShapeReconcileRequest(desired));

        Assert.True(outcome.Succeeded);
        Assert.Equal(3, outcome.Plan.Of(ShapeReconcileAction.Create).Count);
        Assert.Empty(outcome.Plan.Of(ShapeReconcileAction.Delete));

        // Back-to-front, so labels land above the bars they name.
        Assert.Equal(
            ["chart:background", "row-1:bar", "row-1:label"],
            Assert.Single(port.ZOrderLists));
    }
}
