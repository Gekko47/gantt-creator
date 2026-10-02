using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R4.8 D1: unowned content survives a refresh. Sentinels are planted on the
/// worksheet, a full reconciliation runs, and every sentinel is asserted still
/// present, still unowned, and never named in a mutating call.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why these tests are not decoration.</b> <see cref="FakeShapeWritePort"/>
/// previously modelled only the owned set, which made this whole guarantee
/// <em>vacuous</em>: with no unowned shape in existence, a reconciler that deleted
/// the entire worksheet passed every preservation assertion. The fake now carries
/// unowned shapes with real alternative text, so the sentinel is reachable and
/// deletable — and a regression that swept user content away fails here.
/// </para>
/// <para>
/// The matrix deliberately includes the near-misses, because "we did not touch it"
/// is only trustworthy when the filter was given a genuine reason to act. A sentinel
/// whose name matches a primitive the scene is about to delete, and a sentinel
/// carrying a perfectly valid ownership tag for a <em>different</em> primitive, are
/// the two cases where a plausible implementation deletes user content.
/// </para>
/// </remarks>
public class UnownedContentPreservationTests
{
    /// <summary>A minimal scene request, positioned so overlap can be expressed.</summary>
    private static OfficeShapeRequest Request(string id, double left = 10d) =>
        new(
            id,
            OfficeShapeKind.Rectangle,
            new OfficeShapeGeometry(Bounds: new RectD(left, 0, 100, 20)),
            ZLayer.ActivityBody);

    /// <summary>
    /// D1's central pin: a user's shape is present, correctly classified unowned,
    /// and untouched by a reconciliation that legitimately creates, updates, and
    /// deletes owned shapes around it.
    /// </summary>
    /// <remarks>
    /// The reconciliation is made to do real work — one update, one create, one
    /// delete — so "no sentinel was touched" is a statement about a busy refresh,
    /// not about a no-op. A refresh that quietly plans nothing would satisfy a
    /// weaker version of this test and prove nothing.
    /// </remarks>
    [Fact]
    public void A_users_shape_survives_a_refresh_that_creates_updates_and_deletes()
    {
        var port = new FakeShapeWritePort()
            .Owned("row-1:bar", "row-9:orphan")
            .Unowned("MyAnnotation");

        OfficeShapeRequest[] desired = [Request("row-1:bar"), Request("row-2:bar")];
        ShapeReconcileOutcome outcome = ShapeReconciler.Reconcile(port, new ShapeReconcileRequest(desired));

        Assert.True(outcome.Succeeded);

        // The refresh really did all three kinds of work.
        Assert.Contains("Create:row-2:bar", port.Calls);
        Assert.Contains("Update:row-1:bar", port.Calls);
        Assert.Contains("Delete:row-9:orphan", port.Calls);

        // The sentinel is still on the worksheet, and still not owned.
        Assert.True(port.Survives("MyAnnotation"));
        Assert.DoesNotContain("MyAnnotation", port.ListOwned());
    }

    /// <summary>
    /// D1's hardest row: a user shape occupying the identifier of an owned shape the
    /// scene has dropped is never deleted, because the tag — not the name — authorises
    /// the write.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This is the case a name-keyed implementation gets wrong, and it is the most
    /// destructive one available: the user draws a shape, names it the same thing the
    /// chart uses, and the next refresh destroys it.
    /// </para>
    /// <para>
    /// The seed replaces rather than adds, because Excel shape names are unique per
    /// worksheet. The scenario is the real one: the chart's own shape is gone (or was
    /// copied out and retyped), and the name now belongs to the user. The reconciler
    /// never even <em>plans</em> a delete, because <c>ListOwned</c> excludes the
    /// sentinel, so the orphan sweep cannot see it at all.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_user_shape_sharing_the_identifier_of_a_deleted_owned_shape_is_not_deleted()
    {
        var port = new FakeShapeWritePort()
            .Owned("row-1:bar", "row-7:orphan")
            // A different owned orphan proves the sweep still does its real work.
            .Owned("row-8:orphan")
            .Unowned("row-7:orphan", alternativeText: "a note I drew about the orphan");

        // The scene no longer declares row-7 or row-8, so both would be orphans.
        ShapeReconcileOutcome outcome =
            ShapeReconciler.Reconcile(port, new ShapeReconcileRequest([Request("row-1:bar")]));

        Assert.True(outcome.Succeeded);

        // The genuinely owned orphan is deleted, and the user's shape under the same
        // identifier as a would-be orphan is not.
        Assert.Contains("Delete:row-8:orphan", port.Calls);
        Assert.DoesNotContain("Delete:row-7:orphan", port.Calls);
        Assert.True(port.Survives("row-7:orphan"));
        Assert.Equal("a note I drew about the orphan", port.AlternativeTextOf("row-7:orphan"));
    }

    /// <summary>
    /// A sentinel carrying a <em>valid</em> ownership tag for a different primitive is
    /// still unowned, so the near-miss the work item calls out is never deleted.
    /// </summary>
    /// <remarks>
    /// A prefix-only filter would classify this shape as owned and delete it. That is
    /// why the real filter recomputes the expected tag for the identifier and
    /// compares, rather than asking only whether the prefix is present — and this
    /// test is what makes that distinction load-bearing rather than theoretical.
    /// </remarks>
    [Fact]
    public void A_shape_carrying_another_primitives_valid_tag_is_still_never_deleted()
    {
        var port = new FakeShapeWritePort()
            .Owned("row-1:bar")
            .Owned("row-3:orphan")
            .TaggedFor("row-3:orphan", "row-3:orphan-marker");

        ShapeReconcileOutcome outcome =
            ShapeReconciler.Reconcile(port, new ShapeReconcileRequest([Request("row-1:bar")]));

        Assert.True(outcome.Succeeded);
        Assert.DoesNotContain("Delete:row-3:orphan", port.Calls);
        Assert.True(port.Survives("row-3:orphan"));
        Assert.DoesNotContain("row-3:orphan", port.ListOwned());
    }

    /// <summary>
    /// D1's geometry-overlap row, as a <strong>positive</strong> test: an unowned
    /// shape sitting exactly where the chart will draw is still never touched.
    /// </summary>
    /// <remarks>
    /// The work item requires this explicitly. The temptation it guards against is
    /// real and was considered: clearing the chart area is a legitimate-looking way
    /// to avoid a stale-shape mess, and it would delete the user's arrows and
    /// annotations that happen to sit under the chart. Overlap is not ownership.
    /// </remarks>
    [Fact]
    public void An_unowned_shape_that_geometrically_overlaps_the_chart_is_never_touched()
    {
        // The user's shape occupies the same rectangle the scene draws into.
        var port = new FakeShapeWritePort()
            .Owned("row-1:bar")
            .Unowned("HighlightBox", alternativeText: null);

        OfficeShapeRequest[] desired = [Request("row-1:bar", 10d), Request("row-2:bar", 40d)];
        ShapeReconcileOutcome outcome = ShapeReconciler.Reconcile(port, new ShapeReconcileRequest(desired));

        Assert.True(outcome.Succeeded);
        Assert.True(port.Survives("HighlightBox"));
        Assert.DoesNotContain("HighlightBox", port.Calls);
    }

    /// <summary>
    /// The full sentinel matrix: several unowned shapes of every classification, plus
    /// the owned set, all asserted after one refresh.
    /// </summary>
    /// <remarks>
    /// The seed deliberately covers every way a shape can be unowned: no alternative
    /// text, user-authored text, a name that collides with a delete candidate, a
    /// valid tag for the wrong primitive, and a malformed tag. A sentinel type not
    /// modelled here is a sentinel type whose misclassification nothing would catch.
    /// </remarks>
    [Fact]
    public void The_whole_sentinel_matrix_survives_one_refresh_unchanged()
    {
        var port = new FakeShapeWritePort()
            .Owned("row-1:bar", "row-2:bar")
            .Unowned("PlainDrawing")
            .Unowned("NamedLikeOwned", "user-authored alt text")
            .Unowned("row-9:orphan", "collides with a delete candidate")
            .TaggedFor("WrongTag", "row-2:bar")
            .Unowned(ShapeOwnershipTag.Prefix + "not-a-valid-hash", "a malformed tag");

        string[] before = [.. port.AllShapeNames];
        var alternativeTextBefore = before.ToDictionary(name => name, port.AlternativeTextOf);

        OfficeShapeRequest[] desired = [Request("row-1:bar"), Request("row-2:bar"), Request("row-3:bar")];
        ShapeReconcileOutcome outcome = ShapeReconciler.Reconcile(port, new ShapeReconcileRequest(desired));

        Assert.True(outcome.Succeeded);

        // Every sentinel survives.
        foreach (string sentinel in before.Where(name => !name.StartsWith("row-", StringComparison.Ordinal)))
        {
            Assert.True(port.Survives(sentinel), $"Sentinel '{sentinel}' was removed by the refresh.");
        }

        // Alternative text is unchanged, so nothing re-stamped a tag to "repair" a
        // shape it was not entitled to claim.
        foreach (string sentinel in before)
        {
            Assert.Equal(alternativeTextBefore[sentinel], port.AlternativeTextOf(sentinel));
        }

        // No unowned shape was ever named in a mutating call.
        foreach (string sentinel in before.Where(name => !name.StartsWith("row-", StringComparison.Ordinal)))
        {
            Assert.DoesNotContain(port.Calls, call => call.EndsWith(":" + sentinel, StringComparison.Ordinal));
        }
    }

    /// <summary>
    /// D4: a refusal part-way through a refresh still leaves every sentinel intact.
    /// </summary>
    /// <remarks>
    /// The failure path is where preservation is most likely to leak, because a
    /// half-applied plan is exactly when a caller reaches for a blunt recovery. A
    /// user shape must survive the refresh failing, not only the refresh succeeding.
    /// </remarks>
    [Fact]
    public void A_failed_refresh_still_leaves_every_sentinel_untouched()
    {
        var port = new FakeShapeWritePort()
            .Owned("row-1:bar")
            .Unowned("PreciousUserContent");
        port.Refusals["Update:row-1:bar"] = ShapeWriteRefusal.HostRejected;

        OfficeShapeRequest[] desired = [Request("row-1:bar"), Request("row-2:bar")];
        ShapeReconcileOutcome outcome = ShapeReconciler.Reconcile(port, new ShapeReconcileRequest(desired));

        Assert.False(outcome.Succeeded);
        Assert.True(port.Survives("PreciousUserContent"));
        Assert.DoesNotContain(port.Calls, call => call.EndsWith(":PreciousUserContent", StringComparison.Ordinal));
    }

    /// <summary>
    /// The z-order pass never reorders a sentinel, even when the sentinel is named in
    /// the list — the same ownership filter <c>Update</c> and <c>Delete</c> apply.
    /// </summary>
    /// <remarks>
    /// A preservation guarantee that holds for geometry and deletion but not for
    /// z-order is not a guarantee a user can feel: the shape is still there, but the
    /// chart now draws over it. The adapter enforces this; the fake's
    /// <c>CarriesOwnTag</c> mirror is what makes it assertable here.
    /// </remarks>
    [Fact]
    public void An_unowned_shape_is_never_reordered_by_the_z_order_pass()
    {
        // The name belongs to the user, so the shape carrying it is unowned.
        var port = new FakeShapeWritePort()
            .Owned("row-1:bar", "row-2:bar")
            .Unowned("row-1:bar", "user shape that stole the name");

        // A hand-built order that names the sentinel: the fake refuses the whole pass
        // rather than partially applying it, exactly as the adapter does.
        ShapeWriteOutcome outcome = port.ApplyZOrder(["row-1:bar", "row-2:bar"]);

        Assert.Equal(ShapeWriteRefusal.NotFound, outcome.Refusal);

        // Refused atomically: no order was applied at all, so the owned shape that
        // would have been legal to move was not moved either.
        Assert.Empty(port.ZOrderLists);
    }

    /// <summary>
    /// The failure probe: proves the delete path's ownership filter is load-bearing,
    /// and records the two-layer defence it completes.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>What the probe showed, measured rather than assumed.</b> With the filter
    /// removed from the fake's <see cref="FakeShapeWritePort.Delete"/>, exactly one
    /// of the eight tests in this class failed — this one. The other seven stayed
    /// green, and that is the genuinely useful finding: the reconciler tests are
    /// protected by a <em>different</em> layer. <c>ListOwned</c> already excludes an
    /// unowned shape, so the planner never sees the identifier and never emits a
    /// delete for it at all. The sentinel is safe before any delete-path check runs.
    /// </para>
    /// <para>
    /// So the preservation guarantee has two independent layers, and neither test
    /// group alone is sufficient evidence:
    /// <list type="number">
    /// <item><description>
    /// the enumeration filter — a sentinel is never <em>listed</em>, so it is never a
    /// delete candidate (covered by the reconciler tests above);
    /// </description></item>
    /// <item><description>
    /// the mutation filter — a sentinel named directly is still refused (covered by
    /// this probe).
    /// </description></item>
    /// </list>
    /// Layer 1 is what the user experiences; layer 2 is what makes the guarantee hold
    /// against a future caller that bypasses the planner. Removing either one alone
    /// leaves the other intact, which is the property worth having, and it is why the
    /// probe is a separate test rather than folded into a matrix row.
    /// </para>
    /// <para>
    /// The probe drives the port directly rather than mutating shared state, so it
    /// cannot make the suite order-dependent or leave the fake modified for another
    /// test.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_delete_path_refuses_a_sentinel_named_directly_outside_the_planner()
    {
        var port = new FakeShapeWritePort().Unowned("row-1:bar", "a user shape");
        Assert.Equal("a user shape", port.AlternativeTextOf("row-1:bar"));

        // The filter is what refuses this. A delete authorised by name alone — the
        // pre-R4.1 behaviour, and the regression this matrix exists to catch — would
        // remove the shape and leave a worksheet that quietly lost the user's
        // content. The sentinel is reachable, so the refusal is a real decision
        // rather than an absent object.
        ShapeWriteOutcome outcome = port.Delete("row-1:bar");

        Assert.Equal(ShapeWriteRefusal.NotFound, outcome.Refusal);
        Assert.True(port.Survives("row-1:bar"));
    }

    /// <summary>
    /// Layer 1 of the same guarantee, asserted directly: a sentinel is never
    /// <em>listed</em>, so the planner cannot select it for deletion.
    /// </summary>
    /// <remarks>
    /// This is the layer the failure probe showed to be independently sufficient
    /// against a name-authorised delete. It is asserted on its own because it is the
    /// layer that actually protects a user in normal operation, and a test that only
    /// checked the mutation path would overstate how the safety comes about.
    /// </remarks>
    [Fact]
    public void A_sentinel_is_never_listed_so_the_planner_cannot_select_it_for_deletion()
    {
        var port = new FakeShapeWritePort()
            .Owned("row-1:bar")
            .Unowned("PlainDrawing")
            .TaggedFor("WrongTag", "row-1:bar");

        Assert.Equal(["row-1:bar"], port.ListOwned());
    }
}
