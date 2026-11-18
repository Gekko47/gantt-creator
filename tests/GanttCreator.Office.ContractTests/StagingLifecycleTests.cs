using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R6.5: the staging lifecycle leaves zero temporary shapes and restores application state
/// on the success path and at every injected failure point, cleans up from its birth
/// registry (never a sheet scan), and reports residue rather than claiming success.
/// </summary>
/// <remarks>
/// <para>
/// Every row of the implementation plan's failure-injection matrix has a positive test here
/// (AGENTS.md validator rule), over the R4.1 <see cref="FakeShapeWritePort"/> and the
/// grouping fake, with the state scope modelled by <see cref="RecordingScope"/>. The
/// coalesced-refresh row reuses the scheme-writer harness, because the live-vs-staging
/// blindness is a property of the real writer's ownership scheme and cannot be modelled by
/// a scheme-blind fake.
/// </para>
/// <para>
/// The delete-failure rows assert F2's reading of the acceptance criterion: an injected
/// delete failure cannot yield zero residue by construction, so the rows assert the honest
/// outcome — <see cref="StagingLifecycleRefusal.CleanupResidue"/> named, never success, and
/// state restored — rather than a residue-free result the injection itself forbids.
/// </para>
/// </remarks>
public sealed class StagingLifecycleTests
{
    private const double _chartLeft = 500;
    private const double _chartTop = 60;
    private const double _chartWidth = 320;
    private const double _chartHeight = 280;

    private static readonly SceneStyle _fill = new("Fill", fillColour: ColourHex.Parse("#AABBCC"));
    private static readonly SceneStyle _stroke = new("Stroke", strokeColour: ColourHex.Parse("#334455"), outlineWidthPt: 0.5);

    [Fact]
    public void The_success_path_groups_copies_cleans_up_in_order_and_restores_state()
    {
        Fixture fixture = Build();
        FakeShapeWritePort port = new();
        _ = port.Unowned("User note");
        using RecordingScope scope = new();
        RecordingGroupingPort grouping = new(ExpectedGroupBounds(fixture.ChartBounds))
        {
            AfterGroup = () => _ = port.Owned(ShapeStagingTag.GroupPrimitiveId),
        };
        List<string> copied = [];
        List<string> records = [];

        StagingLifecycleOutcome outcome = StagingLifecycle.Run(new(
            scope,
            port,
            grouping,
            fixture.Resolution,
            Copy: group =>
            {
                copied.Add(group.GroupName!);
                return StagingCopyResult.Success;
            },
            TechnicalRecord: records.Add));

        Assert.True(outcome.Succeeded, "Lifecycle refused: " + outcome.Refusal);
        Assert.Empty(outcome.Residue);
        Assert.NotNull(outcome.Composition);
        Assert.NotNull(outcome.Group);
        Assert.NotNull(outcome.Copy);
        Assert.Equal([ShapeStagingTag.GroupName], copied);
        Assert.Empty(records);

        // Checklist J: the five settings and the selection are captured, and the dispose
        // (the selection restore) is the last thing the scope sees.
        Assert.Equal(
            [
                "SuppressScreenUpdating",
                "SuppressEvents",
                "SuppressAlerts",
                "SetStatusBarText:" + StagingLifecycle.StatusText,
                "CaptureSelection",
                "Dispose",
            ],
            scope.Calls);
        Assert.True(scope.Disposed);

        // D2: the group first (it was created last), then every child newest-first.
        string[] ids = [.. outcome.Composition!.StagedPrimitiveIds];
        List<string> expectedDeletes = ["Delete:" + ShapeStagingTag.GroupPrimitiveId];
        expectedDeletes.AddRange(ids.Reverse().Select(id => "Delete:" + id));
        Assert.Equal(expectedDeletes, DeleteCalls(port));

        // Zero residue, and the unowned sentinel was never touched.
        Assert.Empty(port.ListOwned());
        Assert.True(port.Survives("User note"));
    }

    [Fact]
    public void An_empty_composition_runs_clean_without_grouping_or_copying()
    {
        // F7: a supported empty composition (nothing staged) is a clean run - the copy
        // seam is not invoked and the finally has an empty registry.
        ExportBoundsOutcome resolution = Resolution();
        FakeShapeWritePort port = new();
        using RecordingScope scope = new();
        int copyCalls = 0;
        RectD chartBounds = new(_chartLeft, _chartTop, _chartWidth, _chartHeight);

        StagingLifecycleOutcome outcome = StagingLifecycle.Run(new(
            scope,
            port,
            new RecordingGroupingPort(ExpectedGroupBounds(chartBounds)),
            resolution,
            Copy: _ =>
            {
                copyCalls++;
                return StagingCopyResult.Success;
            }));

        Assert.True(outcome.Succeeded, "Lifecycle refused: " + outcome.Refusal);
        Assert.Empty(outcome.Composition!.StagedPrimitiveIds);
        Assert.Null(outcome.Group);
        Assert.Null(outcome.Copy);
        Assert.Equal(0, copyCalls);

        // The two ownership reads (the composer's pre-flight and the verification pass)
        // are the only port calls: nothing was created and nothing was deleted.
        Assert.DoesNotContain(
            port.Calls,
            call => call.StartsWith("Create:", StringComparison.Ordinal)
                || call.StartsWith("Delete:", StringComparison.Ordinal));
        Assert.True(scope.Disposed);
    }

    [Fact]
    public void A_mid_composition_create_refusal_is_cleaned_up_by_the_composer()
    {
        // Row 1: the composer's own reverse-order cleanup runs, so the lifecycle's
        // registry is empty and the sheet is as it was found.
        Fixture fixture = Build();
        FakeShapeWritePort port = new();
        port.Refusals["Create#2"] = ShapeWriteRefusal.HostRejected;
        using RecordingScope scope = new();

        StagingLifecycleOutcome outcome = StagingLifecycle.Run(new(
            scope,
            port,
            new RecordingGroupingPort(ExpectedGroupBounds(fixture.ChartBounds)),
            fixture.Resolution));

        Assert.Equal(StagingLifecycleRefusal.CompositionRefused, outcome.Refusal);
        Assert.Empty(outcome.Residue);
        Assert.Empty(port.ListOwned());
        _ = Assert.Single(DeleteCalls(port));
        Assert.True(scope.Disposed);
    }

    [Fact]
    public void A_throwing_composer_create_cleans_up_through_the_composer_and_rethrows()
    {
        // Row 1c: an exception from Create bypassed the composer's refusal-only cleanup
        // before R6.5; the birth list must clean on the exception path too, and the
        // exception is the report.
        Fixture fixture = Build();
        ThrowingCreatePort port = new(new FakeShapeWritePort(), failAt: 2);
        using RecordingScope scope = new();

        InvalidOperationException thrown = Assert.Throws<InvalidOperationException>(() =>
            StagingLifecycle.Run(new(
                scope,
                port,
                new RecordingGroupingPort(ExpectedGroupBounds(fixture.ChartBounds)),
                fixture.Resolution)));

        Assert.Equal("create failed", thrown.Message);
        Assert.Empty(port.Inner.ListOwned());
        Assert.True(scope.Disposed);
    }

    [Fact]
    public void A_create_refusal_with_a_failed_composer_cleanup_is_reported_as_residue()
    {
        // F1/F3: the composer's cleanup is best-effort and returns no staged ids on a
        // host refusal, so the lifecycle's verification read is the only thing that can
        // catch the shape the composer could not delete.
        Fixture fixture = Build();
        string firstId = StagedIds(fixture)[0];
        FakeShapeWritePort port = new();
        port.Refusals["Create#2"] = ShapeWriteRefusal.HostRejected;
        port.Refusals["Delete:" + firstId] = ShapeWriteRefusal.HostRejected;
        using RecordingScope scope = new();
        List<string> records = [];

        StagingLifecycleOutcome outcome = StagingLifecycle.Run(new(
            scope,
            port,
            new RecordingGroupingPort(ExpectedGroupBounds(fixture.ChartBounds)),
            fixture.Resolution,
            TechnicalRecord: records.Add));

        Assert.Equal(StagingLifecycleRefusal.CleanupResidue, outcome.Refusal);
        Assert.Equal([firstId], outcome.Residue);
        Assert.NotEmpty(records);
        Assert.True(scope.Disposed);
    }

    [Fact]
    public void A_grouping_refusal_cleans_up_the_children_in_reverse_order()
    {
        // Row 2: the grouper refused without producing a group, so no group delete is
        // attempted and the registered children are removed newest-first.
        Fixture fixture = Build();
        FakeShapeWritePort port = new();
        using RecordingScope scope = new();
        string[] ids = StagedIds(fixture);

        StagingLifecycleOutcome outcome = StagingLifecycle.Run(new(
            scope,
            port,
            new RecordingGroupingPort(ExpectedGroupBounds(fixture.ChartBounds))
            {
                RefuseGrouping = ShapeWriteRefusal.HostRejected,
            },
            fixture.Resolution));

        Assert.Equal(StagingLifecycleRefusal.GroupingRefused, outcome.Refusal);
        Assert.Empty(outcome.Residue);
        Assert.Equal(ids.Reverse().Select(id => "Delete:" + id), DeleteCalls(port));
        Assert.Empty(port.ListOwned());
        Assert.True(scope.Disposed);
    }

    [Fact]
    public void A_group_verification_failure_restores_then_cleans_up_the_children()
    {
        // Row 3: the host grouped the wrong children, so the grouper ungrouped; the
        // lifecycle then deletes the children it registered.
        Fixture fixture = Build();
        FakeShapeWritePort port = new();
        using RecordingScope scope = new();
        string[] ids = StagedIds(fixture);
        string[] expectedNames = [.. ids.Select(id => ShapeStagingTag.NamePrefix + id)];
        RecordingGroupingPort grouping = new(ExpectedGroupBounds(fixture.ChartBounds))
        {
            ChildrenOverride = [.. expectedNames.Reverse()],
        };

        StagingLifecycleOutcome outcome = StagingLifecycle.Run(new(
            scope,
            port,
            grouping,
            fixture.Resolution));

        Assert.Equal(StagingLifecycleRefusal.GroupingRefused, outcome.Refusal);
        Assert.Equal(StagedGroupRefusal.ChildOrderMismatch, outcome.Group!.Refusal);
        Assert.Equal([ShapeStagingTag.GroupName], grouping.UngroupedNames);
        Assert.Equal(ids.Reverse().Select(id => "Delete:" + id), DeleteCalls(port));
        Assert.Empty(port.ListOwned());
    }

    [Fact]
    public void A_throwing_copy_seam_still_cleans_up_and_restores_state()
    {
        // Row 4: the exception is the operation's report; it propagates only after the
        // finally has deleted the group and children and disposed the scope.
        Fixture fixture = Build();
        FakeShapeWritePort port = new();
        using RecordingScope scope = new();
        RecordingGroupingPort grouping = new(ExpectedGroupBounds(fixture.ChartBounds))
        {
            AfterGroup = () => _ = port.Owned(ShapeStagingTag.GroupPrimitiveId),
        };

        InvalidOperationException thrown = Assert.Throws<InvalidOperationException>(() =>
            StagingLifecycle.Run(new(
                scope,
                port,
                grouping,
                fixture.Resolution,
                Copy: _ => throw new InvalidOperationException("the copy failed"))));

        Assert.Equal("the copy failed", thrown.Message);
        Assert.Empty(port.ListOwned());
        Assert.Contains("Delete:" + ShapeStagingTag.GroupPrimitiveId, port.Calls);
        Assert.True(scope.Disposed);
        Assert.Equal("Dispose", scope.Calls[^1]);
    }

    [Fact]
    public void A_copy_refusal_is_reported_and_cleaned_up()
    {
        // Row 5: the seam's typed refusal rides the outcome; cleanup still ran.
        Fixture fixture = Build();
        FakeShapeWritePort port = new();
        using RecordingScope scope = new();
        StagingCopyResult refused = StagingCopyResult.Refused("clipboard unavailable");

        StagingLifecycleOutcome outcome = StagingLifecycle.Run(new(
            scope,
            port,
            new RecordingGroupingPort(ExpectedGroupBounds(fixture.ChartBounds)),
            fixture.Resolution,
            Copy: _ => refused));

        Assert.Equal(StagingLifecycleRefusal.CopyRefused, outcome.Refusal);
        Assert.Equal(refused, outcome.Copy);
        Assert.Empty(outcome.Residue);
        Assert.Empty(port.ListOwned());
        Assert.True(scope.Disposed);
    }

    [Fact]
    public void A_failed_cleanup_delete_is_reported_as_residue_never_success()
    {
        // Row 6: the injected delete failure cannot yield zero residue by construction,
        // so the row asserts the honest outcome (F2) - residue named, never success.
        Fixture fixture = Build();
        string[] ids = StagedIds(fixture);
        FakeShapeWritePort port = new();
        port.Refusals["Delete:" + ids[0]] = ShapeWriteRefusal.HostRejected;
        using RecordingScope scope = new();
        List<string> records = [];

        StagingLifecycleOutcome outcome = StagingLifecycle.Run(new(
            scope,
            port,
            new RecordingGroupingPort(ExpectedGroupBounds(fixture.ChartBounds)),
            fixture.Resolution,
            TechnicalRecord: records.Add));

        Assert.Equal(StagingLifecycleRefusal.CleanupResidue, outcome.Refusal);
        Assert.Equal([ids[0]], outcome.Residue);
        _ = Assert.Single(records);
        Assert.True(port.Survives(ids[0]));
        Assert.True(scope.Disposed);
    }

    [Fact]
    public void A_state_scope_member_failure_never_leaves_a_staged_shape()
    {
        // Row 7: the scope is best-effort by contract; if a member throws before staging,
        // the exception propagates and the finally has nothing to clean.
        Fixture fixture = Build();
        FakeShapeWritePort port = new();
        using RecordingScope scope = new() { ThrowOnSuppressAlerts = true };

        _ = Assert.Throws<InvalidOperationException>(() =>
            StagingLifecycle.Run(new(
                scope,
                port,
                new RecordingGroupingPort(ExpectedGroupBounds(fixture.ChartBounds)),
                fixture.Resolution)));

        // Nothing was staged, so the only port call is the finally's verification read.
        Assert.DoesNotContain(
            port.Calls,
            call => call.StartsWith("Create:", StringComparison.Ordinal)
                || call.StartsWith("Delete:", StringComparison.Ordinal));
        Assert.True(scope.Disposed);
    }

    [Fact]
    public void Unowned_sentinels_survive_every_path()
    {
        // Row 10: a user's drawing is never reached by the registry or the verification
        // read, on the success path or a refusal path.
        Fixture fixture = Build();

        FakeShapeWritePort successPort = new();
        _ = successPort.Unowned("User note");
        using RecordingScope successScope = new();
        _ = StagingLifecycle.Run(new(
            successScope,
            successPort,
            new RecordingGroupingPort(ExpectedGroupBounds(fixture.ChartBounds)),
            fixture.Resolution));
        Assert.True(successPort.Survives("User note"));
        Assert.DoesNotContain("Delete:User note", successPort.Calls);

        FakeShapeWritePort refusalPort = new();
        _ = refusalPort.Unowned("User note");
        using RecordingScope refusalScope = new();
        _ = StagingLifecycle.Run(new(
            refusalScope,
            refusalPort,
            new RecordingGroupingPort(ExpectedGroupBounds(fixture.ChartBounds))
            {
                RefuseGrouping = ShapeWriteRefusal.HostRejected,
            },
            fixture.Resolution));
        Assert.True(refusalPort.Survives("User note"));
        Assert.DoesNotContain("Delete:User note", refusalPort.Calls);
    }

    [Fact]
    public void A_coalesced_refresh_during_the_copy_leaves_both_sets_intact()
    {
        // Row 8 (D1 amendment): a refresh fires at the copy failure point, while the
        // staged set exists. The live chart reconciles and every staged shape survives
        // the refresh; the lifecycle's cleanup then removes only the staged set.
        Fixture fixture = Build();
        SharedShapes sheet = new();
        SchemeWriter live = new(sheet, ShapeOwnershipScheme.Live);
        SchemeWriter staging = new(sheet, ShapeOwnershipScheme.Staging);
        Assert.True(live.Create(StagingSchemeFixtures.RectangleRequest("live:bar-1")).Succeeded);
        Assert.True(live.Create(StagingSchemeFixtures.RectangleRequest("live:bar-2")).Succeeded);
        string[] stagedDuringRefresh = [];
        int shapesDuringRefresh = 0;
        using RecordingScope scope = new();

        StagingLifecycleOutcome outcome = StagingLifecycle.Run(new(
            scope,
            staging,
            new RecordingGroupingPort(ExpectedGroupBounds(fixture.ChartBounds)),
            fixture.Resolution,
            Copy: group =>
            {
                // The coalesced pass: the reconciler plans against the LIVE owned set and
                // executes through the live writer, exactly as the production refresh does.
                ShapeReconcilePlan plan = ShapeReconciler.Plan(
                    [
                        StagingSchemeFixtures.RectangleRequest("live:bar-1"),
                        StagingSchemeFixtures.RectangleRequest("live:bar-3"),
                    ],
                    live.ListOwned());
                foreach (ShapeReconcileOperation operation in plan.Operations)
                {
                    ShapeWriteOutcome write = operation.Action switch
                    {
                        ShapeReconcileAction.Update => live.Update(operation.Request!),
                        ShapeReconcileAction.Create => live.Create(operation.Request!),
                        ShapeReconcileAction.Delete => live.Delete(operation.PrimitiveId),
                        _ => throw new InvalidOperationException(
                            "An unhandled reconcile action: " + operation.Action),
                    };
                    Assert.True(write.Succeeded, "Reconcile refused: " + write.Refusal);
                }

                stagedDuringRefresh = [.. group.ChildOrder];
                shapesDuringRefresh = sheet.Shapes.Count;

                // Asserted here, while the staged set exists: every staged shape survived
                // the refresh, and the live chart converged beside it.
                foreach (string name in stagedDuringRefresh)
                {
                    Assert.Contains(
                        sheet.Shapes,
                        shape => string.Equals(shape.Name, name, StringComparison.Ordinal));
                }

                Assert.Equal(["live:bar-1", "live:bar-3"], live.ListOwned());
                return StagingCopyResult.Success;
            }));

        Assert.True(outcome.Succeeded, "Lifecycle refused: " + outcome.Refusal);
        Assert.Equal(2 + stagedDuringRefresh.Length, shapesDuringRefresh);

        // Then the cleanup removed the staged set and nothing else.
        Assert.Empty(staging.ListOwned());
        Assert.Equal(["live:bar-1", "live:bar-3"], live.ListOwned());
        Assert.Equal(2, sheet.Shapes.Count);
    }

    [Fact]
    public void Null_inputs_are_refused_without_touching_anything()
    {
        // Row 11: one positive test per validator.
        Fixture fixture = Build();
        FakeShapeWritePort port = new();
        using RecordingScope scope = new();
        RecordingGroupingPort grouping = new(ExpectedGroupBounds(fixture.ChartBounds));

        Assert.Equal(StagingLifecycleRefusal.NullRequest, StagingLifecycle.Run(null).Refusal);
        Assert.Equal(
            StagingLifecycleRefusal.NullStateScope,
            StagingLifecycle.Run(new(null, port, grouping, fixture.Resolution)).Refusal);
        Assert.Equal(
            StagingLifecycleRefusal.NullStagingPort,
            StagingLifecycle.Run(new(scope, null, grouping, fixture.Resolution)).Refusal);
        Assert.Equal(
            StagingLifecycleRefusal.NullGroupingPort,
            StagingLifecycle.Run(new(scope, port, null, fixture.Resolution)).Refusal);
        Assert.Equal(
            StagingLifecycleRefusal.NullResolution,
            StagingLifecycle.Run(new(scope, port, grouping, null)).Refusal);

        Assert.Empty(port.Calls);
        Assert.Empty(scope.Calls);
        Assert.Empty(grouping.RequestedNames);
    }

    private static IReadOnlyList<string> DeleteCalls(FakeShapeWritePort port) =>
        [.. port.Calls.Where(call => call.StartsWith("Delete:", StringComparison.Ordinal))];

    private static RectD ExpectedGroupBounds(RectD chartBounds) =>
        new(
            chartBounds.Right + EditablePanelStagingComposer.StagingMarginPt,
            chartBounds.Top,
            chartBounds.Width,
            chartBounds.Height);

    /// <summary>Learns the deterministic staged ids from a probe composition.</summary>
    private static string[] StagedIds(Fixture fixture)
    {
        StagingCompositionOutcome probe =
            EditablePanelStagingComposer.ComposeFullScene(fixture.Resolution, new FakeShapeWritePort());
        Assert.True(probe.Succeeded, "Probe composition refused: " + probe.Refusal);
        return [.. probe.StagedPrimitiveIds];
    }

    private sealed record Fixture(ExportBoundsOutcome Resolution, RectD ChartBounds);

    private static Fixture Build()
    {
        ExportBoundsOutcome resolution = Resolution(DefaultPrimitives());
        Assert.NotNull(resolution.Bounds);
        return new Fixture(resolution, resolution.Bounds.ChartBounds);
    }

    private static ExportBoundsOutcome Resolution(params ScenePrimitive[] primitives)
    {
        RectD chart = new(_chartLeft, _chartTop, _chartWidth, _chartHeight);
        RectD plot = new(_chartLeft + 100, _chartTop + 20, 300, 240);
        SceneCreationOutcome created = GanttScene.TryCreate(chart, plot, primitives, []);
        Assert.NotNull(created.Scene);
        return new ExportBoundsOutcome(new ExportBounds(created.Scene.ChartBounds), null, null, created.Scene);
    }

    private static ScenePrimitive[] DefaultPrimitives() =>
    [
        new SceneRect(
            "row-1:bar",
            SceneOwnerId.Chart,
            ZLayer.ActivityBody,
            new RectD(640, 140, 80, 10),
            _fill),
        new SceneLine(
            "chart:frame:left",
            SceneOwnerId.Chart,
            ZLayer.Frame,
            new PointD(_chartLeft, _chartTop),
            new PointD(_chartLeft, _chartTop + _chartHeight),
            _stroke),
    ];

    /// <summary>The R4.2 state scope, recorded rather than driven against Excel.</summary>
    private sealed class RecordingScope : IApplicationStateScope
    {
        public List<string> Calls { get; } = [];

        public bool Disposed { get; private set; }

        public bool ThrowOnSuppressAlerts { get; init; }

        public void SuppressScreenUpdating() => Calls.Add("SuppressScreenUpdating");

        public void SuppressEvents() => Calls.Add("SuppressEvents");

        public void SuppressAlerts()
        {
            Calls.Add("SuppressAlerts");
            if (ThrowOnSuppressAlerts)
            {
                throw new InvalidOperationException("alerts unavailable");
            }
        }

        public void SuppressStatusBar() => Calls.Add("SuppressStatusBar");

        public void SetStatusBarText(string? text) => Calls.Add("SetStatusBarText:" + text);

        public void CaptureSelection() => Calls.Add("CaptureSelection");

        public void Dispose()
        {
            // Idempotent: the lifecycle disposes the scope, and the test's `using`
            // disposes the fake again, which must not add a second recorded call.
            if (Disposed)
            {
                return;
            }

            Calls.Add("Dispose");
            Disposed = true;
        }
    }

    /// <summary>A port that throws at the Nth create and forwards everything else.</summary>
    /// <remarks>
    /// F9: every member is forwarded faithfully - including <c>BeginBatch</c> and
    /// <c>ListOwned</c> - so the behaviour under test is the wrapped port's.
    /// </remarks>
    private sealed class ThrowingCreatePort(IShapeWritePort inner, int failAt) : IShapeWritePort
    {
        private int _creates;

        public IShapeWritePort Inner => inner;

        public ShapeWriteOutcome Create(OfficeShapeRequest request)
        {
            _creates++;
            return _creates == failAt
                ? throw new InvalidOperationException("create failed")
                : inner.Create(request);
        }

        public ShapeWriteOutcome Update(OfficeShapeRequest request) => inner.Update(request);

        public ShapeWriteOutcome UpdateGeometry(OfficeShapeRequest request) => inner.UpdateGeometry(request);

        public ShapeWriteOutcome Delete(string primitiveId) => inner.Delete(primitiveId);

        public IDisposable BeginBatch() => inner.BeginBatch();

        public IReadOnlyList<string> ListOwned() => inner.ListOwned();

        public ShapeWriteOutcome LockOwnedShapes() => inner.LockOwnedShapes();

        public ShapeWriteOutcome UnlockShape(string primitiveId) => inner.UnlockShape(primitiveId);

        public bool GeometryMatches(OfficeShapeRequest request) => inner.GeometryMatches(request);

        public ShapeWriteOutcome ApplyZOrder(IReadOnlyList<string> backToFront) =>
            inner.ApplyZOrder(backToFront);
    }

    /// <summary>The R6.4 grouping fake, with the rollback recorded.</summary>
    private sealed class RecordingGroupingPort(RectD bounds) : IShapeGroupingPort
    {
        public IReadOnlyList<string> RequestedNames { get; private set; } = [];

        public List<string> UngroupedNames { get; } = [];

        public ShapeWriteRefusal? RefuseGrouping { get; init; }

        public IReadOnlyList<string>? ChildrenOverride { get; init; }

        public Action? AfterGroup { get; init; }

        public ShapeGroupReadback GroupStaged(IReadOnlyList<string> shapeNames)
        {
            RequestedNames = [.. shapeNames];

            if (RefuseGrouping is { } refusal)
            {
                return new ShapeGroupReadback(null, [], null, refusal);
            }

            AfterGroup?.Invoke();
            IReadOnlyList<string> children = ChildrenOverride ?? shapeNames;
            return new ShapeGroupReadback(ShapeStagingTag.GroupName, [.. children], bounds, null);
        }

        public ShapeWriteOutcome UngroupStaged(string groupName)
        {
            UngroupedNames.Add(groupName);
            return ShapeWriteOutcome.Ok();
        }
    }
}
