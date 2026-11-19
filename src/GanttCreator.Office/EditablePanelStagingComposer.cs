using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office;

/// <summary>Why the editable-panel staging composition produced nothing.</summary>
public enum StagingCompositionRefusal
{
    /// <summary>The resolved export bounds were absent.</summary>
    NullResolution = 0,

    /// <summary>No staging write port was supplied.</summary>
    NullPort = 1,

    /// <summary>
    /// The bounds resolution itself refused, so there is no composition to stage. The
    /// underlying export or scene refusal rides
    /// <see cref="StagingCompositionOutcome.Resolution"/>.
    /// </summary>
    ExportRefused = 2,

    /// <summary>
    /// The resolution claimed success but carried no scene, which a real resolution
    /// cannot produce. Refused rather than staging a partial composition.
    /// </summary>
    MissingScene = 3,

    /// <summary>
    /// Staged shapes from an earlier composition are still on the sheet. Staging again
    /// would collide on the deterministic names and could orphan a partial second set,
    /// so the composer refuses until the previous composition is cleaned up.
    /// </summary>
    StagingAlreadyPopulated = 4,

    /// <summary>
    /// The panel-only sub-scene could not be created from the export scene's primitives.
    /// A valid export scene cannot produce this; it is a defect surfaced rather than
    /// staged.
    /// </summary>
    InvalidSubScene = 5,

    /// <summary>
    /// A panel primitive's family has no shape translation. Refused before any create,
    /// so nothing is staged.
    /// </summary>
    UnsupportedPrimitive = 6,

    /// <summary>
    /// The host refused a create. The shapes staged before the refusal were deleted
    /// before this outcome was returned; <see cref="StagingCompositionOutcome.WriteRefusal"/>
    /// carries the port's refusal.
    /// </summary>
    HostRejected = 7,
}

/// <summary>The typed result of staging the editable-export data panel.</summary>
/// <param name="StagedPrimitiveIds">
/// The staged primitive identifiers, in the scene's back-to-front order, or empty when
/// nothing was staged (including after a cleaned-up host refusal).
/// </param>
/// <param name="StagedBounds">
/// The staging origin plus the export composition's size - the envelope every staged
/// shape must lie inside - or <see langword="null"/> when nothing was staged.
/// </param>
/// <param name="Refusal">The composition refusal, or <see langword="null"/> on success.</param>
/// <param name="Resolution">
/// The bounds resolution the composition consumed, so a refusal's deeper cause (an
/// export classification or a scene refusal) is never lost.
/// </param>
/// <param name="WriteRefusal">The port's refusal for <see cref="StagingCompositionRefusal.HostRejected"/>.</param>
public sealed record StagingCompositionOutcome(
    IReadOnlyList<string> StagedPrimitiveIds,
    RectD? StagedBounds,
    StagingCompositionRefusal? Refusal,
    ExportBoundsOutcome? Resolution,
    ShapeWriteRefusal? WriteRefusal)
{
    /// <summary>Gets whether the composition staged what the scene asked for.</summary>
    public bool Succeeded => Refusal is null;
}

/// <summary>
/// R6.2/R6.3: stages the editable-export composition as native shapes at a deterministic
/// staging origin on the live worksheet, from the scene's own primitives.
/// </summary>
/// <remarks>
/// <para>
/// <b>Scene-driven and cell-free (work-item D1).</b> Every staged shape comes from a
/// primitive <see cref="PanelBuilder"/> already emitted into the export scene. The
/// composer never reads or writes a worksheet cell, never adds a sheet, and never
/// re-derives geometry, text, style, or order - it filters, translates, and creates.
/// </para>
/// <para>
/// <b>Exact reproduction (D2).</b> Shape count equals the panel primitive count, bounds
/// equal the primitive bounds translated by the staging delta, texts are verbatim (a
/// blank cell emits no text shape, matching the builder), styles are the theme-resolved
/// values already on the primitives, and creation order is the scene's own order. The
/// collapsed panel is the composed panel: the scene's projection already measured
/// collapsed children at zero height, so the composer never re-filters rows.
/// </para>
/// <para>
/// <b>Staging identity (D4, D6).</b> The write port is expected to be an
/// <see cref="ExcelShapeWriter"/> constructed with
/// <see cref="ShapeOwnershipScheme.Staging"/>: every staged shape carries the staging
/// name prefix and staging tag, so the live reconciler and a mid-staging coalesced
/// refresh are structurally blind to it.
/// </para>
/// <para>
/// <b>Origin (D7).</b> Staging lands beside the chart, to the right of
/// <see cref="GanttScene.ChartBounds"/> - the one direction the table does not grow
/// into - so R6.3 can stage the full scene at the same coordinates and the live
/// inspection can assert that every staged shape lies inside
/// <see cref="StagingCompositionOutcome.StagedBounds"/>.
/// </para>
/// <para>
/// <b>Lifecycle (D3).</b> The composer pre-flights an empty staging collection and, on
/// any refusal after the first create, deletes the shapes it staged before returning.
/// Fault injection beyond these paths and the tag sweep are R6.5's row; nothing in this
/// class is exposed through a command until R6.8.
/// </para>
/// <para>
/// <b>Full scene (R6.3).</b> <see cref="ComposeFullScene"/> stages every primitive the
/// export scene carries - the panel from R6.2 plus every chart entity - through the same
/// one-vector translation, pre-flight, and cleanup. It re-stamps each request out of the
/// ADR-0052 host-owned resize class (D1): a staged shape is temporary and deleted, never
/// reconciled, so a user row or column resize must not stretch it while it exists.
/// </para>
/// <para>
/// <b>Request assembly is not this row.</b> The caller supplies the resolver's outcome;
/// building an <em>export</em> <see cref="SceneBuildRequest"/> from the workbook is
/// R6.3/R6.8's concern, because the live request factory pins the live profile by
/// design.
/// </para>
/// </remarks>
public static class EditablePanelStagingComposer
{
    /// <summary>
    /// The gap between the chart's right edge and the staging origin, in points.
    /// </summary>
    /// <remarks>
    /// Deterministic and stated rather than tuned: any positive margin keeps the staged
    /// panel outside the chart, and 20pt is wide enough that a staged shape's antialiased
    /// stroke cannot visually touch the frame. The R6.3 live inspection is where the
    /// value is judged on the host; changing it moves no scene, no schema, and no tag.
    /// </remarks>
    public const double StagingMarginPt = 20;

    /// <summary>Resolves the staging origin for one chart.</summary>
    /// <param name="chartBounds">The export scene's chart bounds.</param>
    /// <returns>The top-left point the export composition is staged at.</returns>
    public static PointD StagingOriginFor(RectD chartBounds) =>
        new(chartBounds.Right + StagingMarginPt, chartBounds.Top);

    /// <summary>Stages the export scene's panel primitives through the staging write port.</summary>
    /// <param name="resolution">
    /// The bounds resolution whose scene and bounds the composition consumes, as produced
    /// by <see cref="ExportBoundsResolver.TryResolve"/>.
    /// </param>
    /// <param name="stagingPort">
    /// The staging-scheme shape write port. Its collection must be empty on entry.
    /// </param>
    /// <returns>The staged identifiers and envelope, or a typed refusal.</returns>
    public static StagingCompositionOutcome Compose(
        ExportBoundsOutcome? resolution,
        IShapeWritePort? stagingPort)
    {
        if (resolution is null)
        {
            return Refused(StagingCompositionRefusal.NullResolution, null);
        }

        if (stagingPort is null)
        {
            return Refused(StagingCompositionRefusal.NullPort, resolution);
        }

        if (!resolution.Succeeded || resolution.Bounds is null || resolution.Scene is null)
        {
            return Refused(
                resolution.Succeeded
                    ? StagingCompositionRefusal.MissingScene
                    : StagingCompositionRefusal.ExportRefused,
                resolution);
        }

        GanttScene scene = resolution.Scene;
        ExportBounds bounds = resolution.Bounds;

        ScenePrimitive[] panelPrimitives =
            [.. scene.Primitives.Where(PanelPrimitiveRoles.IsPanelPrimitive)];

        // D5 / ADR-0065 D4: with no visible column the scene carries no panel primitive,
        // and the composition is the chart alone. Nothing to stage is a supported empty
        // composition, not a refusal - and nothing is touched, not even the pre-flight.
        if (panelPrimitives.Length == 0)
        {
            return new StagingCompositionOutcome([], null, null, resolution, null);
        }

        // D3: staging names are deterministic, so a previous composition's shapes would
        // collide on the first create and a changed scene could stage a partial second
        // set that nothing owns. Refuse until the caller cleans up.
        if (stagingPort.ListOwned().Count > 0)
        {
            return Refused(StagingCompositionRefusal.StagingAlreadyPopulated, resolution);
        }

        // D9: a panel-only sub-scene, built from the export scene's own primitives with
        // the parent scene's bounds, so the renderer's translation and ordering rules are
        // reused unchanged. TryCreate re-sorts a subset of a total order, which preserves
        // the parent scene's relative order (unique primitive ids).
        SceneCreationOutcome subScene = GanttScene.TryCreate(
            scene.ChartBounds,
            scene.PlotBounds,
            panelPrimitives,
            scene.Warnings);
        if (subScene.Scene is null)
        {
            return Refused(StagingCompositionRefusal.InvalidSubScene, resolution);
        }

        // D7: the delta is stagingOrigin minus the resolution's own placement delta, so
        // R6.1's bounds model stays the single authority for the translation.
        PointD origin = StagingOriginFor(bounds.ChartBounds);
        ChartOriginDelta delta = new(
            origin.X - bounds.PlacementDelta.X,
            origin.Y - bounds.PlacementDelta.Y);

        SceneTranslationOutcome translation = new SceneShapeRenderer(delta).Translate(subScene.Scene);
        if (!translation.Complete)
        {
            // A panel primitive family with no translation: refuse before the first
            // create, so the sheet is untouched.
            return Refused(StagingCompositionRefusal.UnsupportedPrimitive, resolution);
        }

        return Stage(translation.Requests, origin, bounds.Bounds, resolution, stagingPort);
    }

    /// <summary>
    /// R6.3: stages the complete export scene - every panel primitive R6.2 stages plus
    /// every chart entity - at the same deterministic staging origin.
    /// </summary>
    /// <param name="resolution">
    /// The bounds resolution whose scene and bounds the composition consumes, as produced
    /// by <see cref="ExportBoundsResolver.TryResolve"/>.
    /// </param>
    /// <param name="stagingPort">
    /// The staging-scheme shape write port. Its collection must be empty on entry.
    /// </param>
    /// <returns>The staged identifiers and envelope, or a typed refusal.</returns>
    /// <remarks>
    /// <para>
    /// <b>No filter and no sub-scene.</b> The parent scene is already a total order, so
    /// this path translates it directly; the panel's filter and sub-scene exist only
    /// because R6.2's row stages the panel alone. A zero-visible-column composition is
    /// not an empty composition here: the scene still carries the chart, and the chart
    /// still stages beside the export bounds R6.1 resolved.
    /// </para>
    /// <para>
    /// <b>D2 bounds equality.</b> Entity guide section 1 exports no whitespace outside
    /// <see cref="GanttScene.ChartBounds"/>, so the union of these staged envelopes is
    /// expected to equal <see cref="StagingCompositionOutcome.StagedBounds"/> within
    /// <see cref="GeometryMath.Epsilon"/>; the contract tests assert that with
    /// <see cref="StagedEnvelopeCheck"/> rather than inside the composer.
    /// </para>
    /// </remarks>
    public static StagingCompositionOutcome ComposeFullScene(
        ExportBoundsOutcome? resolution,
        IShapeWritePort? stagingPort)
    {
        if (resolution is null)
        {
            return Refused(StagingCompositionRefusal.NullResolution, null);
        }

        if (stagingPort is null)
        {
            return Refused(StagingCompositionRefusal.NullPort, resolution);
        }

        if (!resolution.Succeeded || resolution.Bounds is null || resolution.Scene is null)
        {
            return Refused(
                resolution.Succeeded
                    ? StagingCompositionRefusal.MissingScene
                    : StagingCompositionRefusal.ExportRefused,
                resolution);
        }

        GanttScene scene = resolution.Scene;
        ExportBounds bounds = resolution.Bounds;

        if (stagingPort.ListOwned().Count > 0)
        {
            return Refused(StagingCompositionRefusal.StagingAlreadyPopulated, resolution);
        }

        PointD origin = StagingOriginFor(bounds.ChartBounds);
        ChartOriginDelta delta = new(
            origin.X - bounds.PlacementDelta.X,
            origin.Y - bounds.PlacementDelta.Y);

        SceneTranslationOutcome translation = new SceneShapeRenderer(delta).Translate(scene);
        if (!translation.Complete)
        {
            // A primitive family with no translation: refuse before the first create,
            // so the sheet is untouched.
            return Refused(StagingCompositionRefusal.UnsupportedPrimitive, resolution);
        }

        if (translation.Requests.Count == 0)
        {
            // A valid scene always carries primitives, but an empty request list is a
            // supported empty composition rather than a reason to touch the sheet.
            return new StagingCompositionOutcome([], null, null, resolution, null);
        }

        // D1: staged shapes never enter the ADR-0052 host-owned resize classes. They are
        // temporary and deleted, never reconciled, so the host must not resize them when
        // a user changes a row height or column width while the set exists.
        OfficeShapeRequest[] requests =
        [
            .. translation.Requests.Select(request => request with { HostOwnsExtent = false }),
        ];

        return Stage(requests, origin, bounds.Bounds, resolution, stagingPort);
    }

    /// <summary>
    /// Creates the translated requests through the port, cleaning up on a host refusal.
    /// </summary>
    /// <param name="requests">The translated requests, in scene order.</param>
    /// <param name="origin">The staging origin.</param>
    /// <param name="exportBounds">The zero-origin export bounds.</param>
    /// <param name="resolution">The resolution being staged.</param>
    /// <param name="port">The staging write port.</param>
    /// <returns>The staged identifiers and envelope, or a typed host refusal.</returns>
    private static StagingCompositionOutcome Stage(
        IReadOnlyList<OfficeShapeRequest> requests,
        PointD origin,
        RectD exportBounds,
        ExportBoundsOutcome resolution,
        IShapeWritePort port)
    {
        List<string> staged = [];
        using IDisposable batch = port.BeginBatch();
        foreach (OfficeShapeRequest request in requests)
        {
            ShapeWriteOutcome write;
            try
            {
                write = port.Create(request);
            }
#pragma warning disable CA1031
            catch (Exception)
#pragma warning restore CA1031
            {
                // R6.5: an exception is the operation's report, but shapes already staged
                // must not outlive it. The birth list is the only registry this class owns,
                // so cleanup runs here rather than in a caller that cannot see it. The
                // exception is rethrown unchanged after the reverse-order cleanup.
                CleanupStaged(port, staged);
                throw;
            }

            if (!write.Succeeded)
            {
                CleanupStaged(port, staged);
                return new StagingCompositionOutcome(
                    [],
                    null,
                    StagingCompositionRefusal.HostRejected,
                    resolution,
                    write.Refusal);
            }

            staged.Add(request.PrimitiveId);
        }

        RectD stagedBounds = new(origin.X, origin.Y, exportBounds.Width, exportBounds.Height);
        return new StagingCompositionOutcome(staged, stagedBounds, null, resolution, null);
    }

    /// <summary>
    /// Deletes the shapes staged so far, newest first, so a refused composition leaves no
    /// temporary shape behind (D3).
    /// </summary>
    /// <param name="stagingPort">The staging write port.</param>
    /// <param name="staged">The identifiers created before the failure, in creation order.</param>
    /// <remarks>
    /// Best effort, on the same reasoning as the writer's unowned-shape discard: the
    /// caller is already being told about the original refusal, and a cleanup failure
    /// that replaced it with an exception would hide the cause. A shape that survives a
    /// failed cleanup is still staging-tagged, so R6.5's sweep can reach it.
    /// </remarks>
    private static void CleanupStaged(IShapeWritePort stagingPort, List<string> staged)
    {
        for (int index = staged.Count - 1; index >= 0; index--)
        {
            // CA1031: a cleanup delete must never replace the refusal being reported.
#pragma warning disable CA1031
            try
            {
                _ = stagingPort.Delete(staged[index]);
            }
            catch (Exception)
            {
                // Intentionally empty: the original refusal is the report.
            }
#pragma warning restore CA1031
        }
    }

    private static StagingCompositionOutcome Refused(
        StagingCompositionRefusal refusal,
        ExportBoundsOutcome? resolution) =>
        new([], null, refusal, resolution, null);
}
