using GanttCreator.Core.Scene;

namespace GanttCreator.Office;

/// <summary>Why the staging lifecycle produced no verified operation result (R6.5).</summary>
public enum StagingLifecycleRefusal
{
    /// <summary>No lifecycle request was supplied.</summary>
    NullRequest = 0,

    /// <summary>No application-state scope was supplied.</summary>
    NullStateScope = 1,

    /// <summary>No staging shape write port was supplied.</summary>
    NullStagingPort = 2,

    /// <summary>No grouping port was supplied.</summary>
    NullGroupingPort = 3,

    /// <summary>No exported-bounds resolution was supplied.</summary>
    NullResolution = 4,

    /// <summary>
    /// The composition refused. The composer has already deleted the shapes it staged;
    /// the underlying refusal rides <see cref="StagingLifecycleOutcome.Composition"/>.
    /// </summary>
    CompositionRefused = 5,

    /// <summary>
    /// The grouping stage refused. The grouper has already restored the children; the
    /// underlying refusal rides <see cref="StagingLifecycleOutcome.Group"/>.
    /// </summary>
    GroupingRefused = 6,

    /// <summary>
    /// The copy seam reported a refusal. <see cref="StagingLifecycleOutcome.Copy"/>
    /// carries the seam's result; cleanup still ran.
    /// </summary>
    CopyRefused = 7,

    /// <summary>
    /// Cleanup could not remove every temporary shape. <see
    /// cref="StagingLifecycleOutcome.Residue"/> names what the staging scan still
    /// reported, so the next operation's <c>StagingAlreadyPopulated</c> pre-flight
    /// refuses rather than colliding. Zero residue is absolute: this refusal is never
    /// reported as a success.
    /// </summary>
    CleanupResidue = 8,
}

/// <summary>
/// The result carrier the R6.6 copy seam returns. R6.5 defines no clipboard behaviour; the
/// verified copy (<see cref="StagedShapeCopier"/>) fills it.
/// </summary>
/// <param name="Succeeded">Whether the copy step reports success.</param>
/// <param name="Detail">
/// The diagnostic detail: on success the verified formats and clipboard sequence number, on
/// a refusal the actionable reason. Never schedule content.
/// </param>
/// <param name="Refusal">The stage's typed refusal, or <see langword="null"/> on success.</param>
/// <param name="PortRefusal">
/// The copy port's own refusal when the stage refused at the copy step, or
/// <see langword="null"/>. Symmetric with <see cref="StagedGroupOutcome.GroupRefusal"/>, so a
/// command maps messages on typed members rather than parsing text.
/// </param>
/// <remarks>
/// The R6.5 <c>Success</c> and <c>Refused(detail)</c> factories keep their signatures, so the
/// lifecycle's own facts are unchanged by R6.6. The measured drawing format, the clipboard
/// sequence number and the paste-materialisation assertion belong to R6.6 (ADR-0066 D2); this
/// type carries a refusal without swallowing it.
/// </remarks>
public sealed record StagingCopyResult(
    bool Succeeded,
    string? Detail,
    StagedCopyRefusal? Refusal = null,
    ShapeWriteRefusal? PortRefusal = null)
{
    /// <summary>Gets the successful, detail-free result.</summary>
    public static StagingCopyResult Success { get; } = new(true, null);

    /// <summary>Creates a refused result.</summary>
    /// <param name="detail">Why the copy step refused.</param>
    /// <returns>The refused result.</returns>
    public static StagingCopyResult Refused(string detail) => new(false, detail);
}

/// <summary>The typed result of one staging lifecycle run (R6.5).</summary>
/// <param name="Composition">
/// The composition stage's outcome, or <see langword="null"/> when the run refused
/// before composing.
/// </param>
/// <param name="Group">
/// The grouping stage's outcome, or <see langword="null"/> when the run stopped before
/// grouping (including a supported empty composition).
/// </param>
/// <param name="Copy">The copy seam's result, or <see langword="null"/> when none ran.</param>
/// <param name="Refusal">The lifecycle refusal, or <see langword="null"/> on success.</param>
/// <param name="Residue">
/// The primitive identifiers whose staged shape the host still reported after cleanup,
/// in the order the staging scan returned them. Always empty on a clean run; a non-empty
/// list forces <see cref="StagingLifecycleRefusal.CleanupResidue"/>.
/// </param>
public sealed record StagingLifecycleOutcome(
    StagingCompositionOutcome? Composition,
    StagedGroupOutcome? Group,
    StagingCopyResult? Copy,
    StagingLifecycleRefusal? Refusal,
    IReadOnlyList<string> Residue)
{
    /// <summary>Gets whether the run completed with every temporary shape removed.</summary>
    public bool Succeeded => Refusal is null;
}

/// <summary>The inputs one staging lifecycle run consumes.</summary>
/// <param name="StateScope">
/// The application-state scope that suppresses screen updating, events and alerts,
/// writes the status-bar text, and captures the selection for restoration (ADR-0020,
/// R4.2). Required.
/// </param>
/// <param name="StagingPort">
/// The staging-scheme shape write port (an <see cref="ExcelShapeWriter"/> constructed
/// with <see cref="ShapeOwnershipScheme.Staging"/>). Its collection must be empty on
/// entry; the composer pre-flights that. Required.
/// </param>
/// <param name="GroupingPort">The R6.4 grouping port. Required.</param>
/// <param name="Resolution">
/// The R6.1 export-bounds resolution whose scene and bounds the composition stages, as
/// produced by <see cref="ExportBoundsResolver.TryResolve"/>. Required.
/// </param>
/// <param name="Copy">
/// The copy seam R6.6 fills. A <see langword="null"/> delegate means the caller runs no
/// copy step (the lifecycle still stages, groups, and cleans up); the real verified copy
/// and its clipboard assertions are deliberately R6.6's row, not this one.
/// </param>
/// <param name="TechnicalRecord">
/// Receives one technical record per cleanup failure (and one summary record when
/// residue remains), so checklist K's "one technical record" is satisfiable without a
/// logging dependency in this class.
/// </param>
public sealed record StagingLifecycleRequest(
    IApplicationStateScope? StateScope,
    IShapeWritePort? StagingPort,
    IShapeGroupingPort? GroupingPort,
    ExportBoundsOutcome? Resolution,
    Func<StagedGroupOutcome, StagingCopyResult>? Copy = null,
    Action<string>? TechnicalRecord = null);

/// <summary>
/// R6.5: owns the editable-export staging operation from birth to removal — compose,
/// group, copy, and a <c>finally</c> cleanup that guarantees zero temporary shapes remain
/// and restores application state, success or any injected failure.
/// </summary>
/// <remarks>
/// <para>
/// <b>The decided design is the simple delete path (ADR-0066 D1).</b> Staging shapes are
/// deleted in <c>finally</c> immediately after the copy seam returns; the short-lived
/// staging-workbook fallback is not required by the measured evidence and is not built.
/// </para>
/// <para>
/// <b>Birth-registration, never a sheet scan (work-item D1).</b> Every temporary shape is
/// registered at creation: the composer's <see cref="StagingCompositionOutcome.StagedPrimitiveIds"/>
/// list is the birth record (R6.2's D3 hook), in scene creation order, and the verified
/// group is registered as the newest birth. Cleanup walks that registry; it never
/// enumerates the sheet to decide what to delete, because an unowned shape is
/// untouchable.
/// </para>
/// <para>
/// <b>Order-deterministic cleanup (D2).</b> The group is deleted first — it was created
/// after every child, so this is reverse creation order — and the children are then
/// attempted newest-first. A host that removes the children with the group reports them
/// absent (<see cref="ShapeWriteRefusal.NotFound"/> is the goal state, not a failure); a
/// host that keeps them removes each child individually. The deletion-order question the
/// work item defers to the Tier-A upgrade is probed by the live gate's success body.
/// </para>
/// <para>
/// <b>Residue is reported, never hidden (D3, checklist K).</b> A cleanup delete that
/// fails, throws, or leaves a shape the staging scan still reports produces
/// <see cref="StagingLifecycleRefusal.CleanupResidue"/> with the identifiers named — an
/// operation that could not reach zero residue is never reported as a success. A residue
/// shape is staging-tagged, so the next operation's <c>StagingAlreadyPopulated</c>
/// pre-flight refuses rather than colliding.
/// </para>
/// <para>
/// <b>The verification read is not the forbidden scan (D1).</b> After the registry-driven
/// deletes, one ownership-proven staging <c>ListOwned()</c> read confirms zero residue.
/// It deletes nothing, and the staging tag is recomputed from the shape name, so the read
/// can structurally never classify an unowned shape as staged. It is never the driver of
/// a delete.
/// </para>
/// <para>
/// <b>State restoration is unconditional (checklist J).</b> The state scope's disposal
/// runs after shape cleanup inside the same <c>finally</c>, so selection and the five
/// application settings are restored even when the body threw. Every delete precedes the
/// selection restore by construction.
/// </para>
/// <para>
/// <b>Nothing is exposed through a command yet.</b> R6.6 fills the copy seam and R6.8
/// owns the click.
/// </para>
/// </remarks>
public static class StagingLifecycle
{
    /// <summary>The status-bar text shown while the staging operation runs.</summary>
    public const string StatusText = "Gantt Creator: preparing editable export";

    /// <summary>Runs the staging lifecycle: compose, group, copy, then clean up.</summary>
    /// <param name="request">The run's inputs; see <see cref="StagingLifecycleRequest"/>.</param>
    /// <returns>The typed outcome, with residue named whenever cleanup was incomplete.</returns>
    /// <remarks>
    /// A copy seam that throws propagates after the <c>finally</c> has cleaned up and
    /// restored state — the exception is the operation's report and is never swallowed.
    /// A seam that returns a refusal yields <see cref="StagingLifecycleRefusal.CopyRefused"/>
    /// with the result carried.
    /// </remarks>
    public static StagingLifecycleOutcome Run(StagingLifecycleRequest? request)
    {
        if (request is null)
        {
            return Refused(StagingLifecycleRefusal.NullRequest, null, null, null, []);
        }

        if (request.StateScope is null)
        {
            return Refused(StagingLifecycleRefusal.NullStateScope, null, null, null, []);
        }

        if (request.StagingPort is null)
        {
            return Refused(StagingLifecycleRefusal.NullStagingPort, null, null, null, []);
        }

        if (request.GroupingPort is null)
        {
            return Refused(StagingLifecycleRefusal.NullGroupingPort, null, null, null, []);
        }

        if (request.Resolution is null)
        {
            return Refused(StagingLifecycleRefusal.NullResolution, null, null, null, []);
        }

        IApplicationStateScope scope = request.StateScope;
        IShapeWritePort port = request.StagingPort;
        List<string> registry = [];
        List<string> residue = [];
        string? groupPrimitiveId = null;
        StagingCompositionOutcome? composition = null;
        StagedGroupOutcome? group = null;
        StagingCopyResult? copy = null;
        StagingLifecycleRefusal? refusal = null;

        try
        {
            scope.SuppressScreenUpdating();
            scope.SuppressEvents();
            scope.SuppressAlerts();
            scope.SetStatusBarText(StatusText);
            scope.CaptureSelection();

            composition = EditablePanelStagingComposer.ComposeFullScene(request.Resolution, port);
            if (!composition.Succeeded)
            {
                refusal = StagingLifecycleRefusal.CompositionRefused;
            }
            else if (composition.StagedPrimitiveIds.Count > 0)
            {
                // D1: the composer's staged list is the birth record, in creation order.
                registry.AddRange(composition.StagedPrimitiveIds);

                group = StagedShapeGrouper.Group(composition, request.GroupingPort);
                if (!group.Succeeded)
                {
                    refusal = StagingLifecycleRefusal.GroupingRefused;
                }
                else
                {
                    // The verified group is the newest birth; its primitive identifier is
                    // colon-free, so the staging scheme maps it to the group's live name.
                    groupPrimitiveId = ShapeStagingTag.GroupPrimitiveId;

                    copy = request.Copy?.Invoke(group);
                    if (copy is not null && !copy.Succeeded)
                    {
                        refusal = StagingLifecycleRefusal.CopyRefused;
                    }
                }
            }

            // A supported empty composition (nothing staged) is a clean run: the
            // finally has an empty registry and the copy seam is not invoked (F7).
        }
        finally
        {
            Cleanup(port, registry, groupPrimitiveId, residue, request.TechnicalRecord);

            // The verification read: whatever the staging scan still reports is residue.
            // The scan is ownership-proven and deletes nothing (D1). A survivor the
            // deletes did not account for is recorded once, here, so one failure point
            // yields one technical record (checklist K).
            foreach (string survivor in port.ListOwned())
            {
                string identifier = Normalize(survivor);
                if (!residue.Contains(identifier, StringComparer.Ordinal))
                {
                    residue.Add(identifier);
                    request.TechnicalRecord?.Invoke(
                        $"Staging cleanup left '{identifier}' after every delete succeeded.");
                }
            }

            // Checklist J: the selection restore and the five-setting restore run after
            // every delete, whatever the body's result.
            scope.Dispose();
        }

        if (residue.Count > 0)
        {
            refusal = StagingLifecycleRefusal.CleanupResidue;
        }

        return new StagingLifecycleOutcome(composition, group, copy, refusal, residue);
    }

    /// <summary>
    /// Deletes the verified group first and then every registered child, newest first.
    /// </summary>
    /// <param name="port">The staging shape write port.</param>
    /// <param name="registry">The registered primitive identifiers, in creation order.</param>
    /// <param name="groupPrimitiveId">The verified group's primitive identifier, or null.</param>
    /// <param name="residue">Receives each identifier whose delete failed.</param>
    /// <param name="technicalRecord">Receives one record per failed delete.</param>
    /// <remarks>
    /// D2's order is reverse creation order: the group was created after every child. The
    /// children are still attempted afterwards because the host's grouping-deletion
    /// behaviour is design intent until the live probe records it; a child the group
    /// already removed answers <see cref="ShapeWriteRefusal.NotFound"/>, which is the
    /// goal state.
    /// </remarks>
    private static void Cleanup(
        IShapeWritePort port,
        List<string> registry,
        string? groupPrimitiveId,
        List<string> residue,
        Action<string>? technicalRecord)
    {
        if (groupPrimitiveId is not null)
        {
            _ = TryDelete(port, groupPrimitiveId, residue, technicalRecord);
        }

        for (int index = registry.Count - 1; index >= 0; index--)
        {
            _ = TryDelete(port, registry[index], residue, technicalRecord);
        }
    }

    /// <summary>
    /// Deletes one registered shape, treating absence as success and recording any
    /// failure as residue rather than letting it replace the operation's own result.
    /// </summary>
    /// <param name="port">The staging shape write port.</param>
    /// <param name="primitiveId">The registered primitive identifier.</param>
    /// <param name="residue">Receives the identifier when the shape survived.</param>
    /// <param name="technicalRecord">Receives the failure record.</param>
    /// <returns><see langword="true"/> when the shape is gone after the call.</returns>
    private static bool TryDelete(
        IShapeWritePort port,
        string primitiveId,
        List<string> residue,
        Action<string>? technicalRecord)
    {
        ShapeWriteOutcome outcome;
        try
        {
            outcome = port.Delete(primitiveId);
        }
        // CA1031: cleanup must never replace the operation's own result with an
        // exception; the residue and its record ARE the report (checklist K).
#pragma warning disable CA1031
        catch (Exception exception)
#pragma warning restore CA1031
        {
            residue.Add(primitiveId);
            technicalRecord?.Invoke(
                $"Staging cleanup could not delete '{primitiveId}': {exception.GetType().Name}.");
            return false;
        }

        // Absent is the goal state: a host that deleted the children with the group
        // answers NotFound, and a second cleanup pass must not report that as residue.
        if (outcome.Succeeded || outcome.Refusal == ShapeWriteRefusal.NotFound)
        {
            return true;
        }

        residue.Add(primitiveId);
        technicalRecord?.Invoke(
            $"Staging cleanup could not delete '{primitiveId}': {outcome.Refusal}.");
        return false;
    }

    /// <summary>
    /// Reduces a host-reported staging name to its primitive identifier, so residue is
    /// always named in the form the registry uses.
    /// </summary>
    /// <param name="name">The name the staging scan reported.</param>
    /// <returns>The primitive identifier.</returns>
    private static string Normalize(string name) =>
        name.StartsWith(ShapeStagingTag.NamePrefix, StringComparison.Ordinal)
            ? name[ShapeStagingTag.NamePrefix.Length..]
            : name;

    /// <summary>Creates a refusal outcome with nothing staged.</summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <param name="composition">The composition outcome, when one exists.</param>
    /// <param name="group">The grouping outcome, when one exists.</param>
    /// <param name="copy">The copy result, when one exists.</param>
    /// <param name="residue">The residue identifiers.</param>
    /// <returns>The refusal outcome.</returns>
    private static StagingLifecycleOutcome Refused(
        StagingLifecycleRefusal refusal,
        StagingCompositionOutcome? composition,
        StagedGroupOutcome? group,
        StagingCopyResult? copy,
        IReadOnlyList<string> residue) =>
        new(composition, group, copy, refusal, residue);
}
