namespace GanttCreator.Office;

/// <summary>What reconciliation decided to do with one owned shape.</summary>
/// <remarks>
/// The members are ordered by pipeline stage, not alphabetically: a delete
/// releases a name that a create may reuse, so <see cref="Delete"/> must be
/// planned before <see cref="Create"/> is issued. See
/// <see cref="ShapeReconciler"/> for the ordering rule.
/// </remarks>
public enum ShapeReconcileAction
{
    /// <summary>
    /// The scene declares this identifier and the host already holds an owned
    /// shape for it, so the shape is re-applied in place.
    /// </summary>
    /// <remarks>
    /// Emitted for <em>every</em> matched identifier on every reconciliation, not
    /// only for the ones that changed. That is the deliberate consequence of
    /// R4.7 D5: the planner cannot prove a matched shape already matches, so it
    /// re-applies and the write itself is idempotent. See
    /// <see cref="ShapeReconciler"/> for why a property diff was rejected.
    /// </remarks>
    Update = 0,

    /// <summary>The scene declares this identifier and the host holds no owned shape for it.</summary>
    Create = 1,

    /// <summary>
    /// The host holds an owned shape whose identifier the scene no longer
    /// declares, so it is removed.
    /// </summary>
    /// <remarks>
    /// "Owned" is the whole qualification. An unowned shape is never in
    /// <c>ListOwned()</c>, so it can never appear here and can never be deleted
    /// by a refresh (R4.8 owns the preservation guarantee).
    /// </remarks>
    Delete = 2,
}

/// <summary>One planned reconciliation operation.</summary>
/// <param name="Action">What to do with the shape.</param>
/// <param name="PrimitiveId">
/// The stable scene identifier this operation addresses. Carried explicitly rather
/// than read from <paramref name="Request"/> so that a delete — which has no
/// request, because it removes a shape the scene no longer declares — is still
/// addressable and reportable.
/// </param>
/// <param name="Request">
/// The shape request, or <see langword="null"/> for a
/// <see cref="ShapeReconcileAction.Delete"/>.
/// </param>
public sealed record ShapeReconcileOperation(
    ShapeReconcileAction Action,
    string PrimitiveId,
    OfficeShapeRequest? Request);

/// <summary>The complete set of operations one reconciliation decided on.</summary>
/// <param name="Operations">
/// The planned operations in execution order: every update, then every create,
/// then every delete.
/// </param>
/// <param name="BackToFront">
/// The identifiers in the scene's back-to-front order, for the single
/// <see cref="IShapeWritePort.ApplyZOrder"/> pass.
/// </param>
/// <remarks>
/// This type is a <em>plan</em>: producing it performs no host call. That
/// separation is what makes reconciliation testable without Excel and what lets
/// the idempotence pin be a pure assertion.
/// </remarks>
public sealed record ShapeReconcilePlan(
    IReadOnlyList<ShapeReconcileOperation> Operations,
    IReadOnlyList<string> BackToFront)
{
    /// <summary>Gets the operations of one kind, in execution order.</summary>
    /// <param name="action">The kind to filter by.</param>
    /// <returns>The matching operations.</returns>
    public IReadOnlyList<ShapeReconcileOperation> Of(ShapeReconcileAction action) =>
        [.. Operations.Where(operation => operation.Action == action)];

    /// <summary>Gets whether the plan would mutate nothing at all.</summary>
    /// <remarks>
    /// Always <see langword="false"/> for a well-formed plan, because every
    /// matched identifier produces an update. It is retained because a scene
    /// with no primitives produces a genuinely empty plan, and the property is
    /// how a caller asserts that case without knowing the planner's internals.
    /// </remarks>
    public bool IsEmpty => Operations.Count == 0;
}

/// <summary>One reconciliation: a scene, the owned set, and whether it may run.</summary>
/// <param name="Desired">The scene's translated shape requests.</param>
/// <param name="Owned">
/// The identifiers the host currently holds, or <see langword="null"/> to read
/// them from the port.
/// </param>
/// <param name="HasBlockingErrors">
/// Whether validation produced blocking errors. When <see langword="true"/>,
/// <see cref="ShapeReconciler.Reconcile"/> returns
/// <see cref="ShapeReconcileRefusal.BlockingErrors"/> having issued no shape
/// write at all, so the last valid chart survives a failed refresh (D3).
/// </param>
public sealed record ShapeReconcileRequest(
    IReadOnlyList<OfficeShapeRequest> Desired,
    IReadOnlyList<string>? Owned = null,
    bool HasBlockingErrors = false);

/// <summary>Why a reconciliation stopped without completing.</summary>
public enum ShapeReconcileRefusal
{
    /// <summary>
    /// Validation reported blocking errors, so no shape was written (D3).
    /// </summary>
    BlockingErrors = 0,

    /// <summary>A create, update, or delete call was refused by the host or the guard.</summary>
    /// <remarks>
    /// <see cref="IShapeWritePort"/> already distinguishes why, so this value
    /// carries no reason of its own: the executor reports the port's
    /// <see cref="ShapeWriteRefusal"/> and the identifier that produced it. What
    /// this row adds is <em>stopping</em> — the first refusal ends the
    /// reconciliation, so a half-applied plan is never mistaken for a success.
    /// </remarks>
    OperationRefused = 1,

    /// <summary>The z-order pass was refused after every shape write succeeded.</summary>
    /// <remarks>
    /// Distinct from <see cref="OperationRefused"/> because the shapes are all
    /// correct and only their stacking is wrong — a visibly wrong chart rather
    /// than a missing one, and the reason a caller should not retry blindly.
    /// </remarks>
    ZOrderRefused = 2,
}

/// <summary>One failed step of a reconciliation, with the host's own reason.</summary>
/// <param name="PrimitiveId">The scene identifier the step addressed.</param>
/// <param name="Refusal">The refusal the port reported.</param>
public sealed record ShapeReconcileFailure(string PrimitiveId, ShapeWriteRefusal Refusal);

/// <summary>The typed result of executing one reconciliation.</summary>
/// <param name="Refusal">Why it stopped, or <see langword="null"/> on success.</param>
/// <param name="Plan">The plan that was executed.</param>
/// <param name="Failure">The first failed step, or <see langword="null"/> on success.</param>
/// <param name="CompletedCount">How many shape writes succeeded before stopping.</param>
/// <remarks>
/// On success <see cref="CompletedCount"/> is every planned write. On refusal it
/// is the prefix that landed, which is what makes a partial reconciliation
/// reportable rather than silent.
/// </remarks>
public sealed record ShapeReconcileOutcome(
    ShapeReconcileRefusal? Refusal,
    ShapeReconcilePlan Plan,
    ShapeReconcileFailure? Failure,
    int CompletedCount)
{
    /// <summary>Gets whether every planned step completed.</summary>
    public bool Succeeded => Refusal is null;

    /// <summary>Gets whether at least one shape write was issued.</summary>
    /// <remarks>
    /// The D3 pin reads this on a blocked reconciliation and requires
    /// <see langword="false"/>: a blocked refresh must not touch a single shape.
    /// </remarks>
    public bool MutatedAnyShape => CompletedCount > 0;
}

