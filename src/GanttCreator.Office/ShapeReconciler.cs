using System.Diagnostics;

namespace GanttCreator.Office;

/// <summary>
/// Plans — and, through <see cref="Reconcile"/>, performs — the update, create,
/// and delete operations that bring the host's owned shape set in line with the
/// scene's, keyed by stable scene identifier (R4.7).
/// </summary>
/// <remarks>
/// <para>
/// <b>D5 — the plan is a pure function of (desired, owned).</b> Planning performs
/// no host call, so the idempotence property is provable without Excel.
/// </para>
/// <para>
/// <b>D5 — a full property diff was rejected, and this is the reason.</b> The
/// obvious design reads each owned shape back and updates only what changed,
/// which would let an unchanged scene produce an empty operation set. It is
/// unsound on this write model: <see cref="OfficeShapeRequest.TextColour"/> being
/// <see langword="null"/> means <em>"leave the host font alone"</em>, so after the
/// write a shape holds whatever the host held before and the intent is no longer
/// observable. A read-back diff therefore cannot distinguish "we wrote black"
/// from "the theme supplied black" from "we never wrote", and a tolerance policy
/// loose enough to cope would also hide real changes.
/// </para>
/// <para>
/// The property actually required — and the one this type proves — is
/// <em>fixed-point</em> idempotence: reconciling an already-converged chart
/// produces an identical plan and an identical owned set. That is a pure function
/// of two collections, needs no tolerance, and needs no host read-back. It is
/// also strictly stronger evidence than "this refresh happened to write nothing",
/// which would hold only by accident.
/// </para>
/// <para>
/// <b>D1 — a matched shape is updated in place.</b> R4.4, R4.6, and R4.11 have
/// landed, so <see cref="IShapeWritePort.Update"/> re-applies geometry, style,
/// and text content. An earlier draft of this row specified delete-and-recreate
/// for a style-only change on the grounds that those members were still
/// create-only; that premise expired with R4.6, so update-in-place is both
/// correct and the architecture's stated preference.
/// </para>
/// <para>
/// <b>D2 — update-first.</b> Delete-and-recreate as an atomic fallback exists but
/// is not selected here: the architecture permits it only with recorded benchmark
/// evidence, and that is R4.10's row.
/// </para>
/// <para>
/// <b>D3 — blocking errors short-circuit before any mutation.</b> Reconciliation
/// is not attempted at all when
/// <see cref="ShapeReconcileRequest.HasBlockingErrors"/> is set, so the owned set
/// is untouched.
/// </para>
/// </remarks>
public static partial class ShapeReconciler
{
    /// <summary>
    /// Plans the operations that reconcile the host's owned shapes with the
    /// scene's shape requests. Performs no host call.
    /// </summary>
    /// <param name="desired">The scene's translated shape requests.</param>
    /// <param name="owned">
    /// The identifiers the host currently holds, as returned by
    /// <see cref="IShapeWritePort.ListOwned"/>.
    /// </param>
    /// <param name="unchangedIds">Cached IDs whose desired payload is unchanged.</param>
    /// <param name="scopes">
    /// How much of each changed shape has to be re-applied, or <see langword="null"/> to
    /// re-apply everything. A shape absent from the map gets the full update, so a caller
    /// that cannot classify a change never gets a partial write.
    /// </param>
    /// <returns>The planned operations and the back-to-front z-order list.</returns>
    /// <exception cref="ArgumentNullException">
    /// Thrown when <paramref name="desired"/> or <paramref name="owned"/> is
    /// <see langword="null"/>.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// Thrown when <paramref name="desired"/> declares one identifier twice.
    /// </exception>
    /// <remarks>
    /// <para>
    /// Execution order is updates, then creates, then deletes. Updates and creates
    /// follow the scene's back-to-front order so created shapes land in the right
    /// relative order even before <see cref="IShapeWritePort.ApplyZOrder"/> runs.
    /// Deletes go last so a delete could never vacate a name that a create in the
    /// same plan still needs: <see cref="IShapeWritePort.Create"/> refuses
    /// <see cref="ShapeWriteRefusal.AlreadyExists"/>, so freeing-then-filling is
    /// the only order that is safe if one is ever added.
    /// </para>
    /// <para>
    /// A duplicate identifier is a caller defect and throws, mirroring
    /// <c>GanttScene.TryCreate</c>'s duplicate-identifier refusal. It is not
    /// de-duplicated silently: two requests claiming one name makes the plan
    /// ambiguous, and picking either would make the render depend on iteration
    /// order.
    /// </para>
    /// </remarks>
    public static ShapeReconcilePlan Plan(
        IReadOnlyList<OfficeShapeRequest> desired,
        IReadOnlyList<string> owned,
        IReadOnlySet<string>? unchangedIds = null,
        IReadOnlyDictionary<string, ShapeUpdateScope>? scopes = null)
    {
        ArgumentNullException.ThrowIfNull(desired);
        ArgumentNullException.ThrowIfNull(owned);

        // A HashSet answers the membership question; the sorted delete list below
        // is what keeps the plan deterministic regardless of the host's z-order,
        // which both R4.5 and R4.7 mutate.
        HashSet<string> ownedIds = new(owned, StringComparer.Ordinal);
        HashSet<string> declared = new(StringComparer.Ordinal);
        List<OfficeShapeRequest> updates = [];
        List<OfficeShapeRequest> creates = [];

        foreach (OfficeShapeRequest request in desired)
        {
            if (!declared.Add(request.PrimitiveId))
            {
                throw new ArgumentException(
                    "The scene declared the primitive identifier '"
                        + request.PrimitiveId
                        + "' more than once, so the reconciliation plan would be ambiguous.",
                    nameof(desired)
                );
            }

            if (ownedIds.Contains(request.PrimitiveId))
            {
                if (unchangedIds?.Contains(request.PrimitiveId) != true)
                {
                    updates.Add(request);
                }
            }
            else
            {
                creates.Add(request);
            }
        }

        List<ShapeReconcileOperation> operations = new(updates.Count + creates.Count + owned.Count);
        foreach (OfficeShapeRequest request in updates)
        {
            operations.Add(new ShapeReconcileOperation(ShapeReconcileAction.Update, request.PrimitiveId, request)
            {
                Scope = scopes is not null && scopes.TryGetValue(request.PrimitiveId, out ShapeUpdateScope scope)
                    ? scope
                    : ShapeUpdateScope.All,
            });
        }

        foreach (OfficeShapeRequest request in creates)
        {
            operations.Add(new ShapeReconcileOperation(ShapeReconcileAction.Create, request.PrimitiveId, request));
        }

        List<string> deletes = [];
        foreach (string identifier in owned)
        {
            if (!declared.Contains(identifier))
            {
                deletes.Add(identifier);
            }
        }

        // Ordinal, matching ListOwned's own sort, so an unchanged scene plans an
        // identical delete list on every host. Sorting here rather than trusting
        // the caller's order is what lets the fixed-point pin hold even if a
        // future caller passes a differently-ordered owned list.
        deletes.Sort(StringComparer.Ordinal);
        foreach (string identifier in deletes)
        {
            operations.Add(new ShapeReconcileOperation(ShapeReconcileAction.Delete, identifier, null));
        }

        // The scene's own order: every desired identifier exists afterwards, so
        // this is the complete owned set in back-to-front order.
        List<string> backToFront = [.. desired.Select(request => request.PrimitiveId)];
        return new ShapeReconcilePlan(operations, backToFront);
    }

    /// <summary>
    /// Executes a reconciliation through the shape-write port: updates, creates,
    /// deletes, then the single z-order pass.
    /// </summary>
    /// <param name="writer">The shape-write port.</param>
    /// <param name="request">The reconciliation to perform.</param>
    /// <returns>The typed result, carrying the plan that was executed.</returns>
    /// <exception cref="ArgumentNullException">
    /// Thrown when <paramref name="writer"/> or <paramref name="request"/> is
    /// <see langword="null"/>.
    /// </exception>
    /// <remarks>
    /// <para>
    /// The first refusal stops the reconciliation and is reported with the
    /// identifier that produced it. Stopping is the point: continuing after a
    /// refused update would leave a chart that is partly from the old scene and
    /// partly from the new one, with nothing to say so.
    /// </para>
    /// <para>
    /// The z-order pass runs only when every shape write succeeded, so a
    /// half-applied plan is never stacked as though it were complete.
    /// <see cref="IShapeWritePort.ApplyZOrder"/> preflights the whole list
    /// before issuing any command, so a refused z-order means <em>no</em> shape
    /// moved in z — the chart is then visibly mis-stacked but complete, which is
    /// why <see cref="ShapeReconcileRefusal.ZOrderRefused"/> is reported
    /// separately from a shape-write refusal.
    /// </para>
    /// <para>
    /// This method does not open the application-state scope and does not
    /// surface validation errors; both belong to the Refresh command (R4.9).
    /// It is deliberately the single reconcile step so that row composes
    /// read → validate → build → scope → <c>Reconcile</c> → surface.
    /// </para>
    /// </remarks>
    public static ShapeReconcileOutcome Reconcile(IShapeWritePort writer, ShapeReconcileRequest request)
    {
        ArgumentNullException.ThrowIfNull(writer);
        ArgumentNullException.ThrowIfNull(request);

        // D3. The owned set is read only when the refresh is allowed to run, so a
        // blocked refresh issues no host call at all - not even a read. Reading
        // first and deciding later would be observably different from never
        // asking.
        if (request.HasBlockingErrors)
        {
            return new ShapeReconcileOutcome(ShapeReconcileRefusal.BlockingErrors, _emptyPlan, null, 0);
        }

        // One batch for the whole reconcile: the target shapes collection and the
        // workbook guard are resolved once, and every proxy that a pass obtains is
        // reused by the next (ADR-0054). The verification reads, the plan execution
        // and the z-order pass all run inside it, so a move plans and writes against
        // one resolved collection instead of resolving it per operation.
        //
        // Nothing about correctness moves: every operation still validates its own
        // request, still confirms the ownership tag before writing, and still refuses
        // with the same outcome. The batch only removes the repeat calls that returned
        // the same answer.
        using IDisposable batch = writer.BeginBatch();

        IReadOnlyList<string> owned = request.Owned ?? writer.ListOwned();
        ShapeDeltaSnapshot? previous = request.PreviousSnapshot;
        bool cacheMatchesOwned = previous is not null
            && owned.Count == previous.Signatures.Count
            && owned.All(previous.Signatures.ContainsKey);
        HashSet<string>? unchangedIds = null;
        if (cacheMatchesOwned && previous is not null)
        {
            unchangedIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (OfficeShapeRequest item in request.Desired)
            {
                if (previous.Signatures.TryGetValue(item.PrimitiveId, out string? signature)
                    && string.Equals(signature, ShapeDeltaCache.Signature(item), StringComparison.Ordinal))
                {
                    _ = unchangedIds.Add(item.PrimitiveId);
                }
            }

            // Verify the candidate skips against the LIVE host before trusting them.
            //
            // The signature proves the scene's payload is unchanged, not that the shape on
            // the sheet still matches it: shapes are host objects, and Excel moves and
            // resizes xlMoveAndSize shapes whenever rows are inserted, hidden, shown, or
            // re-heighted - none of which changes the scene. A skipped shape whose live
            // geometry has drifted is a chart that no Refresh repairs, and it is the class
            // of defect the last two live sessions produced. Geometry is verifiable where
            // text and style are not (R4.7 D5 rejected a full property diff because a null
            // colour means "leave the host alone"; four absolute numbers carry no such
            // ambiguity), so the check is sound and costs a handful of property reads per
            // unchanged shape instead of a rewrite.
            if (unchangedIds.Count > 0)
            {
                List<string> drifted = [];
                foreach (OfficeShapeRequest item in request.Desired)
                {
                    if (unchangedIds.Contains(item.PrimitiveId) && !writer.GeometryMatches(item))
                    {
                        drifted.Add(item.PrimitiveId);
                    }
                }

                foreach (string identifier in drifted)
                {
                    _ = unchangedIds.Remove(identifier);
                }
            }
        }

        // Classify what changed, so a shape that only MOVED is moved rather than
        // re-styled and re-lettered (ADR-0053). The part hashes come from the same cache
        // entries that decided "unchanged", and the classification is a pure in-memory
        // comparison: no extra host call decides it.
        //
        // Only geometry-only is scoped. A shape whose appearance changed (a style-key change,
        // a type change, a label change) or whose parts BOTH changed is re-applied in full:
        // writing the style is the honest repair, and a full update is what the verified
        // skip's counterpart has always done.
        Dictionary<string, ShapeUpdateScope>? scopes = null;
        if (unchangedIds is not null && previous is not null && previous.Parts.Count > 0)
        {
            foreach (OfficeShapeRequest item in request.Desired)
            {
                if (unchangedIds.Contains(item.PrimitiveId)
                    || !previous.Parts.TryGetValue(item.PrimitiveId, out ShapePartSignatures? parts))
                {
                    continue;
                }

                bool geometrySame = string.Equals(
                    parts.Geometry,
                    ShapeDeltaCache.GeometrySignature(item),
                    StringComparison.Ordinal);
                bool appearanceSame = string.Equals(
                    parts.Appearance,
                    ShapeDeltaCache.AppearanceSignature(item),
                    StringComparison.Ordinal);

                // A kind change is never a move: the host builds a different shape for a
                // diamond than for a bar, and ApplyGeometry on the wrong kind is not a
                // repair. Re-applying everything is the honest answer.
                bool kindSame = string.Equals(parts.Kind, item.Kind.ToString(), StringComparison.Ordinal);
                if (!kindSame || geometrySame || !appearanceSame)
                {
                    continue;
                }

                // The shape moved — or the host moved it for us. An inserted or deleted
                // worksheet row moves every xlMove shape with its cells, so the live box can
                // already BE the scene's box. Verifying that costs a handful of property
                // reads and removes the write entirely, which is what makes an insert cheap.
                if (writer.GeometryMatches(item))
                {
                    _ = unchangedIds.Add(item.PrimitiveId);
                    continue;
                }

                scopes ??= new Dictionary<string, ShapeUpdateScope>(StringComparer.Ordinal);
                scopes[item.PrimitiveId] = ShapeUpdateScope.Geometry;
            }
        }

        ShapeReconcilePlan plan = Plan(request.Desired, owned, unchangedIds, scopes);

        int completed = 0;
        foreach (ShapeReconcileOperation operation in plan.Operations)
        {
            ShapeWriteOutcome outcome = operation.Action switch
            {
                ShapeReconcileAction.Update when operation.Scope == ShapeUpdateScope.Geometry =>
                    writer.UpdateGeometry(operation.Request!),
                ShapeReconcileAction.Update => writer.Update(operation.Request!),
                ShapeReconcileAction.Create => writer.Create(operation.Request!),
                ShapeReconcileAction.Delete => writer.Delete(operation.PrimitiveId),
                _ => throw new InvalidOperationException(
                    "The reconciliation plan carried the unimplemented action '"
                        + operation.Action
                        + "'. This is a defect in ShapeReconciler.Plan, not a host condition."
                ),
            };

            if (!outcome.Succeeded)
            {
                return new ShapeReconcileOutcome(
                    ShapeReconcileRefusal.OperationRefused,
                    plan,
                    new ShapeReconcileFailure(operation.PrimitiveId, outcome.Refusal!.Value),
                    completed
                );
            }

            completed++;
        }

        // The z-order pass runs only once every shape write has succeeded, so a
        // half-applied plan is never stacked as though it were complete.
        bool orderUnchanged = cacheMatchesOwned && previous is not null
            && SurvivorOrderUnchanged(previous.BackToFront, plan);
        ShapeWriteOutcome zOrder = ShapeWriteOutcome.Ok();
        if (!orderUnchanged)
        {
            long zOrderStarted = Stopwatch.GetTimestamp();
            zOrder = writer.ApplyZOrder(plan.BackToFront);
            double zOrderMilliseconds = Stopwatch.GetElapsedTime(zOrderStarted).TotalMilliseconds;
            return zOrder.Succeeded
                ? new ShapeReconcileOutcome(null, plan, null, completed)
                {
                    ZOrderApplied = true,
                    ZOrderMilliseconds = zOrderMilliseconds,
                }
                : new ShapeReconcileOutcome(ShapeReconcileRefusal.ZOrderRefused, plan, null, completed)
                {
                    ZOrderApplied = true,
                    ZOrderMilliseconds = zOrderMilliseconds,
                };
        }

        return new ShapeReconcileOutcome(null, plan, null, completed);
    }

    /// <summary>
    /// Whether the surviving owned shapes are already in the desired relative order.
    /// </summary>
    /// <param name="previous">The back-to-front order last applied.</param>
    /// <param name="plan">The plan just executed.</param>
    /// <returns><see langword="true"/> when the pass can be skipped.</returns>
    /// <remarks>
    /// <para>
    /// <b>Deletes do not reorder the survivors.</b> Removing a shape leaves every other
    /// shape's relative position untouched, so a delete-only reconcile whose survivors
    /// match the last applied order needs no z-order pass — the pass costs one lookup and
    /// one host command per shape and was running on every such refresh.
    /// </para>
    /// <para>
    /// <b>A CREATE is different and must never take this path.</b> A new shape is added
    /// in front of the existing ones, so a create needs the pass to place it behind the
    /// shapes the scene puts behind it. The caller therefore also requires that the plan
    /// contains no create; this method only answers the relative-order question.
    /// </para>
    /// </remarks>
    private static bool SurvivorOrderUnchanged(
        IReadOnlyList<string> previous,
        ShapeReconcilePlan plan)
    {
        if (plan.Of(ShapeReconcileAction.Create).Count > 0)
        {
            return false;
        }

        // Compare the previous order with the deleted shapes removed against the desired
        // order: removing a shape does not move the survivors, so their relative order is
        // what the pass would have to reproduce.
        HashSet<string> planned = new(plan.BackToFront, StringComparer.Ordinal);
        List<string> previousSurvivors = [.. previous.Where(planned.Contains)];
        return previousSurvivors.SequenceEqual(plan.BackToFront, StringComparer.Ordinal);
    }

    /// <summary>
    /// The plan reported when planning is not permitted, so a blocked
    /// reconciliation still reports a well-formed empty plan rather than a null.
    /// </summary>
    private static readonly ShapeReconcilePlan _emptyPlan = new([], []);
}
