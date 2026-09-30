using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// An in-memory <see cref="IShapeWritePort"/> that records the calls a
/// reconciliation makes, so R4.7's contract is provable without Excel.
/// </summary>
/// <remarks>
/// The fake models the host's behaviour that reconciliation actually depends on:
/// a create makes an identifier owned, a delete removes it, and a second
/// <see cref="ListOwned"/> reports the converged set. That is what lets the
/// idempotence pin be a two-step test rather than a single assertion.
/// </remarks>
internal sealed class FakeShapeWritePort : IShapeWritePort
{
    private readonly SortedSet<string> _owned = new(StringComparer.Ordinal);

    /// <summary>Gets one entry per port call, in the order the executor made it.</summary>
    public List<string> Calls { get; } = [];

    /// <summary>
    /// Gets the shape-mutating calls only, in order — the create, update, and
    /// delete calls, excluding the reads and the z-order pass.
    /// </summary>
    /// <remarks>
    /// Filtered by method name rather than by "not ZOrder", because
    /// <c>ListOwned</c> is a call too and a reconciliation legitimately makes
    /// one before its writes.
    /// </remarks>
    public IReadOnlyList<string> ShapeCalls =>
        [..
            Calls.Where(call =>
                call.StartsWith("Create", StringComparison.Ordinal)
                || call.StartsWith("Update", StringComparison.Ordinal)
                || call.StartsWith("Delete", StringComparison.Ordinal))];

    /// <summary>Gets the z-order list each call received, in order.</summary>
    public List<IReadOnlyList<string>> ZOrderLists { get; } = [];

    /// <summary>How many times <see cref="ListOwned"/> was called.</summary>
    public int ListOwnedCallCount { get; private set; }

    /// <summary>
    /// A refusal to return from the named call, keyed by method name then by the
    /// identifier or the ordinal of the call.
    /// </summary>
    /// <remarks>
    /// Keyed flexibly on purpose: the "stop on the third create" test and the
    /// "refuse this identifier" test are different shapes of failure injection,
    /// and a single string key expresses both without two fake types.
    /// </remarks>
    public Dictionary<string, ShapeWriteRefusal?> Refusals { get; } = new(StringComparer.Ordinal);

    /// <summary>Seeds the owned set, as a previous refresh would have left it.</summary>
    /// <param name="identifiers">The identifiers the host already holds.</param>
    public FakeShapeWritePort Owned(params string[] identifiers)
    {
        foreach (string identifier in identifiers)
        {
            _owned.Add(identifier);
        }

        return this;
    }

    public ShapeWriteOutcome Create(OfficeShapeRequest request)
    {
        Calls.Add("Create:" + request.PrimitiveId);
        if (RefusalFor("Create", request.PrimitiveId) is { } refusal)
        {
            return ShapeWriteOutcome.Refused(refusal);
        }

        if (!_owned.Add(request.PrimitiveId))
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.AlreadyExists);
        }

        return ShapeWriteOutcome.Ok();
    }

    public ShapeWriteOutcome Update(OfficeShapeRequest request)
    {
        Calls.Add("Update:" + request.PrimitiveId);
        if (RefusalFor("Update", request.PrimitiveId) is { } refusal)
        {
            return ShapeWriteOutcome.Refused(refusal);
        }

        // The real port refuses an update for an identifier it does not hold
        // (NotFound), and reconciliation must never be the reason that happens.
        return _owned.Contains(request.PrimitiveId)
            ? ShapeWriteOutcome.Ok()
            : ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
    }

    public ShapeWriteOutcome Delete(string primitiveId)
    {
        Calls.Add("Delete:" + primitiveId);
        if (RefusalFor("Delete", primitiveId) is { } refusal)
        {
            return ShapeWriteOutcome.Refused(refusal);
        }

        return _owned.Remove(primitiveId)
            ? ShapeWriteOutcome.Ok()
            : ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
    }

    public IReadOnlyList<string> ListOwned()
    {
        ListOwnedCallCount++;
        Calls.Add("ListOwned");
        return [.. _owned];
    }

    public ShapeWriteOutcome ApplyZOrder(IReadOnlyList<string> backToFront)
    {
        Calls.Add("ZOrder");
        ZOrderLists.Add(backToFront);
        return Refusals.TryGetValue("ZOrder", out ShapeWriteRefusal? refusal) && refusal is { } reason
            ? ShapeWriteOutcome.Refused(reason)
            : ShapeWriteOutcome.Ok();
    }

    /// <summary>
    /// Resolves a configured refusal, either by identifier
    /// (<c>"Update:row-2:bar"</c>) or by call ordinal
    /// (<c>"Update#3"</c>), or returns <see langword="null"/> for none.
    /// </summary>
    private ShapeWriteRefusal? RefusalFor(string method, string primitiveId)
    {
        if (Refusals.TryGetValue(method + ":" + primitiveId, out ShapeWriteRefusal? byName))
        {
            return byName;
        }

        int ordinal = Calls.Count(call => call.StartsWith(method, StringComparison.Ordinal));
        return Refusals.TryGetValue(method + "#" + ordinal, out ShapeWriteRefusal? byOrdinal)
            ? byOrdinal
            : null;
    }
}
