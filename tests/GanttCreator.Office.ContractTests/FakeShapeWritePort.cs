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

    /// <summary>
    /// Every shape the host is modelled to hold, owned or not, keyed by name.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Added by R4.8. Before this row the fake modelled <em>only</em> the owned set,
    /// which made the entire preservation guarantee untestable rather than merely
    /// untested: with no representation of an unowned shape, a reconciler that
    /// deleted the entire worksheet — every user shape, every annotation — produced
    /// the same green result as one that preserved them, because
    /// <see cref="Delete"/> only ever had owned identifiers to remove. The
    /// preservation tests were passing <em>vacuously</em>.
    /// </para>
    /// <para>
    /// The point of the dictionary is that a sentinel is a real member of the host
    /// that <see cref="ListOwned"/> must exclude and that <see cref="Delete"/> /
    /// <see cref="Update"/> must refuse. That is the whole contract, and it can
    /// only fail if the sentinel is reachable.
    /// </para>
    /// </remarks>
    private readonly Dictionary<string, FakeShape> _shapes = new(StringComparer.Ordinal);

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
            _shapes[identifier] = FakeShape.ForPrimitive(identifier);
        }

        return this;
    }

    /// <summary>
    /// Seeds an unowned shape — a user's own drawing, annotation, or pasted object —
    /// which a refresh must leave completely alone.
    /// </summary>
    /// <param name="name">The shape's name, which may or may not collide with an identifier.</param>
    /// <param name="alternativeText">
    /// The shape's alternative text. <see langword="null"/> models a shape with none,
    /// which is the common case for user-drawn content.
    /// </param>
    /// <returns>This fake, for chaining.</returns>
    /// <remarks>
    /// <para>
    /// The sentinel is registered in <see cref="_shapes"/> but deliberately NOT in
    /// <c>_owned</c>: that asymmetry is the test. A reconciler must reach it through
    /// <see cref="ListOwned"/> and never find it.
    /// </para>
    /// <para>
    /// <b>A name collision REPLACES rather than adds.</b> Writing this after
    /// <see cref="Owned"/> for the same name models what Excel actually holds: shape
    /// names are unique per worksheet, so there is no such thing as an owned shape
    /// and a user shape side by side under one name. The realistic case is that a
    /// user shape occupies the name the chart once used — the chart's own shape was
    /// deleted, or the user copied one out and retyped its name — and the entry is
    /// now unowned. Seeding both and expecting two shapes would model a host state
    /// Excel cannot produce, and the resulting test would assert fiction.
    /// </para>
    /// </remarks>
    public FakeShapeWritePort Unowned(string name, string? alternativeText = null)
    {
        _shapes[name] = new FakeShape(name, alternativeText, Owned: false);
        _owned.Remove(name);
        return this;
    }

    /// <summary>
    /// Seeds a shape that carries a <em>valid</em> ownership tag, but for some other
    /// primitive — the R4.8 D1 near-miss that a prefix-only filter would misclassify.
    /// </summary>
    /// <param name="name">The shape's name as the host holds it.</param>
    /// <param name="taggedFor">The primitive whose tag it wrongly carries.</param>
    /// <returns>This fake, for chaining.</returns>
    public FakeShapeWritePort TaggedFor(string name, string taggedFor)
    {
        _shapes[name] = new FakeShape(name, ShapeOwnershipTag.ForPrimitiveId(taggedFor), Owned: false);
        _owned.Remove(name);
        return this;
    }

    /// <summary>
    /// Every shape the host is modelled to hold, owned or not, in insertion order —
    /// the state a preservation assertion compares against.
    /// </summary>
    public IReadOnlyList<string> AllShapeNames => [.. _shapes.Keys];

    /// <summary>
    /// The alternative text of a shape still on the worksheet, or <see langword="null"/>
    /// when no shape by that name remains.
    /// </summary>
    /// <param name="name">The shape's name.</param>
    /// <returns>The modelled alternative text, or <see langword="null"/>.</returns>
    public string? AlternativeTextOf(string name) =>
        _shapes.TryGetValue(name, out FakeShape? shape) ? shape.AlternativeText : null;

    /// <summary>
    /// Whether a shape with this name is still on the worksheet at all — the
    /// assertion a deleted sentinel fails.
    /// </summary>
    /// <param name="name">The shape's name.</param>
    /// <returns><see langword="true"/> when the shape survives.</returns>
    public bool Survives(string name) => _shapes.ContainsKey(name);

    /// <summary>One modelled worksheet shape, owned or unowned.</summary>
    /// <param name="Name">The shape's name.</param>
    /// <param name="AlternativeText">The shape's alternative text.</param>
    /// <param name="Owned">Whether it carries its own valid ownership tag.</param>
    private sealed record FakeShape(string Name, string? AlternativeText, bool Owned)
    {
        /// <summary>An owned shape stamped exactly as <c>ApplyOwnership</c> would.</summary>
        public static FakeShape ForPrimitive(string identifier) =>
            new(identifier, ShapeOwnershipTag.ForPrimitiveId(identifier), Owned: true);
    }

    public ShapeWriteOutcome Create(OfficeShapeRequest request)
    {
        Calls.Add("Create:" + request.PrimitiveId);
        if (RefusalFor("Create", request.PrimitiveId) is { } refusal)
        {
            return ShapeWriteOutcome.Refused(refusal);
        }

        if (_shapes.ContainsKey(request.PrimitiveId))
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.AlreadyExists);
        }

        // A newly created shape is owned: the real adapter stamps Name and the
        // alternative-text tag, and that stamp is what makes the shape visible to
        // the next refresh's ListOwned(). Without it the shape would sit on the
        // sheet forever, indistinguishable from a user's own.
        _shapes[request.PrimitiveId] = FakeShape.ForPrimitive(request.PrimitiveId);
        _owned.Add(request.PrimitiveId);

        return ShapeWriteOutcome.Ok();
    }

    public ShapeWriteOutcome Update(OfficeShapeRequest request)
    {
        Calls.Add("Update:" + request.PrimitiveId);
        if (RefusalFor("Update", request.PrimitiveId) is { } refusal)
        {
            return ShapeWriteOutcome.Refused(refusal);
        }

        // The real port refuses an update for an identifier it does not hold as an
        // OWNED shape (NotFound), and reconciliation must never be the reason that
        // happens. The check is the ownership filter, not name membership: a user
        // shape that happens to share an identifier is refused exactly as the
        // adapter refuses it (R4.8 D1).
        return CarriesOwnTag(request.PrimitiveId)
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

        // The load-bearing R4.8 assertion. Deleting is authorised by the ownership
        // tag, not by the name, so a user's shape that happens to share an
        // identifier is refused NotFound and never removed. A fake that keyed only
        // on _owned would let a reconciler delete user content and stay green.
        if (!CarriesOwnTag(primitiveId))
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
        }

        _owned.Remove(primitiveId);
        _shapes.Remove(primitiveId);
        return ShapeWriteOutcome.Ok();
    }

    /// <summary>
    /// Whether the shape the host holds under this name carries the ownership tag
    /// for that same identifier — the fake's model of the adapter's
    /// <c>CarriesOwnershipTagFor</c>.
    /// </summary>
    /// <param name="primitiveId">The identifier the caller is acting on.</param>
    /// <returns><see langword="true"/> only on an exact tag match.</returns>
    private bool CarriesOwnTag(string primitiveId) =>
        _shapes.TryGetValue(primitiveId, out FakeShape? shape)
        && shape.Owned
        && string.Equals(
            shape.AlternativeText,
            ShapeOwnershipTag.ForPrimitiveId(primitiveId),
            StringComparison.Ordinal);

    public IReadOnlyList<string> ListOwned()
    {
        ListOwnedCallCount++;
        Calls.Add("ListOwned");
        return [.. _owned];
    }

    public ShapeWriteOutcome ApplyZOrder(IReadOnlyList<string> backToFront)
    {
        Calls.Add("ZOrder");
        if (backToFront.Count == 0)
        {
            return ShapeWriteOutcome.Ok();
        }

        // R4.8: the z-order pass is a mutation path like any other, so it is
        // authorised by the ownership tag and not by the name. The real adapter
        // proves ownership for every named shape in a preflight and refuses the whole
        // pass rather than applying a partial order; without the same check here, a
        // sentinel could be reordered by a refresh and this fake would report success
        // — a preservation guarantee that silently did not hold on one path.
        foreach (var primitiveId in backToFront)
        {
            if (!CarriesOwnTag(primitiveId))
            {
                return ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
            }
        }

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
