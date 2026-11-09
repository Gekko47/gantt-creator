using System.Globalization;
using System.Runtime.InteropServices;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="IShapeWritePort"/> over the Excel application object
/// supplied by the host at add-in load.
/// </summary>
/// <param name="application">
/// The Excel application object (for example <c>ExcelDnaUtil.Application</c>), or
/// <see langword="null"/> when the host supplied none (unit tests, non-Excel
/// host). A foreign object fails the interface cast and degrades to the
/// no-active-workbook refusal with no mutation.
/// </param>
/// <param name="protectionGuard">The shared read-only workbook-protection guard.</param>
/// <param name="trace">
/// Optional diagnostic sink for the splitter band's live geometry around a write, or
/// <see langword="null"/> for no trace. Bounded to that one primitive so the cost is
/// a handful of property reads per Refresh; the production sink must not throw.
/// </param>
/// <remarks>
/// <para>
/// ADR-0008 D4: the protection guard is the first read-only check, before any
/// COM mutation. A protected worksheet refuses the whole write.
/// </para>
/// <para>
/// COM ownership: the <c>Application</c>, <c>Workbook</c>, <c>Worksheet</c>,
/// <c>Shapes</c>, and <c>Shape</c> objects reached here are Excel-owned shared
/// roots. This adapter takes no ownership of them, never calls
/// <c>FinalReleaseComObject</c>, and force-releases nothing (the ownership
/// policy of <see cref="ExcelApplicationAdapter"/>). Every proxy is held in a
/// local and used without chained member expressions.
/// </para>
/// <para>
/// The <c>internal virtual</c> accessors isolate the Excel COM parameterised
/// properties (indexers) and the shape-creation calls. Expression trees cannot
/// contain indexed properties (CS0855), so contract tests substitute these
/// seams and every other member through Moq; the real indexer behaviour is
/// exercised by the tagged live-Office integration test.
/// </para>
/// <para>
/// ADR-0019: the adapter writes both ownership members itself. The shape
/// <c>Name</c> is the full primitive identifier and the shape
/// <c>AlternativeText</c> is the bounded <see cref="ShapeOwnershipTag"/> value.
/// A caller can never write an untagged shape through this port, which is what
/// makes R4.8's preservation guarantee structural rather than conventional.
/// </para>
/// </remarks>
public class ExcelShapeWriter(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null,
    Action<string>? trace = null) : IShapeWritePort
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <summary>Optional diagnostic sink for the splitter band's live geometry.</summary>
    private readonly Action<string>? _trace = trace;

    /// <summary>Whether a reconciliation batch is open (ADR-0054).</summary>
    private bool _batchActive;

    /// <summary>
    /// The shapes collection resolved once for the open batch, or <see langword="null"/>
    /// outside a batch or after a refused resolution.
    /// </summary>
    private Excel.Shapes? _batchShapes;

    /// <summary>
    /// Shape proxies resolved during the open batch, keyed by exact name. A name can still
    /// appear or disappear inside a batch (creates and deletes), so this is a memo of
    /// successful lookups, not an authoritative list.
    /// </summary>
    private Dictionary<string, Excel.Shape>? _batchProxies;

    /// <inheritdoc />
    /// <remarks>
    /// Opening a batch hoists the two costs that do not depend on the shape: the worksheet
    /// and table resolution (walks <c>Sheets</c> and <c>ListObjects</c>, then the guard's
    /// per-target check) and the leading protection query. Everything that DOES depend on
    /// the shape -- validation, the ownership tag, the name lookup -- still runs per
    /// operation.
    /// </remarks>
    public IDisposable BeginBatch()
    {
        if (_batchActive)
        {
            throw new InvalidOperationException(
                "A shape-write batch is already open. Batches do not nest, because the inner "
                + "scope would release the resolution the outer one is still using.");
        }

        _batchActive = true;
        _batchShapes = null;
        _batchProxies = new Dictionary<string, Excel.Shape>(StringComparer.Ordinal);
        return new BatchScope(this);
    }

    /// <summary>Ends a batch and releases the proxies it was holding.</summary>
    private sealed class BatchScope(ExcelShapeWriter owner) : IDisposable
    {
        /// <inheritdoc />
        public void Dispose()
        {
            owner._batchActive = false;
            owner._batchShapes = null;
            owner._batchProxies = null;
        }
    }

    /// <summary>Remembers a resolved proxy for the rest of the batch, when one is open.</summary>
    /// <param name="name">The exact shape name.</param>
    /// <param name="shape">The resolved shape, or <see langword="null"/>.</param>
    /// <returns>The shape, so a lookup can return it directly.</returns>
    private Excel.Shape? Remember(string name, Excel.Shape? shape)
    {
        if (shape is not null && _batchProxies is not null)
        {
            _batchProxies[name] = shape;
        }

        return shape;
    }

    /// <summary>
    /// Maps a protection outcome to the refusal a mutating operation returns, or
    /// <see langword="null"/> when the target is writable.
    /// </summary>
    /// <param name="protection">The outcome of the guard consultation.</param>
    /// <returns>The refusal, or <see langword="null"/> when the write may proceed.</returns>
    /// <remarks>
    /// Only the mapping lives here. The consultation itself stays visible in each entry
    /// point, because ADR-0008 D4's architecture gate inspects the entry method's own tokens
    /// and exists precisely because a guard once sat in a helper nothing reached.
    /// </remarks>
    private static ShapeWriteOutcome? RefusalFor(ProtectionGuardOutcome protection) =>
        protection == ProtectionGuardOutcome.NotProtected
            ? null
            : ShapeWriteOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? ShapeWriteRefusal.NoActiveWorkbook
                    : ShapeWriteRefusal.TargetProtected);

    /// <summary>Reads a shape's placement. Test seam over the COM property.</summary>
    /// <param name="shape">The live shape.</param>
    /// <returns>The host's placement for the shape.</returns>
    internal virtual Excel.XlPlacement ReadPlacement(Excel.Shape shape) => shape.Placement;

    /// <summary>Reads a shape's one-based z-order position. Test seam over the COM property.</summary>
    /// <param name="shape">The live shape.</param>
    /// <returns>The host's position, where 1 is the front-most shape.</returns>
    internal virtual int GetZOrderPosition(Excel.Shape shape) => shape.ZOrderPosition;

    /// <summary>
    /// The longest primitive identifier this adapter will hand to the host as a
    /// shape name.
    /// </summary>
    /// <remarks>
    /// KNOWN-LIMITATIONS <strong>L18</strong> records the live observation: Excel
    /// rejects a shape <c>Name</c> beyond roughly 255 characters with
    /// <c>ArgumentException: The specified value is out of range.</c> The name
    /// carries the full <see cref="OfficeShapeRequest.PrimitiveId"/>, and ADR-0017
    /// leaves that identifier unbounded for a shared owner, so the limit is
    /// reachable in principle.
    /// <para>
    /// The observed figure is approximate, so this bound is deliberately the
    /// conservative side of it: a name at or under the cap is still attempted and
    /// a host that disagrees is caught by the discard path in
    /// <see cref="ApplyOwnershipOrDiscard"/>. The alternative error direction -
    /// attempting a name the host will refuse - is the one that leaves debris.
    /// </para>
    /// <para>
    /// L18 assigns the <em>policy</em> for an over-long identifier (a bounded name
    /// with the full identifier in the tag, or a contributor-set ceiling) to R4.7.
    /// This constant is not that policy; it only refuses an identifier this row
    /// cannot render.
    /// </para>
    /// </remarks>
    internal const int MaxShapeNameLength = 255;

    /// <inheritdoc />
    public ShapeWriteOutcome Create(OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);

        if (!Validate(request, out ShapeWriteRefusal? invalid))
        {
            return ShapeWriteOutcome.Refused(invalid!.Value);
        }

        // Refused before any host call. The name is written during ApplyOwnership,
        // which is AFTER AddShape, so without this check an over-long identifier
        // would place a shape on the sheet only to fail on the name write.
        if (request.PrimitiveId.Length > MaxShapeNameLength)
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.HostRejected);
        }

        // ADR-0008 D4's first statement, before any workbook access. ADR-0054: inside a
        // reconcile batch the target was already checked when it was resolved and events are
        // off for the whole pass, so the repeated query can only return the same answer.
        if (!_batchActive && RefusalFor(_protectionGuard.Query()) is { } guardRefusal)
        {
            return guardRefusal;
        }

        Excel.Shapes? shapes = ResolveShapes(out ShapeWriteOutcome? refusal);
        if (shapes is null)
        {
            return refusal!;
        }

        if (FindShapeByName(shapes, request.PrimitiveId) is not null)
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.AlreadyExists);
        }

        Excel.Shape? created = AddShape(shapes, request);
        if (created is null)
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.HostRejected);
        }

        // The host must not own the shape's SIZE -- except for the one class that
        // spans the plot's full height, where the host owning it IS the wanted
        // behaviour (ADR-0052 D6-D7). Excel creates shapes with
        // Placement = xlMoveAndSize, which ties the displayed height to the rows the
        // shape spans: for a bar or a label that resizes the shape behind the renderer's
        // back, so xlMove is written. For a delineator or a plot-spanning band the
        // scene's height is derived from those same rows, so the host resizing it is the
        // chart following its own rows and is kept. Measured 2026-11-03: the scene
        // declared the splitter band at 18pt on every refresh while the live band was
        // four rows tall in one state and a line in another -- which is why the
        // classification is a geometric one (full plot height) and not "all shapes".
        ApplyPlacement(created, request.HostOwnsExtent);

        // R4.6. The style is written on the create path for the same reason the
        // text is: it is generated output the scene owns, and a shape created in
        // the host's default accent colour with the resolved colours never
        // reaching it is the same silent-staleness class the R4.4 content write
        // closed. Order matters — the style is applied BEFORE the ownership stamp,
        // so a shape that cannot be styled is discarded by the same
        // ApplyOwnershipOrDiscard path rather than left half-written.
        //
        // The failure is reported, not swallowed. ApplyStyleOrDiscard returning
        // normally after discarding the shape would leave Create to stamp
        // ownership onto a shape that is no longer on the sheet: a write to a
        // deleted shape that either throws (surfacing as a second, misleading
        // refusal) or succeeds and returns Ok for a shape that does not exist.
        // Neither is acceptable, so a failed style write ends the create here and
        // the ownership stamp is never attempted.
        if (!ApplyStyleOrDiscard(created, request))
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.HostRejected);
        }

        // No per-shape z-order call here. The caller creates the scene's
        // primitives in Translate's back-to-front order and a new shape is
        // inserted in front of the existing ones, so the default insertion order
        // already IS the scene's order. ApplyZOrder is the single authoritative
        // pass, and it establishes the final order independently of whatever
        // order the host happened to be in.
        ShapeWriteOutcome stamped = ApplyOwnershipOrDiscard(created, request.PrimitiveId);
        if (stamped.Succeeded)
        {
            // The new shape is a valid proxy for the rest of the batch: the geometry check
            // and the z-order preflight would otherwise look it up again (ADR-0054).
            _ = Remember(request.PrimitiveId, created);
            TraceBandShape("create", created, request.PrimitiveId, before: null);
        }

        return stamped;
    }

    /// <inheritdoc />
    public ShapeWriteOutcome Update(OfficeShapeRequest request) => UpdateCore(request, applyAppearance: true);

    /// <inheritdoc />
    public ShapeWriteOutcome UpdateGeometry(OfficeShapeRequest request) => UpdateCore(request, applyAppearance: false);

    /// <summary>
    /// The shared update path: authorise, move, and — for a full update — re-apply the
    /// content and the style.
    /// </summary>
    /// <param name="request">The scene's desired shape.</param>
    /// <param name="applyAppearance">
    /// Whether to re-apply content and style as well as geometry. <see langword="false"/>
    /// is the move-only path, used when the payload cache proves nothing but the box
    /// changed (ADR-0053).
    /// </param>
    /// <returns>The typed result.</returns>
    private ShapeWriteOutcome UpdateCore(OfficeShapeRequest request, bool applyAppearance)
    {
        ArgumentNullException.ThrowIfNull(request);

        if (!Validate(request, out ShapeWriteRefusal? invalid))
        {
            return ShapeWriteOutcome.Refused(invalid!.Value);
        }

        // ADR-0008 D4's first statement, before any workbook access. ADR-0054: inside a
        // reconcile batch the target was already checked when it was resolved and events are
        // off for the whole pass, so the repeated query can only return the same answer.
        if (!_batchActive && RefusalFor(_protectionGuard.Query()) is { } guardRefusal)
        {
            return guardRefusal;
        }

        Excel.Shapes? shapes = ResolveShapes(out ShapeWriteOutcome? refusal);
        if (shapes is null)
        {
            return refusal!;
        }

        Excel.Shape? existing = FindShapeByName(shapes, request.PrimitiveId);
        if (existing is null)
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
        }

        // The ownership filter, not the name, authorises the write - exactly as it
        // does for Delete. A name match alone would let a refresh move and rewrite
        // a shape the user drew that happens to share an identifier, which is
        // precisely the R4.8 preservation guarantee. A user shape that keeps the
        // name but edits the alternative text is equally unowned: re-stamping the
        // tag would "repair" it, which is R9.4's job to report, not this row's.
        //
        // NotFound rather than a new refusal: from the caller's point of view
        // "no owned shape with this identifier" is one condition, and R4.7
        // reconciles from ListOwned(), which already excludes such a shape, so it
        // will not reach this path with an unowned identifier in normal operation.
        if (!CarriesOwnershipTagFor(existing, request.PrimitiveId))
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
        }

        string? bandBefore = DescribeBandShape(existing, request.PrimitiveId);
        if (!ApplyGeometry(existing, request.Kind, request.Geometry))
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.InvalidGeometry);
        }

        // Heal a shape created before the placement rule changed (see Create), and keep
        // the host's resize behaviour equal to what this shape's class requires. Every
        // update used to write `Placement` unconditionally -- a host write per shape,
        // most of them writing the value the shape already held. One cheap read now
        // decides, so the write happens only on the first pass after the rule changed
        // (or for a shape a user re-anchored), and an existing workbook still heals itself
        // without a rebuild of its shapes (ADR-0054). The rule is per class: a shape that
        // ADR-0052 put on xlMove, and that the amendment now classifies as host-owned,
        // heals to xlMoveAndSize on its next refresh and vice versa.
        if (ReadPlacement(existing) != RequiredPlacement(request))
        {
            ApplyPlacement(existing, request.HostOwnsExtent);
        }

        TraceBandShape(applyAppearance ? "update" : "move", existing, request.PrimitiveId, bandBefore);

        // The move-only path stops here. Everything below writes generated appearance
        // (text, fill, stroke, hatch, typography) which the caller has already proven
        // unchanged; re-applying it would cost the ~15–18 ms a full update costs for a
        // shape that only changed where it sits. A hand-edited fill is deliberately NOT
        // repaired here — the explicit Refresh does that (ADR-0053).
        if (!applyAppearance)
        {
            return ShapeWriteOutcome.Ok();
        }

        // R4.4: a text box's CONTENT is owned generated data, exactly as its
        // geometry is, so an update re-applies it. Without this, a refresh that
        // changed a label's string, font, or alignment would move the box and
        // leave the stale text on the sheet - the shape would sit in the right
        // place saying the wrong thing, which is the silent-staleness class this
        // project treats as a defect. R4.7 reaches this path on every reconcile
        // that finds an existing shape.
        //
        // This is deliberately NOT the same treatment as the name and tag below.
        // A user never authors a label's text; the scene does, so overwriting it
        // is correct. The name is the reconciliation KEY and the tag is the
        // ownership PROOF, and re-stamping those would silently "repair" a shape
        // a user edited, which is R9.4's job to report, not R4.7's to hide.
        //
        // Both writes are contained, on the same reasoning as the create path's
        // ApplyTextOrDiscard and ApplyStyleOrDiscard: a COM write the host refuses
        // must not escape an update as an unhandled COMException from inside a
        // render command. R4.7's reconciler acts on the typed result, so a
        // HostRejected is what it needs to report - an escaping exception is not.
        //
        // A contract fault still propagates: MapHatchPattern and MapAlignment
        // throw ArgumentOutOfRangeException for a scene value with no host
        // mapping, which is a drift defect between the guide and the adapter, not
        // a host failure, and swallowing it into a refusal would hide exactly the
        // thing the ratchet exists to surface.
        //
        // Containment is NOT a rollback. ApplyGeometry has already run, so a
        // refused text or style write leaves the shape at its new geometry with
        // its old content or colour. That is the deliberate trade: the shape stays
        // on the sheet under its ownership tag, so the next reconcile reaches it
        // again, where deleting and recreating an owned shape to undo a refusal
        // would throw away the user's work the moment the host was briefly busy.
        // The refusal is reported rather than hidden, so a chart that never
        // converges is visible instead of silent.
        if (request.Kind == OfficeShapeKind.TextBox && !TryApplyContent(existing, request))
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.HostRejected);
        }

        // R4.6. An update re-applies the style for the same reason it re-applies
        // the text: the fill, stroke, width, and hatch are owned generated data.
        // Without this, a refresh that changed only a colour would move the shape
        // and leave the old colour behind - the shape would sit in the right place
        // looking wrong, which is invisible to every other assertion in this
        // class because geometry and content would both be correct.
        return TryApplyStyle(existing, request)
            ? ShapeWriteOutcome.Ok()
            : ShapeWriteOutcome.Refused(ShapeWriteRefusal.HostRejected);
    }

    /// <inheritdoc />
    public ShapeWriteOutcome Delete(string primitiveId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(primitiveId);

        // ADR-0008 D4's first statement, before any workbook access. ADR-0054: inside a
        // reconcile batch the target was already checked when it was resolved and events are
        // off for the whole pass, so the repeated query can only return the same answer.
        if (!_batchActive && RefusalFor(_protectionGuard.Query()) is { } guardRefusal)
        {
            return guardRefusal;
        }

        Excel.Shapes? shapes = ResolveShapes(out ShapeWriteOutcome? refusal);
        if (shapes is null)
        {
            return refusal!;
        }

        Excel.Shape? existing = FindShapeByName(shapes, primitiveId);
        if (existing is null)
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
        }

        // The ownership filter, not the name, authorises the delete. A user
        // shape that happens to share a name is never deleted (R4.8 D1).
        if (!CarriesOwnershipTagFor(existing, primitiveId))
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
        }

        existing.Delete();

        // A deleted name must not be served from the batch memo afterwards: a later create
        // in the same batch legitimately reuses it.
        if (_batchProxies is not null)
        {
            _ = _batchProxies.Remove(primitiveId);
        }

        return ShapeWriteOutcome.Ok();
    }

    /// <inheritdoc />
    public IReadOnlyList<string> ListOwned()
    {
        Excel.Shapes? shapes = ResolveShapes(out _);
        if (shapes is null)
        {
            return [];
        }

        List<string> owned = [];
        int count = GetShapeCount(shapes);
        int inventoryLogged = 0;
        for (int index = 1; index <= count; index++)
        {
            Excel.Shape? shape = GetShapeAt(shapes, index);
            if (shape is null)
            {
                continue;
            }

            string name = shape.Name;
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            // Read the tag only for a name this writer could have created (O4). The
            // alternative-text read is a host round trip; a name outside our own two
            // shapes cannot be one of ours, so the read buys nothing for it. The proof
            // is in CouldBeOwnedName, and the tag comparison stays exactly as it was
            // for every name that passes this test -- including the recompute that
            // makes a user's renamed copy of one of our shapes fail ownership.
            bool ownedTag = CouldBeOwnedName(name) && CarriesOwnershipTagFor(shape, name);
            if (ownedTag)
            {
                owned.Add(name);
            }

            // Phase-0 inventory: a chart-shaped shape that is NOT owned, or any
            // splitter band, is logged once per shape per refresh. Names carry the
            // chart's own "<rowId>:<role>" form, so a shape the reconciler cannot see
            // (an untagged band from an older build) is identifiable here rather than
            // inferred from the pixels. Bounded to 40 lines so a heavily annotated
            // sheet cannot flood the log.
            if (inventoryLogged < 40
                && (name.Contains("plitter", StringComparison.OrdinalIgnoreCase)
                    || (!ownedTag && name.Contains(':', StringComparison.Ordinal))))
            {
                TraceInventoryShape(shape, name, ownedTag);
                inventoryLogged++;
            }
        }

        // Ordinal sort so the returned order is a property of the shape names
        // rather than of the host's z-order, which R4.5 and R4.7 both mutate.
        owned.Sort(StringComparer.Ordinal);
        return owned;
    }

    /// <inheritdoc />
    public ShapeWriteOutcome UnlockShape(string primitiveId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(primitiveId);

        // Same guard-first rule as every other entry point, and it is load-bearing here too: an
        // unlock clears `Locked`, which the host refuses on a sheet the add-in does not own.
        if (!_batchActive && RefusalFor(_protectionGuard.Query()) is { } guardRefusal)
        {
            return guardRefusal;
        }

        Excel.Shapes? shapes = ResolveShapes(out ShapeWriteOutcome? refusal);
        if (shapes is null)
        {
            return refusal!;
        }

        Excel.Shape? existing = FindShapeByName(shapes, primitiveId);

        // The tag authorises the unlock, not the name: handing the user a writable shape the add-in
        // does not own would turn "unlock" into a way to edit anything on the sheet (ADR-0055 D6).
        if (existing is null || !CarriesOwnershipTagFor(existing, primitiveId))
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
        }

        if (ReadLocked(existing))
        {
            ApplyUnlock(existing);
        }

        return ShapeWriteOutcome.Ok();
    }

    /// <inheritdoc />
    public ShapeWriteOutcome LockOwnedShapes()
    {
        if (!_batchActive && RefusalFor(_protectionGuard.Query()) is { } guardRefusal)
        {
            return guardRefusal;
        }

        Excel.Shapes? shapes = ResolveShapes(out ShapeWriteOutcome? refusal);
        if (shapes is null)
        {
            return refusal!;
        }

        int count = GetShapeCount(shapes);
        for (int index = 1; index <= count; index++)
        {
            Excel.Shape? shape = GetShapeAt(shapes, index);
            if (shape is null)
            {
                continue;
            }

            string name = shape.Name;
            if (string.IsNullOrWhiteSpace(name) || !CouldBeOwnedName(name))
            {
                continue;
            }

            // The tag, not the name, is the ownership proof -- the same rule Delete and Update
            // follow, so a user's shape that shares a name is left unlocked.
            if (!CarriesOwnershipTagFor(shape, name))
            {
                continue;
            }

            if (!ReadLocked(shape))
            {
                ApplyLock(shape);
            }
        }

        return ShapeWriteOutcome.Ok();
    }

    /// <summary>Reads a shape's lock state. Test seam over the COM property.</summary>
    /// <param name="shape">The live shape.</param>
    /// <returns>The host's lock state.</returns>
    internal virtual bool ReadLocked(Excel.Shape shape) => shape.Locked;

    /// <summary>Locks a shape. Test seam over the COM property.</summary>
    /// <param name="shape">The live shape.</param>
    internal virtual void ApplyLock(Excel.Shape shape) => shape.Locked = true;

    /// <summary>Clears a shape's lock. Test seam over the COM property.</summary>
    /// <param name="shape">The live shape.</param>
    /// <remarks>The PIA types this property as <c>Boolean</c>, not <c>MsoTriState</c>.</remarks>
    internal virtual void ApplyUnlock(Excel.Shape shape) => shape.Locked = false;

    /// <inheritdoc />
    public bool GeometryMatches(OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);

        if (request.Geometry is null || !TryGetBoundingBox(request.Geometry, out RectD desired))
        {
            return false;
        }

        Excel.Shapes? shapes = ResolveShapes(out _);
        if (shapes is null)
        {
            return false;
        }

        Excel.Shape? shape = FindShapeByName(shapes, request.PrimitiveId);
        if (shape is null || !CarriesOwnershipTagFor(shape, request.PrimitiveId))
        {
            return false;
        }

        // Placement is part of the match, and the required placement is the request's
        // class. A scene-owned shape still on the host's default xlMoveAndSize can be
        // resized by a row-visibility change at any moment, so skipping it would preserve
        // exactly the drift this verification exists to catch; the rewrite re-places it
        // with xlMove. The reverse holds for a host-owned shape: it must carry
        // xlMoveAndSize, or the live stretch ADR-0052 D6 exists to restore is absent, and
        // the rewrite re-places it (ADR-0052 D7).
        if (shape.Placement != RequiredPlacement(request))
        {
            return false;
        }

        // Height is the one member the host is allowed to have moved for a host-owned
        // shape: Excel resizes an xlMoveAndSize shape proportionally, so a
        // collapse/expand cycle can leave a fraction of a point of drift that is not a
        // scene change and must not be rewritten on every pass. The wider tolerance is
        // stated (D8), not tuned; left, top and width keep the strict one because the
        // host does not resize along those axes for this class.
        return MatchesLive(shape.Left, desired.X)
            && MatchesLive(shape.Top, desired.Y)
            && MatchesLive(shape.Width, desired.Width)
            && (request.HostOwnsExtent
                ? MatchesLiveWithin(shape.Height, desired.Height, HostOwnedGeometry.HostResizeTolerancePt)
                : MatchesLive(shape.Height, desired.Height));
    }

    /// <summary>The tolerance for a live-versus-scene geometry comparison, in points.</summary>
    /// <remarks>
    /// The host stores geometry as <see cref="float"/> and the scene snaps to display
    /// precision at the port boundary, so exact equality would test rounding rather than
    /// geometry. One hundredth of a point is far below anything Excel can show and far
    /// above the <see cref="float"/> rounding.
    /// </remarks>
    internal const double GeometryMatchTolerancePt = 0.01;

    /// <summary>Compares one live host value to the scene's expected point value.</summary>
    private static bool MatchesLive(float actual, double expected) =>
        MatchesLiveWithin(actual, expected, GeometryMatchTolerancePt);

    /// <summary>
    /// Compares one live host value to the scene's expected point value within an
    /// explicit tolerance.
    /// </summary>
    /// <param name="actual">The host's value.</param>
    /// <param name="expected">The scene's value.</param>
    /// <param name="tolerancePt">The tolerance, in points.</param>
    /// <returns><see langword="true"/> when the values are within the tolerance.</returns>
    private static bool MatchesLiveWithin(float actual, double expected, double tolerancePt) =>
        Math.Abs(actual - GeometryMath.SnapToDisplayPrecision(expected)) <= tolerancePt;

    /// <summary>
    /// Derives the bounding box a shape of this kind occupies, exactly as
    /// <see cref="ApplyGeometry"/> writes it.
    /// </summary>
    /// <param name="geometry">The scene geometry.</param>
    /// <param name="bounds">The bounding box, when the geometry carries one.</param>
    /// <returns><see langword="true"/> when a bounding box could be derived.</returns>
    private static bool TryGetBoundingBox(OfficeShapeGeometry geometry, out RectD bounds)
    {
        if (geometry.Bounds is { } rect)
        {
            bounds = rect;
            return true;
        }

        if (geometry.From is { } from && geometry.To is { } to)
        {
            bounds = new RectD(
                Math.Min(from.X, to.X),
                Math.Min(from.Y, to.Y),
                Math.Abs(to.X - from.X),
                Math.Abs(to.Y - from.Y));
            return true;
        }

        bounds = default;
        return false;
    }

    /// <inheritdoc />
    public ShapeWriteOutcome ApplyZOrder(IReadOnlyList<string> backToFront)
    {
        ArgumentNullException.ThrowIfNull(backToFront);

        // ADR-0008 D4's first statement, before any workbook access. ADR-0054: inside a
        // reconcile batch the target was already checked when it was resolved and events are
        // off for the whole pass, so the repeated query can only return the same answer.
        if (!_batchActive && RefusalFor(_protectionGuard.Query()) is { } guardRefusal)
        {
            return guardRefusal;
        }

        Excel.Shapes? shapes = ResolveShapes(out ShapeWriteOutcome? refusal);
        if (shapes is null)
        {
            return refusal!;
        }

        if (backToFront.Count == 0)
        {
            return ShapeWriteOutcome.Ok();
        }

        // Resolve AND prove ownership of every named shape BEFORE mutating any.
        // Resolving inside the mutation loop would let a name that does not exist
        // leave the earlier shapes already moved, which is the partial order this
        // method refuses to produce; two passes keep the refusal atomic.
        //
        // The ownership check is load-bearing and belongs in this same preflight,
        // not in the mutation loop. `Update`, `Delete`, and `ListOwned` all
        // authorise on `CarriesOwnershipTagFor` rather than on the name, so a
        // refresh preserves a user-drawn shape that happens to share a requested
        // primitive ID. Without the check here, `ApplyZOrder` was the one mutation
        // path that would still reorder exactly that shape - a preservation
        // guarantee that holds everywhere else and silently does not hold here.
        // Putting it in the preflight also means an unowned shape refused *later*
        // in the list leaves the whole scene untouched, which is the atomicity the
        // existence check above already provides.
        List<Excel.Shape> resolved = new(backToFront.Count);
        for (int index = 0; index < backToFront.Count; index++)
        {
            string primitiveId = backToFront[index];
            Excel.Shape? shape = FindShapeByName(shapes, primitiveId);
            if (shape is null || !CarriesOwnershipTagFor(shape, primitiveId))
            {
                // NotFound rather than a new refusal: from the caller's point of view
                // "no owned shape with this identifier" is one condition, and it is the
                // reason `Update` and `Delete` already report. An unowned shape is
                // indistinguishable from an absent one by design - re-stamping its
                // tag here would "repair" it, which is R9.4's job to report, not this
                // row's.
                return ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
            }

            resolved.Add(shape);
        }

        // Send each shape to the back in REVERSE order. msoSendToBack is absolute
        // ("make this the backmost shape"), so walking the scene's back-to-front
        // list from the front end and pushing every shape behind everything else
        // leaves the final order equal to backToFront whatever the host's prior
        // order was. Forward order would produce its exact reverse. This is still
        // the single deterministic pass R4.5 D2 requires: no BringToFront, and no
        // per-shape call outside this method.
        //
        // O5 (ADR-0054): only the shapes that are actually out of order are moved. A read
        // costs ~0.03-0.1 ms and a ZOrder command is a write at 0.6-2.7 ms, so reading N
        // positions to skip most of the writes is a large win, and a pass whose order is
        // already correct issues no writes at all.
        //
        // Two facts fix which end can be skipped, and getting either wrong silently leaves
        // the chart mis-stacked rather than merely slower.
        //
        // ONE: `Shape.ZOrderPosition` is the shape's index in the Shapes collection and
        // Shapes(1) is the shape at the BACK of the z-order (Microsoft, Shape.ZOrderPosition
        // remarks). So a correct back-to-front run has strictly INCREASING positions as the
        // index advances.
        //
        // TWO: msoSendToBack is absolute, so every shape this loop touches ends up behind
        // every shape it does not. The untouched shapes therefore always finish in FRONT of
        // the moved ones, which means only the FRONT-most run of the desired order may be
        // left alone: it keeps its relative order and stays above the pass. Skipping a
        // back-most run would push the touched shapes behind it, inverting exactly what was
        // skipped — the defect this comment replaces.
        //
        // So: find the longest front run already in correct relative order, and push
        // everything before it to the back, front-to-back, which is the original full pass
        // applied to a prefix.
        int moveCount = resolved.Count;
        int previousPosition = int.MaxValue;
        for (int index = resolved.Count - 1; index >= 0; index--)
        {
            int position = GetZOrderPosition(resolved[index]);
            if (position <= 0)
            {
                // The host did not report a usable position (a test double, or a shape the
                // host will not describe). Fall back to the full pass rather than trust a
                // partial order: the pass is absolute and always converges.
                moveCount = resolved.Count;
                break;
            }

            if (position >= previousPosition)
            {
                break;
            }

            previousPosition = position;
            moveCount = index;
        }

        for (int index = moveCount - 1; index >= 0; index--)
        {
            resolved[index].ZOrder(MsoZOrderCmd.msoSendToBack);
        }

        return ShapeWriteOutcome.Ok();
    }

    /// <summary>
    /// Resolves the worksheet whose <c>Shapes</c> collection the renderer owns.
    /// </summary>
    /// <param name="refusal">The refusal to return when resolution fails.</param>
    /// <returns>The shapes collection, or <see langword="null"/> with <paramref name="refusal"/> set.</returns>
    /// <remarks>
    /// The protection guard has already run in the entry method, per ADR-0008
    /// D4. The authoritative check against the <em>resolved target</em> runs here,
    /// because the Gantt worksheet is not necessarily the active one: R4.8 D3
    /// forbids the refresh path from activating another sheet, so a guard that
    /// only read the active sheet would authorise a write to an unprotected
    /// active sheet while the real target is protected.
    /// </remarks>
    private Excel.Shapes? ResolveShapes(out ShapeWriteOutcome? refusal)
    {
        // Resolved once per batch: the target cannot change while a refresh runs, because
        // ADR-0020's scope disables events for the whole pass (ADR-0054).
        if (_batchShapes is not null)
        {
            refusal = null;
            return _batchShapes;
        }

        Excel.Application? application = _application;
        if (application is null)
        {
            refusal = ShapeWriteOutcome.Refused(ShapeWriteRefusal.NoActiveWorkbook);
            return null;
        }

        Excel.Workbook? workbook = application.ActiveWorkbook;
        if (workbook is null)
        {
            refusal = ShapeWriteOutcome.Refused(ShapeWriteRefusal.NoActiveWorkbook);
            return null;
        }

        Excel.Sheets? sheets = workbook.Sheets;
        if (sheets is null)
        {
            refusal = ShapeWriteOutcome.Refused(ShapeWriteRefusal.WorksheetMissing);
            return null;
        }

        Excel._Worksheet? gantt = FindGanttWorksheet(sheets);
        if (gantt is null)
        {
            refusal = ShapeWriteOutcome.Refused(ShapeWriteRefusal.WorksheetMissing);
            return null;
        }

        // The authoritative check is against the resolved target, not the active
        // sheet: the Gantt worksheet is not necessarily the active one (R4.8 D3
        // forbids the refresh path from activating another sheet).
        ProtectionGuardOutcome target = _protectionGuard.QueryTarget(gantt);
        if (target != ProtectionGuardOutcome.NotProtected)
        {
            refusal = ShapeWriteOutcome.Refused(ShapeWriteRefusal.TargetProtected);
            return null;
        }

        Excel.Shapes? shapes = GetShapes(gantt);
        if (shapes is null)
        {
            refusal = ShapeWriteOutcome.Refused(ShapeWriteRefusal.WorksheetMissing);
            return null;
        }

        if (_batchActive)
        {
            _batchShapes = shapes;
        }

        refusal = null;
        return shapes;
    }

    /// <summary>
    /// Validates a request before any host call, so a bad request never reaches Excel.
    /// </summary>
    /// <param name="request">The request to validate.</param>
    /// <param name="refusal">The refusal reason when validation fails.</param>
    /// <returns><see langword="true"/> when the request is well-formed.</returns>
    private static bool Validate(OfficeShapeRequest request, out ShapeWriteRefusal? refusal)
    {
        if (string.IsNullOrWhiteSpace(request.PrimitiveId))
        {
            refusal = ShapeWriteRefusal.BlankIdentifier;
            return false;
        }

        // The ownership scan reads a shape's tag only when its name could be one of ours
        // (ListOwned + CouldBeOwnedName, ADR-0054 O4), so a name outside those forms would
        // be written and then never recognised -- an un-updatable, un-deletable shape. The
        // form is enforced here rather than assumed, at the one place a name reaches the
        // host.
        if (!CouldBeOwnedName(request.PrimitiveId))
        {
            refusal = ShapeWriteRefusal.UnsupportedIdentifierForm;
            return false;
        }

        if (!Enum.IsDefined(request.Kind))
        {
            refusal = ShapeWriteRefusal.InvalidGeometry;
            return false;
        }

        if (!HasRequiredGeometry(request.Kind, request.Geometry))
        {
            refusal = ShapeWriteRefusal.InvalidGeometry;
            return false;
        }

        if (request.LineWidthPt is { } width && (!double.IsFinite(width) || width < 0))
        {
            refusal = ShapeWriteRefusal.InvalidGeometry;
            return false;
        }

        if (request.FontSizePt is { } size && (!double.IsFinite(size) || size <= 0))
        {
            refusal = ShapeWriteRefusal.InvalidGeometry;
            return false;
        }

        refusal = null;
        return true;
    }

    /// <summary>
    /// Determines whether a geometry carries the members its kind requires and
    /// whether every supplied value is finite.
    /// </summary>
    /// <param name="kind">The shape kind.</param>
    /// <param name="geometry">The geometry to check.</param>
    /// <returns><see langword="true"/> when the geometry is usable for the kind.</returns>
    /// <remarks>
    /// The kind-specific requirement is deliberate rather than permissive: a
    /// rectangle with no bounds cannot be positioned, and silently defaulting
    /// them would place the shape at the origin, which is the kind of silent zero
    /// the entity guide forbids.
    /// </remarks>
    private static bool HasRequiredGeometry(OfficeShapeKind kind, OfficeShapeGeometry geometry)
    {
        ArgumentNullException.ThrowIfNull(geometry);

        return kind switch
        {
            OfficeShapeKind.Rectangle or OfficeShapeKind.TextBox => IsUsableBounds(geometry.Bounds),
            OfficeShapeKind.Line => IsUsablePoint(geometry.From) && IsUsablePoint(geometry.To),

            // A diamond needs BOTH: the box places the auto-shape, and the
            // points stay on the request so the translation remains auditable.
            OfficeShapeKind.Diamond =>
                IsUsableBounds(geometry.Bounds)
                && geometry.Points is { Count: >= 3 } points
                && points.All(point => IsUsablePoint(point)),

            _ => false,
        };
    }

    private static bool IsUsableBounds(RectD? bounds) =>
        bounds is { } value
        && double.IsFinite(value.X)
        && double.IsFinite(value.Y)
        && double.IsFinite(value.Width)
        && double.IsFinite(value.Height)
        && value.Width > 0
        && value.Height > 0;

    private static bool IsUsablePoint(PointD? point) =>
        point is { } value && double.IsFinite(value.X) && double.IsFinite(value.Y);

    /// <summary>
    /// The primitive identifier of the chart title shape, the one owned shape whose name
    /// is not role-derived.
    /// </summary>
    /// <remarks>
    /// Declared here, next to <see cref="CouldBeOwnedName"/>, and consumed by
    /// <c>GanttRefreshOrchestrator</c>, because the ownership scan is only sound while
    /// every name this writer can create is enumerated in one place. Two literals with
    /// the same text would let a rename of the title silently make the title look
    /// unowned on every refresh.
    /// </remarks>
    public const string ChartTitlePrimitiveId = "ChartTitle";

    /// <summary>
    /// Whether a worksheet shape name could be one this writer created, and so needs its
    /// ownership tag compared.
    /// </summary>
    /// <param name="name">The live shape name, already known not to be blank.</param>
    /// <returns><see langword="true"/> when the name is one of our two possible forms.</returns>
    /// <remarks>
    /// <para>
    /// <b>Why this is a name test and not a tag test.</b> Reading <c>AlternativeText</c>
    /// is a host round trip, and <see cref="ListOwned"/> pays it once per shape on the
    /// sheet — most of which are a user's drawings and annotations that can never be
    /// ours. Skipping the read for those names is worth roughly one round trip per
    /// non-chart shape (O4).
    /// </para>
    /// <para>
    /// <b>Why it is sound.</b> There are exactly two sources of a name this writer
    /// stamps. A scene primitive is named by <c>ScenePrimitive.CreateId</c>, which
    /// always emits <c>"{ownerId}:{role}"</c> and refuses a blank role, so it always
    /// contains a colon. The chart title is the single fixed identifier,
    /// <see cref="ChartTitlePrimitiveId"/>. A name outside both forms cannot be a shape
    /// this writer created, therefore cannot carry this writer's tag, therefore cannot
    /// be owned — the skipped read could only ever have returned <see langword="false"/>.
    /// A user's renamed copy of one of our shapes is still excluded, because ownership
    /// is recomputed from the <em>live</em> name.
    /// </para>
    /// </remarks>
    internal static bool CouldBeOwnedName(string name)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        return name.Contains(':', StringComparison.Ordinal)
            || string.Equals(name, ChartTitlePrimitiveId, StringComparison.Ordinal);
    }

    /// <summary>
    /// Determines whether a shape's alternative text is the ownership tag for
    /// <em>this</em> identifier, not merely a well-formed tag for some identifier.
    /// </summary>
    /// <param name="shape">The shape whose alternative text is read.</param>
    /// <param name="primitiveId">The identifier the caller is acting on.</param>
    /// <returns>
    /// <see langword="true"/> when the shape's alternative text is exactly
    /// <see cref="ShapeOwnershipTag.ForPrimitiveId"/> of <paramref name="primitiveId"/>.
    /// </returns>
    /// <remarks>
    /// A format check alone is not the ownership proof. <c>IsOwnedTag</c> answers
    /// "did the add-in write a tag here", while this answers "did the add-in write
    /// <em>this primitive's</em> tag here". The second is what R4.8 needs: a shape
    /// that carries a valid tag for a different identifier is not the shape the
    /// caller named, so a refresh must neither rewrite nor delete it. The
    /// comparison is ordinal because the tag is machine-generated by
    /// <see cref="ShapeOwnershipTag.ForPrimitiveId"/>.
    /// </remarks>
    private static bool CarriesOwnershipTagFor(Excel.Shape shape, string primitiveId)
    {
        ArgumentNullException.ThrowIfNull(shape);
        ArgumentException.ThrowIfNullOrWhiteSpace(primitiveId);

        return string.Equals(
            shape.AlternativeText,
            ShapeOwnershipTag.ForPrimitiveId(primitiveId),
            StringComparison.Ordinal);
    }

    /// <summary>
    /// Writes the two ADR-0019 ownership members onto a newly created shape.
    /// </summary>
    /// <param name="shape">The created shape.</param>
    /// <param name="primitiveId">The stable scene primitive identifier.</param>
    /// <returns>The typed result.</returns>
    /// <remarks>
    /// <c>internal virtual</c> is a test seam, per the <see cref="ApplyStyle"/>
    /// pattern: the stamp writes <c>Name</c> and <c>AlternativeText</c> on a COM
    /// proxy, and a test needs to assert that the stamp was <em>reached</em> —
    /// not merely that no error surfaced — to pin the rule that a discarded shape
    /// is never stamped.
    /// </remarks>
    internal virtual void ApplyOwnershipStamp(Excel.Shape shape, string primitiveId) =>
        _ = ApplyOwnership(shape, primitiveId);

    /// <summary>
    /// Writes the two ADR-0019 ownership members onto a newly created shape.
    /// </summary>
    /// <param name="shape">The created shape.</param>
    /// <param name="primitiveId">The stable scene primitive identifier.</param>
    /// <returns>The typed result.</returns>
    private static ShapeWriteOutcome ApplyOwnership(Excel.Shape shape, string primitiveId)
    {
        shape.Name = primitiveId;
        shape.AlternativeText = ShapeOwnershipTag.ForPrimitiveId(primitiveId);

        // ADR-0055: an owned shape is locked by default, so the user cannot drag it out of place
        // by accident. `Locked` is inert until the worksheet is protected, which the refresh arms
        // separately; writing it here costs one property call on a path that already writes two,
        // and it means a new shape is locked from birth rather than on a later sweep.
        // The PIA types this property as Boolean, not MsoTriState.
        shape.Locked = true;
        return ShapeWriteOutcome.Ok();
    }

    /// <summary>
    /// Writes the two ADR-0019 ownership members onto a newly created shape,
    /// discarding the shape if the host will not accept them.
    /// </summary>
    /// <param name="shape">The created shape.</param>
    /// <param name="primitiveId">The stable scene primitive identifier.</param>
    /// <returns>
    /// The typed result. A failure to stamp is <see cref="ShapeWriteRefusal.HostRejected"/>
    /// and leaves no shape behind.
    /// </returns>
    /// <remarks>
    /// A shape with no ownership tag is worse than no shape at all: the port's
    /// structural guarantee is that a caller can never write an untagged shape
    /// (the class remarks), and R4.8's preservation filter is a test for that tag.
    /// An untagged shape is invisible to <see cref="ListOwned"/>, so R4.7 would
    /// neither update nor delete it, and it would sit on the sheet forever,
    /// indistinguishable from a user's own shape.
    /// <para>
    /// The host does reject this in practice. KNOWN-LIMITATIONS **L18** records
    /// that Excel refuses a shape name beyond roughly 255 characters with
    /// <c>ArgumentException: The specified value is out of range.</c>, and the
    /// name carries the full <see cref="OfficeShapeRequest.PrimitiveId"/>, which
    /// ADR-0017 leaves unbounded for a shared owner. So a large contributor set
    /// reaches here and fails <em>after</em> <c>AddShape</c> has already put an
    /// untagged shape on the sheet.
    /// </para>
    /// <para>
    /// Deleting the shape is therefore the only way to keep the sheet clean, and
    /// the refusal is reported rather than swallowed so R4.7 can see why the
    /// identifier was not rendered. L18 assigns the reconciliation policy for an
    /// over-long identifier to R4.7; this row only guarantees that whatever that
    /// policy is, it never leaves an unowned shape behind.
    /// </para>
    /// </remarks>
    private ShapeWriteOutcome ApplyOwnershipOrDiscard(Excel.Shape shape, string primitiveId)
    {
        // CA1031: a host refusal on either ownership member must be converted to a
        // typed refusal, not allowed to escape Create as an unhandled COM or
        // argument exception from inside a render command.
#pragma warning disable CA1031
        try
        {
            ApplyOwnershipStamp(shape, primitiveId);
            return ShapeWriteOutcome.Ok();
        }
        catch (Exception)
        {
            DiscardUnownedShape(shape);
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.HostRejected);
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Deletes a shape that could not be stamped, so no unowned shape is left.
    /// </summary>
    /// <param name="shape">The unstamped shape.</param>
    /// <remarks>
    /// Best-effort by design, and deliberately silent: the caller has already
    /// decided this shape is not part of the chart, and there is no
    /// <c>try</c>/<c>finally</c> here whose guarantee a throwing delete could
    /// break. A delete that itself fails leaves an unowned shape, which is the
    /// same failure the caller is already being told about.
    /// </remarks>
    private static void DiscardUnownedShape(Excel.Shape shape)
    {
        // CA1031: see the remarks - the failure is already reported to the caller
        // as HostRejected, so a throwing delete adds nothing and must not replace
        // that typed refusal with an unhandled exception.
#pragma warning disable CA1031
        try
        {
            shape.Delete();
        }
        catch (Exception)
        {
            // Intentionally empty: the typed refusal is the report.
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Writes a geometry onto an existing shape at the single rounding boundary.
    /// </summary>
    /// <param name="shape">The shape to update.</param>
    /// <param name="kind">The shape kind, which selects which members are written.</param>
    /// <param name="geometry">The scene geometry, in points.</param>
    /// <returns><see langword="true"/> when every member was written.</returns>
    /// <remarks>
    /// <b>See <see cref="ApplyPlacement"/> for why the shape is placed with
    /// <c>xlMove</c> and not <c>xlMoveAndSize</c>.</b>
    /// </remarks>
    /// <remarks>
    /// <para>
    /// R4.1 D3: this is the <em>only</em> place a scene point becomes a host
    /// <c>Single</c>, via <c>GeometryMath.SnapToDisplayPrecision</c>. A caller
    /// never pre-rounds, and there is no second conversion anywhere on the path.
    /// </para>
    /// <para>
    /// Every renderable kind is rewritten in place, milestones included: a
    /// diamond auto-shape is positioned by <c>Left</c>/<c>Top</c>/<c>Width</c>/
    /// <c>Height</c> like any other auto-shape, because the host derives its
    /// vertices from that box. (When milestones were freeforms this was not
    /// expressible - a freeform exposes its vertices as a read-only
    /// <c>Vertices</c> collection - which is one of the reasons the diamond
    /// auto-shape is preferred. An <em>undefined</em> kind is the only case that
    /// returns <see langword="false"/>, making the limitation explicit rather
    /// than silently ignoring the requested geometry.
    /// </para>
    /// </remarks>
    private static bool ApplyGeometry(Excel.Shape shape, OfficeShapeKind kind, OfficeShapeGeometry geometry)
    {
        ArgumentNullException.ThrowIfNull(geometry);

        if (!HasRequiredGeometry(kind, geometry))
        {
            return false;
        }

        switch (kind)
        {
            case OfficeShapeKind.Rectangle:
            case OfficeShapeKind.TextBox:
                RectD bounds = geometry.Bounds!.Value;
                shape.Left = GeometryMath.SnapToDisplayPrecision(bounds.X);
                shape.Top = GeometryMath.SnapToDisplayPrecision(bounds.Y);
                shape.Width = GeometryMath.SnapToDisplayPrecision(bounds.Width);
                shape.Height = GeometryMath.SnapToDisplayPrecision(bounds.Height);
                return true;

            case OfficeShapeKind.Line:
                // Probed 2026-09-27: a created line carries its absolute endpoints,
                // but the Shape surface exposes only the bounding box, so an update
                // is expressed as left/top/width/height. A negative-width line
                // (right-to-left) is normalised here rather than rejected, because
                // the host's own bounding box cannot represent it.
                PointD from = geometry.From!.Value;
                PointD to = geometry.To!.Value;
                double left = Math.Min(from.X, to.X);
                double top = Math.Min(from.Y, to.Y);
                shape.Left = GeometryMath.SnapToDisplayPrecision(left);
                shape.Top = GeometryMath.SnapToDisplayPrecision(top);
                shape.Width = GeometryMath.SnapToDisplayPrecision(Math.Abs(to.X - from.X));
                shape.Height = GeometryMath.SnapToDisplayPrecision(Math.Abs(to.Y - from.Y));
                return true;

            case OfficeShapeKind.Diamond:
                // A diamond auto-shape is positioned exactly like any other
                // auto-shape, so an in-place update only has to move and resize
                // it. The vertices are derived from the box by the host, so
                // there is no vertex collection to rewrite.
                RectD diamond = geometry.Bounds!.Value;
                shape.Left = GeometryMath.SnapToDisplayPrecision(diamond.X);
                shape.Top = GeometryMath.SnapToDisplayPrecision(diamond.Y);
                shape.Width = GeometryMath.SnapToDisplayPrecision(diamond.Width);
                shape.Height = GeometryMath.SnapToDisplayPrecision(diamond.Height);
                return true;

            default:
                // An undefined kind is not expressible. Returning false here
                // makes the limitation explicit rather than silently ignoring
                // the requested geometry.
                return false;
        }
    }

    /// <summary>
    /// Places the shape with the ownership the request's class requires: <c>xlMove</c>
    /// for a scene-owned shape, <c>xlMoveAndSize</c> for a host-owned one.
    /// </summary>
    /// <param name="shape">The shape to place.</param>
    /// <param name="hostOwnsExtent">
    /// Whether the shape spans the plot's full vertical extent and the host therefore
    /// owns its height (ADR-0052 D6).
    /// </param>
    /// <remarks>
    /// <para>
    /// <b>Why <c>xlMove</c> for a scene-owned shape and not the host's default
    /// <c>xlMoveAndSize</c>.</b> The scene owns the geometry, so the host resizing a shape
    /// is the renderer losing control of its own output. Excel creates shapes with
    /// <c>xlMoveAndSize</c>, which ties the displayed height to the rows the shape spans,
    /// so hiding, un-hiding or re-heighting a row inside that span resizes the shape
    /// without any scene change. Measured 2026-11-03: the scene declared the splitter band
    /// at 18pt on every refresh while the live band was four rows tall in one state and a
    /// thin line in another — a rewritten 18pt rect did not survive a row-visibility
    /// change.
    /// </para>
    /// <para>
    /// <c>xlMove</c> preserves the single host behaviour the row inserter depends on:
    /// an inserted worksheet row still moves the shape down with its cell (measured
    /// 2026-10-03, <c>scripts/probe-rowinsert-anchoring.ps1</c> measured the move; the
    /// sizing half is what this removes). Placement is re-applied on every update, so a
    /// workbook whose shapes were created under the old rule heals on its next Refresh.
    /// </para>
    /// <para>
    /// <b>Why the host-owned class gets the opposite answer (ADR-0052 D6–D7, amended
    /// 2026-11-08).</b> A primitive that spans the plot's full vertical extent has a
    /// height the scene derives from the same measured rows Excel uses, so host resizing
    /// is not the renderer losing control — it is the chart following its own rows, which
    /// is what makes a delineator or a band stretch live on collapse/expand before any
    /// refresh runs. Verification accepts that geometry for this class only, with the
    /// wider height tolerance D8 states.
    /// </para>
    /// </remarks>
    private static void ApplyPlacement(Excel.Shape shape, bool hostOwnsExtent) =>
        shape.Placement = hostOwnsExtent
            ? Excel.XlPlacement.xlMoveAndSize
            : Excel.XlPlacement.xlMove;

    /// <summary>The placement a request's class requires in the host.</summary>
    private static Excel.XlPlacement RequiredPlacement(OfficeShapeRequest request) =>
        request.HostOwnsExtent ? Excel.XlPlacement.xlMoveAndSize : Excel.XlPlacement.xlMove;

    /// <summary>Whether the identifier names a splitter band primitive.</summary>
    /// <param name="primitiveId">The stable primitive identifier.</param>
    /// <returns><see langword="true"/> for a splitter band.</returns>
    private static bool IsSplitterBand(string primitiveId) =>
        primitiveId.EndsWith(":splitter-band", StringComparison.Ordinal);

    /// <summary>
    /// Describes the band's live placement and geometry before a write, or
    /// <see langword="null"/> when this shape is not traced.
    /// </summary>
    /// <param name="shape">The live shape.</param>
    /// <param name="primitiveId">The stable primitive identifier.</param>
    /// <returns>The description, or null.</returns>
    private string? DescribeBandShape(Excel.Shape shape, string primitiveId) =>
        _trace is null || !IsSplitterBand(primitiveId) ? null : BandGeometry(shape);

    /// <summary>
    /// Writes the band's live placement and geometry around one create or update, so a
    /// live log can tell "the host ignored the height we wrote" from "the scene asked
    /// for the wrong height".
    /// </summary>
    /// <param name="action">The write kind, for the log.</param>
    /// <param name="shape">The live shape after the write.</param>
    /// <param name="primitiveId">The stable primitive identifier.</param>
    /// <param name="before">The description captured before the write, or null.</param>
    private void TraceBandShape(string action, Excel.Shape shape, string primitiveId, string? before)
    {
        if (_trace is null || !IsSplitterBand(primitiveId))
        {
            return;
        }

        _trace(string.Create(
            CultureInfo.InvariantCulture,
            $"gantt.shape write={action} id={primitiveId} before=[{before ?? "n/a"}] after=[{BandGeometry(shape)}]"));
    }

    /// <summary>
    /// Logs one inventory line for a chart-shaped or untagged shape, so an artifact the
    /// reconciler cannot see is identifiable rather than inferred from the pixels.
    /// </summary>
    /// <param name="shape">The live shape.</param>
    /// <param name="name">The shape's name.</param>
    /// <param name="ownedTag">Whether the shape carries a valid ownership tag.</param>
    private void TraceInventoryShape(Excel.Shape shape, string name, bool ownedTag)
    {
        if (_trace is null)
        {
            return;
        }

        _trace(string.Create(
            CultureInfo.InvariantCulture,
            $"gantt.shape inventory name={name} owned={ownedTag} place={shape.Placement} l={shape.Left:0.###} t={shape.Top:0.###} w={shape.Width:0.###} h={shape.Height:0.###}"));
    }

    /// <summary>Formats a shape's placement, top and height as one invariant fragment.</summary>
    /// <param name="shape">The live shape.</param>
    /// <returns>The fragment.</returns>
    private static string BandGeometry(Excel.Shape shape) =>
        string.Create(
            CultureInfo.InvariantCulture,
            $"place={shape.Placement} t={shape.Top:0.###} h={shape.Height:0.###}");

    // ---- Test seams (internal virtual, per the ExcelWorkbookInitialiser pattern) ----

    /// <summary>
    /// Returns the worksheet carrying the Gantt table. Test seam over the COM
    /// parameterised <c>Worksheets.Item</c> property and the table lookup
    /// (CS0855).
    /// </summary>
    /// <param name="sheets">The workbook's sheet collection.</param>
    /// <returns>The Gantt worksheet, or <see langword="null"/> when absent.</returns>
    internal virtual Excel._Worksheet? FindGanttWorksheet(Excel.Sheets sheets)
    {
        ArgumentNullException.ThrowIfNull(sheets);

        int count = GetSheetCount(sheets);
        for (int index = 1; index <= count; index++)
        {
            Excel._Worksheet? sheet = GetSheetAt(sheets, index);
            if (sheet is not null && HasGanttTable(sheet))
            {
                return sheet;
            }
        }

        return null;
    }

    /// <summary>Reads the sheet count. Test seam over <c>Worksheets.Count</c>.</summary>
    /// <param name="sheets">The sheet collection.</param>
    /// <returns>The number of sheets.</returns>
    internal virtual int GetSheetCount(Excel.Sheets sheets) => sheets.Count;

    /// <summary>Reads one sheet by its 1-based index. Test seam over <c>Sheets.Item</c>.</summary>
    /// <param name="sheets">The sheet collection.</param>
    /// <param name="index">The 1-based sheet index.</param>
    /// <returns>The sheet, or <see langword="null"/> when the index does not resolve.</returns>
    internal virtual Excel._Worksheet? GetSheetAt(Excel.Sheets sheets, int index) =>
        sheets[index] as Excel._Worksheet;

    /// <summary>
    /// Determines whether a sheet carries the Gantt table. Test seam over the
    /// <c>ListObjects</c> indexed property and the <c>Name</c> read.
    /// </summary>
    /// <param name="sheet">The candidate worksheet.</param>
    /// <returns><see langword="true"/> when the sheet carries <c>tblGanttData</c>.</returns>
    internal virtual bool HasGanttTable(Excel._Worksheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        Excel.ListObjects? objects = sheet.ListObjects;
        if (objects is null)
        {
            return false;
        }

        int count = GetListObjectCount(objects);
        for (int index = 1; index <= count; index++)
        {
            Excel.ListObject? table = GetListObjectAt(objects, index);
            if (table is not null && string.Equals(table.Name, GanttTableSchema.TableName, StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>Reads the table count. Test seam over <c>ListObjects.Count</c>.</summary>
    /// <param name="objects">The table collection.</param>
    /// <returns>The number of tables on the sheet.</returns>
    internal virtual int GetListObjectCount(Excel.ListObjects objects) => objects.Count;

    /// <summary>Reads one table by its 1-based index. Test seam over <c>ListObjects.Item</c>.</summary>
    /// <param name="objects">The table collection.</param>
    /// <param name="index">The 1-based table index.</param>
    /// <returns>The table, or <see langword="null"/> when the index does not resolve.</returns>
    internal virtual Excel.ListObject? GetListObjectAt(Excel.ListObjects objects, int index) =>
        objects[index];

    /// <summary>Reads a worksheet's shapes collection. Test seam over the <c>_Worksheet.Shapes</c> cast.</summary>
    /// <param name="sheet">The worksheet.</param>
    /// <returns>The shapes collection, or <see langword="null"/>.</returns>
    internal virtual Excel.Shapes? GetShapes(Excel._Worksheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        return sheet.Shapes;
    }

    /// <summary>Reads the shape count. Test seam over <c>Shapes.Count</c>.</summary>
    /// <param name="shapes">The shapes collection.</param>
    /// <returns>The number of shapes on the sheet.</returns>
    internal virtual int GetShapeCount(Excel.Shapes shapes) => shapes.Count;

    /// <summary>Reads one shape by its 1-based index. Test seam over <c>Shapes.Item</c>.</summary>
    /// <param name="shapes">The shapes collection.</param>
    /// <param name="index">The 1-based shape index.</param>
    /// <returns>The shape, or <see langword="null"/> when the index does not resolve.</returns>
    internal virtual Excel.Shape? GetShapeAt(Excel.Shapes shapes, int index) => shapes.Item(index);

    /// <summary>
    /// Finds a shape by its exact name. Test seam over the name lookup, which
    /// raises a COM error rather than returning null for a missing name; scanning
    /// by index turns that into a null so callers need no try/catch.
    /// </summary>
    /// <param name="shapes">The shapes collection.</param>
    /// <param name="name">The exact shape name.</param>
    /// <returns>The shape, or <see langword="null"/> when no shape carries the name.</returns>
    internal virtual Excel.Shape? FindShapeByName(Excel.Shapes shapes, string name)
    {
        ArgumentNullException.ThrowIfNull(shapes);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);

        // Inside a batch the same name is looked up by the geometry check, the update and
        // the z-order pass; one indexed lookup serves all three (ADR-0054).
        if (_batchProxies is not null && _batchProxies.TryGetValue(name, out Excel.Shape? cached))
        {
            return cached;
        }

        // The host's name indexer is the fast path: one lookup instead of a scan of
        // every shape with a COM `Name` read per shape. That scan was O(N) here and
        // this method is called once per create, update and delete, plus once per
        // shape in the ApplyZOrder preflight, so a full refresh cost O(N^2) name
        // reads (~6,400 at 80 shapes, ~10^6 at 1,000).
        //
        // A missing name RAISES (0x800A03EC) rather than returning null, which is why
        // the scan existed. Catching the miss here keeps the "null means absent"
        // contract for every caller, and a miss is rare compared with a hit.
        try
        {
            if (GetShapeByExactName(shapes, name) is { } resolved)
            {
                return Remember(name, resolved);
            }
        }
        catch (Exception exception) when (exception is ArgumentException or COMException or InvalidOperationException)
        {
            // The host reports a missing name by RAISING, and the type depends on the call
            // site. Measured on Excel 16.0 x64, 2026-11-03 (owner session `s6aea0f5a`):
            // `Shapes.Item("row:missing")` raised
            // `System.ArgumentException: The item with the specified name wasn't found.`
            // (0x80070057) — which crashed `Create`'s already-exists check before this was
            // caught. `COMException` and `InvalidOperationException` are the other
            // encodings the interop boundary produces. All three are the documented miss.
            return null;
        }

        // The index resolved nothing without raising. A miss normally raises, so this is
        // the degraded path (a test double, or a host whose indexer behaves differently);
        // the scan keeps the "null means absent" contract for every caller.
        int count = GetShapeCount(shapes);
        for (int index = 1; index <= count; index++)
        {
            Excel.Shape? shape = GetShapeAt(shapes, index);
            if (shape is not null && string.Equals(shape.Name, name, StringComparison.Ordinal))
            {
                return Remember(name, shape);
            }
        }

        return null;
    }

    /// <summary>
    /// Resolves a shape through the host's name indexer. Test seam over the
    /// parameterised COM property, which cannot be expressed in a Moq expression
    /// (CS0855); overrides model both the hit and the raising miss.
    /// </summary>
    /// <param name="shapes">The worksheet's shapes collection.</param>
    /// <param name="name">The exact shape name.</param>
    /// <returns>The shape when the host found it.</returns>
    /// <exception cref="COMException">Raised by the host when no shape carries the name.</exception>
    internal virtual Excel.Shape? GetShapeByExactName(Excel.Shapes shapes, string name)
    {
        ArgumentNullException.ThrowIfNull(shapes);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        return shapes.Item(name);
    }

    /// <summary>
    /// Creates one shape of the requested kind. Test seam over the three
    /// shape-creation calls and the freeform builder sequence, so contract tests
    /// assert the operation without a live Excel.
    /// </summary>
    /// <param name="shapes">The shapes collection.</param>
    /// <param name="request">The shape to create.</param>
    /// <returns>The created shape, or <see langword="null"/> when the host refused.</returns>
    internal virtual Excel.Shape? AddShape(Excel.Shapes shapes, OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(shapes);
        ArgumentNullException.ThrowIfNull(request);

        OfficeShapeGeometry geometry = request.Geometry;

        switch (request.Kind)
        {
            case OfficeShapeKind.Rectangle:
                RectD rect = geometry.Bounds!.Value;
                return shapes.AddShape(
                    MsoAutoShapeType.msoShapeRectangle,
                    GeometryMath.SnapToDisplayPrecision(rect.X),
                    GeometryMath.SnapToDisplayPrecision(rect.Y),
                    GeometryMath.SnapToDisplayPrecision(rect.Width),
                    GeometryMath.SnapToDisplayPrecision(rect.Height));

            case OfficeShapeKind.Line:
                PointD from = geometry.From!.Value;
                PointD to = geometry.To!.Value;
                return shapes.AddLine(
                    GeometryMath.SnapToDisplayPrecision(from.X),
                    GeometryMath.SnapToDisplayPrecision(from.Y),
                    GeometryMath.SnapToDisplayPrecision(to.X),
                    GeometryMath.SnapToDisplayPrecision(to.Y));

            case OfficeShapeKind.TextBox:
                RectD text = geometry.Bounds!.Value;
                Excel.Shape? textBox = shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    GeometryMath.SnapToDisplayPrecision(text.X),
                    GeometryMath.SnapToDisplayPrecision(text.Y),
                    GeometryMath.SnapToDisplayPrecision(text.Width),
                    GeometryMath.SnapToDisplayPrecision(text.Height));

                if (textBox is not null && !ApplyTextOrDiscard(textBox, request))
                {
                    return null;
                }

                return textBox;

            case OfficeShapeKind.Diamond:
                return AddDiamond(shapes, request);

            default:
                // An undefined kind is refused rather than placed. Returning null
                // makes Create report HostRejected, which is a visible refusal and
                // not a silently misplaced shape.
                return null;
        }
    }

    /// <summary>
    /// Creates the milestone diamond from the request's bounding box.
    /// </summary>
    /// <param name="shapes">The shapes collection.</param>
    /// <param name="request">The request whose bounds place the diamond.</param>
    /// <returns>The created diamond, or <see langword="null"/> when the host refused.</returns>
    /// <remarks>
    /// <para>
    /// R4.5 D1, decided by measurement rather than preference. The row was first
    /// implemented as a host freeform, and the Step-0 probe then established
    /// that <c>BuildFreeform</c> stores each vertex on a whole-EMU grid: the
    /// resulting span is <strong>+1 EMU</strong> (1 EMU = 1/12700 pt) off the
    /// scene value, measured at +1.005, +0.993, +1.017, +0.969 and +0.969 EMU
    /// across exact spans of 10, 20, 37.5, 100 and 253 pt. That is a fixed
    /// one-unit quantisation, not drift: it is constant across a 25x size range.
    /// </para>
    /// <para>
    /// A diamond auto-shape is <em>derived</em> from its
    /// <c>Left</c>/<c>Top</c>/<c>Width</c>/<c>Height</c> box rather than storing
    /// absolute vertices, and measured <strong>0.000 EMU at every span</strong>.
    /// Using it therefore removes the only reason this project had ever needed a
    /// second geometry tolerance, so <see cref="GeometryMath.Epsilon"/> governs
    /// every shape family again.
    /// </para>
    /// <para>
    /// The shape is placed from <see cref="OfficeShapeGeometry.Bounds"/> - the
    /// same member <see cref="ApplyGeometry"/> uses on the update path - rather
    /// than recomputed here from the points. The scene still models a milestone
    /// as four points and they stay on the request so the translation remains
    /// auditable, but <see cref="SceneShapeRenderer"/> has already refused any
    /// polygon that is not a symmetric axis-aligned diamond, so the two boxes are
    /// equal by construction. Recomputing them here created a second place where
    /// the box could be derived, which is exactly the dual-source risk R4.1's
    /// single rounding boundary exists to prevent.
    /// </para>
    /// </remarks>
    private static Excel.Shape? AddDiamond(Excel.Shapes shapes, OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(shapes);
        ArgumentNullException.ThrowIfNull(request);

        if (request.Geometry.Bounds is not { } bounds)
        {
            // Validate has already refused a bounds-less diamond before any host
            // call, so this arm is unreachable through Create. It is the guard
            // that keeps this method total for any future caller.
            return null;
        }

        return shapes.AddShape(
            MsoAutoShapeType.msoShapeDiamond,
            GeometryMath.SnapToDisplayPrecision(bounds.X),
            GeometryMath.SnapToDisplayPrecision(bounds.Y),
            GeometryMath.SnapToDisplayPrecision(bounds.Width),
            GeometryMath.SnapToDisplayPrecision(bounds.Height));
    }

    /// <summary>
    /// Writes the resolved fill and stroke onto a shape.
    /// </summary>
    /// <param name="shape">The shape to style.</param>
    /// <param name="request">The request whose style members are written.</param>
    /// <remarks>
    /// <para>
    /// R4.6. The host's format members are written from
    /// <see cref="OfficeStyleMapper"/>'s output verbatim: no colour is chosen here,
    /// no token is defaulted, and no theme colour is touched. The mapper owns the
    /// arithmetic; this method owns the COM.
    /// </para>
    /// <para>
    /// <see cref="MsoPatternType"/> is the facility the Step-0 probe established for
    /// a hatch, and it is a <em>fixed pitch</em> pattern set: the probe found no
    /// member controlling the <c>HatchPitchPt</c> spacing or the <c>HatchLinePt</c>
    /// stroke width. Those two tokens are therefore <strong>not honoured</strong> by
    /// the native pattern, which is exactly the case entity-guide section 18
    /// anticipates ("may use a native pattern only if the compatibility test proves
    /// equivalent bounds and adequate appearance"). The deviation is recorded for
    /// R8.3's equivalence policy rather than silently approximated, and a
    /// <see cref="GanttHatchPattern"/> the host cannot express is
    /// <strong>refused</strong> rather than drawn as the nearest pattern.
    /// </para>
    /// <para>
    /// <c>internal virtual</c> is a test seam, per the
    /// <see cref="ApplyText"/> pattern: the write reaches COM through a chained
    /// <c>Shape.Fill</c>/<c>Shape.Line</c> walk that a contract test cannot observe,
    /// so the seam lets a test record that the style was applied rather than
    /// inferring it.
    /// </para>
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when a hatch pattern has no host member. Never caught, because an
    /// unexpressible pattern is a contract fault the caller must see, not a host
    /// failure to swallow into a refusal.
    /// </exception>
    internal virtual void ApplyStyle(Excel.Shape shape, OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(shape);
        ArgumentNullException.ThrowIfNull(request);

        OfficeShapeStyle style = OfficeStyleMapper.Map(request);

        Excel.FillFormat? fill = shape.Fill;
        if (fill is not null)
        {
            if (!style.FillVisible)
            {
                // Explicitly off rather than left alone. A shape the host created
                // with its default fill would otherwise keep that fill, and
                // "the scene resolved no fill" has to be able to mean "no fill".
                fill.Visible = MsoTriState.msoFalse;
            }
            else
            {
                Excel.ColorFormat? backColour = fill.BackColor;
                if (backColour is not null && style.FillBackgroundRgb is { } background)
                {
                    backColour.RGB = background;
                }

                Excel.ColorFormat? foreColour = fill.ForeColor;
                if (foreColour is not null && style.FillRgb is { } foreground)
                {
                    foreColour.RGB = foreground;
                }

                fill.Transparency = style.FillTransparency;
                fill.Visible = MsoTriState.msoTrue;

                // Solid() before Patterned(): the host keeps whichever fill type
                // it last had, so a shape created solid and later given a hatch
                // needs the type set explicitly or the colour write above would
                // land on a solid fill and the hatch would never appear.
                if (style.HatchPattern is { } hatch)
                {
                    fill.Patterned(MapHatchPattern(hatch));
                }
                else
                {
                    fill.Solid();
                }
            }
        }

        Excel.LineFormat? line = shape.Line;
        if (line is null)
        {
            return;
        }

        if (!style.LineVisible)
        {
            line.Visible = MsoTriState.msoFalse;
            return;
        }

        Excel.ColorFormat? strokeColour = line.ForeColor;
        if (strokeColour is not null && style.LineRgb is { } stroke)
        {
            strokeColour.RGB = stroke;
        }

        if (style.LineWeightPt is { } weight)
        {
            line.Weight = weight;
        }

        // The dash is a resolved style token, so the renderer writes it
        // verbatim rather than choosing one: a line the scene resolved as
        // dashed must reach the host dashed. Solid is written explicitly
        // because the host keeps whichever dash it last had.
        line.DashStyle = MapDashStyle(style.StrokeDash);

        line.Transparency = style.LineTransparency;
        line.Visible = MsoTriState.msoTrue;
    }

    /// <summary>
    /// Maps a resolved stroke dash pattern onto the host's line-dash constant.
    /// </summary>
    /// <param name="dash">The resolved dash pattern.</param>
    /// <returns>The host's dash constant.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when the dash has no host member. Never caught, because a
    /// contract fault is not a host failure: the caller discards the shape
    /// and reports the refusal (the same contract as
    /// <see cref="MapHatchPattern"/>).
    /// </exception>
    internal static MsoLineDashStyle MapDashStyle(GanttStrokeDash dash) => dash switch
    {
        GanttStrokeDash.Solid => MsoLineDashStyle.msoLineSolid,
        GanttStrokeDash.SquareDot => MsoLineDashStyle.msoLineSquareDot,
        GanttStrokeDash.RoundDot => MsoLineDashStyle.msoLineRoundDot,
        GanttStrokeDash.Dash => MsoLineDashStyle.msoLineDash,
        GanttStrokeDash.DashDot => MsoLineDashStyle.msoLineDashDot,
        GanttStrokeDash.LongDash => MsoLineDashStyle.msoLineLongDash,
        _ => throw new ArgumentOutOfRangeException(nameof(dash), dash, "No host dash member."),
    };

    /// <summary>
    /// Applies the style, discarding the shape if the host refuses it.
    /// </summary>
    /// <param name="shape">The created shape.</param>
    /// <param name="request">The request whose style is applied.</param>
    /// <returns>
    /// <see langword="true"/> when the style landed; <see langword="false"/>
    /// when the host refused it and the shape has been discarded.
    /// </returns>
    /// <remarks>
    /// Best-effort on the same reasoning as
    /// <see cref="ApplyOwnershipOrDiscard"/>: a shape that could not be styled
    /// must not be left on the sheet, because the caller is about to be told the
    /// create failed and would never learn about the orphan.
    /// <para>
    /// The outcome is returned rather than swallowed, because
    /// <see cref="Create"/> must not continue past it. Continuing would stamp
    /// ownership onto a shape this method has just deleted, and that write is
    /// either an escaping error or a false success for a shape that no longer
    /// exists. The caller turns <see langword="false"/> into
    /// <see cref="ShapeWriteRefusal.HostRejected"/> and skips the ownership
    /// stamp entirely.
    /// </para>
    /// </remarks>
    private bool ApplyStyleOrDiscard(Excel.Shape shape, OfficeShapeRequest request)
    {
        // CA1031: see the remarks. A throwing style write must not escape Create
        // as an unhandled COM exception from inside a render command, and must
        // not replace the typed refusal the caller already has.
#pragma warning disable CA1031
        try
        {
            ApplyStyle(shape, request);
        }
        catch (Exception)
        {
            DiscardUnownedShape(shape);
            return false;
        }
#pragma warning restore CA1031

        return true;
    }

    /// <summary>
    /// Writes a new text box's content, discarding the shape if the host refuses.
    /// </summary>
    /// <param name="shape">The created text box.</param>
    /// <param name="request">The request whose text members are written.</param>
    /// <returns>
    /// <see langword="true"/> when the content landed; <see langword="false"/>
    /// when the host refused it and the shape has been discarded.
    /// </returns>
    /// <remarks>
    /// <para>
    /// The create-path twin of <see cref="ApplyStyleOrDiscard"/>, and it exists
    /// because a text box is the one kind whose content is written inside
    /// <see cref="AddShape"/>, before the shape is stamped and before
    /// <see cref="Create"/> has any handle on it. A throwing
    /// <see cref="ApplyText"/> there had two bad outcomes: the exception
    /// escaped the create as an unhandled COM error from inside a render
    /// command, and the text box stayed on the sheet with no ownership tag, so
    /// nothing could ever update or remove it.
    /// </para>
    /// <para>
    /// Returning <see langword="false"/> makes <see cref="AddShape"/> return
    /// <see langword="null"/>, which <see cref="Create"/> already reports as
    /// <see cref="ShapeWriteRefusal.HostRejected"/> — the same typed refusal
    /// every other create failure uses, so the caller needs no new case.
    /// </para>
    /// </remarks>
    private bool ApplyTextOrDiscard(Excel.Shape shape, OfficeShapeRequest request)
    {
        // CA1031: see the remarks. A throwing content write must not escape the
        // create as an unhandled COM exception, and must not leave an untagged
        // text box behind for a caller that was told the create failed.
#pragma warning disable CA1031
        try
        {
            ApplyText(shape, request);
        }
        catch (Exception)
        {
            DiscardUnownedShape(shape);
            return false;
        }
#pragma warning restore CA1031

        return true;
    }

    /// <summary>
    /// Applies a text box's content on the update path, containing a host failure.
    /// </summary>
    /// <param name="shape">The owned shape whose content is rewritten.</param>
    /// <param name="request">The request whose text members are written.</param>
    /// <returns>
    /// <see langword="true"/> when the content landed; <see langword="false"/>
    /// when the host refused the write.
    /// </returns>
    /// <remarks>
    /// The update-path twin of <see cref="ApplyTextOrDiscard"/>, and it differs in
    /// exactly one respect: there is nothing to discard. The shape was found, not
    /// created, so a failed content write leaves an existing owned shape in place
    /// and the caller is told the update failed.
    /// <para>
    /// <see cref="ArgumentOutOfRangeException"/> is deliberately not caught. It is
    /// what <see cref="MapAlignment"/> raises for a scene alignment with no host
    /// mapping - a contract fault the caller must see, and the same rule
    /// <see cref="ApplyTextOrDiscard"/> follows.
    /// </para>
    /// </remarks>
    private bool TryApplyContent(Excel.Shape shape, OfficeShapeRequest request)
    {
        // CA1031: see the remarks. A host refusal must not escape the update as an
        // unhandled COM exception, and must not replace the typed refusal the
        // reconciler already knows how to report.
#pragma warning disable CA1031
        try
        {
            ApplyText(shape, request);
        }
        catch (ArgumentOutOfRangeException)
        {
            throw;
        }
        catch (Exception)
        {
            return false;
        }
#pragma warning restore CA1031

        return true;
    }

    /// <summary>
    /// Applies the style on the update path, containing a host failure.
    /// </summary>
    /// <param name="shape">The owned shape whose style is rewritten.</param>
    /// <param name="request">The request whose style members are written.</param>
    /// <returns>
    /// <see langword="true"/> when the style landed; <see langword="false"/> when
    /// the host refused the write.
    /// </returns>
    /// <remarks>
    /// The update-path twin of <see cref="ApplyStyleOrDiscard"/>, and it differs in
    /// exactly one respect: there is nothing to discard, for the same reason
    /// <see cref="TryApplyContent"/> does not discard. <see cref="ApplyGeometry"/>
    /// has already run by this point and is not undone, so a refused style leaves
    /// the shape moved and un-restyled; that is stated on
    /// <see cref="IShapeWritePort.Update(OfficeShapeRequest)"/> rather than left for
    /// a reader to infer.
    /// <para>
    /// <see cref="ArgumentOutOfRangeException"/> is deliberately not caught, for
    /// the reason <see cref="ApplyStyle"/> documents: it is
    /// <see cref="MapHatchPattern"/>'s contract fault, not a host failure.
    /// </para>
    /// </remarks>
    private bool TryApplyStyle(Excel.Shape shape, OfficeShapeRequest request)
    {
        // CA1031: see the remarks. A host refusal must not escape the update as an
        // unhandled COM exception, and must not replace the typed refusal the
        // reconciler already knows how to report.
#pragma warning disable CA1031
        try
        {
            ApplyStyle(shape, request);
        }
        catch (ArgumentOutOfRangeException)
        {
            throw;
        }
        catch (Exception)
        {
            return false;
        }
#pragma warning restore CA1031

        return true;
    }

    /// <summary>
    /// Maps a resolved hatch pattern onto the host's pattern constant.
    /// </summary>
    /// <param name="pattern">The resolved pattern.</param>
    /// <returns>The host pattern constant.</returns>
    /// <remarks>
    /// <para>
    /// The mapping is closed because <see cref="GanttHatchPattern"/> is. The
    /// default arm throws rather than falling back to a nearby pattern, for the
    /// same reason <see cref="MapAlignment"/> does: a pattern the scene asked for
    /// and the host cannot draw must fail loudly, so the guide and this method
    /// cannot drift apart unnoticed. A silent nearest-match would render a
    /// procurement bar in a pattern nobody approved.
    /// </para>
    /// <para>
    /// Direction is read from the entity guide's own wording, not guessed:
    /// <c>ForwardDiagonal</c> is "top-left to bottom-right", which on a
    /// screen-coordinate host is the <em>downward</em> diagonal.
    /// </para>
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the pattern is not a defined member.</exception>
    internal static MsoPatternType MapHatchPattern(GanttHatchPattern pattern) =>
        pattern switch
        {
            GanttHatchPattern.ForwardDiagonal => MsoPatternType.msoPatternLightDownwardDiagonal,
            GanttHatchPattern.BackwardDiagonal => MsoPatternType.msoPatternLightUpwardDiagonal,
            GanttHatchPattern.Cross => MsoPatternType.msoPatternDiagonalCross,

            // None is listed rather than left to the default arm so the intent is
            // explicit: a solid fill is not a hatch, and OfficeStyleMapper turns
            // None into a null pattern rather than calling this. Reaching here
            // means a caller mapped a "no hatch" as though it were one.
            GanttHatchPattern.None => throw new ArgumentOutOfRangeException(
                nameof(pattern),
                pattern,
                "None means a solid fill and must not be mapped to a host pattern."),

            _ => throw new ArgumentOutOfRangeException(
                nameof(pattern),
                pattern,
                "The scene produced a hatch pattern with no host mapping. Extend MapHatchPattern and the guide together."),
        };

    /// <summary>
    /// Writes the text content, typography, and alignment onto a text shape.
    /// </summary>
    /// <param name="shape">The created text box.</param>
    /// <param name="request">The request whose text members are written.</param>
    /// <remarks>
    /// <para>
    /// R4.4: the renderer <strong>consumes</strong> the scene's resolved text.
    /// The string is written verbatim, including any ellipsis R3.6's overflow
    /// policy already applied, and the resolved <c>TextBounds</c> positions the
    /// shape. Nothing here measures, truncates, re-wraps, or re-selects a label
    /// side; the scene owns all of that.
    /// </para>
    /// <para>
    /// The <c>TextFrame2</c> path is used rather than the legacy
    /// <c>TextFrame.Characters</c> path because the latter's font members are
    /// typed <c>Object</c>, and this repository treats an untyped interop member
    /// as a live hazard - R4.1's <c>Shape.Tag</c> and R4.2's
    /// <c>Application.Calculation</c> were both assumptions about interop member
    /// shapes that did not survive contact. R4.4 D3's original
    /// <c>xlAlignLeft</c>-style constants do not exist in the installed
    /// assembly; the real values are <see cref="MsoParagraphAlignment"/> members.
    /// </para>
    /// <para>
    /// <c>internal virtual</c> is a test seam, per the
    /// <c>ExcelWorkbookInitialiser</c> pattern: the content write reaches COM
    /// through a chained <c>Shape.TextFrame2</c> walk that a contract test
    /// cannot observe, so the seam lets a test record that the update path
    /// applied the content rather than inferring it from the geometry.
    /// </para>
    /// <para>
    /// Word wrap is switched off deliberately. A text box that auto-fits would
    /// resize itself away from the bounds the scene resolved, which is the same
    /// class of silent re-layout the no-remeasure rule forbids.
    /// </para>
    /// </remarks>
    internal virtual void ApplyText(Excel.Shape shape, OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(shape);
        ArgumentNullException.ThrowIfNull(request);

        Excel.TextFrame2? frame = shape.TextFrame2;
        if (frame is null)
        {
            return;
        }

        frame.WordWrap = MsoTriState.msoFalse;
        frame.AutoSize = MsoAutoSize.msoAutoSizeNone;

        // Vertical centring, which the entity guide's section 22 requires at EVERY
        // label position ("Right: ... vertically centred", "Left: ... vertically
        // centred"). This was never written, so Excel fell back to its own default,
        // msoAnchorTop, and every label sat at the top of the row-tall box the scene
        // resolved for it - visible in a live F5 session on 2026-10-04 as labels
        // hugging the top edge of their bar rather than centred in it.
        //
        // Set explicitly rather than relied upon, for the same reason the margins
        // below are zeroed: the scene already resolved the box and the vertical
        // placement, so a host default is a second authority that can disagree with
        // it. Verified against the installed interop assembly (office.dll 16.0) that
        // TextFrame2.VerticalAnchor is a settable Single-valued property of type
        // MsoVerticalAnchor, and that msoAnchorMiddle is the member named Middle
        // (value 3).
        frame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;

        // The internal margins are zeroed because the scene already resolved the
        // box, and a host default inset would silently move the text away from the
        // position the scene chose - the same no-remeasure violation as auto-sizing.
        // Excel's textbox default is roughly 0.1in left/right and 0.05in top/bottom,
        // which is 3.6pt/7.2pt: enough to make a Right-anchored label start visibly
        // right of the bar it labels, and to make a Left-anchored one end short of
        // it. Verified against the installed interop assembly that all four members
        // are settable Single properties on TextFrame2.
        frame.MarginLeft = 0f;
        frame.MarginRight = 0f;
        frame.MarginTop = 0f;
        frame.MarginBottom = 0f;

        TextRange2? textRange = frame.TextRange;
        if (textRange is null)
        {
            return;
        }

        // Verbatim: the scene already applied the overflow policy.
        textRange.Text = request.Text ?? string.Empty;

        if (request.Alignment is { } alignment && textRange.ParagraphFormat is { } paragraphFormat)
        {
            // ParagraphFormat2 is held in a local rather than written through
            // textRange.ParagraphFormat: a chained COM property call creates a
            // second proxy the adapter never names, and the ownership rule is that
            // every proxy is held in a local. The proxy is Excel-owned and shared,
            // so it is not force-released - see the class remarks.
            paragraphFormat.Alignment = MapAlignment(alignment);
        }

        if (textRange.Font is { } font)
        {
            // Same rule for Font2, which is read once and written through for
            // family, size, and weight alike.
            if (request.FontFamily is { } family)
            {
                font.Name = family;
            }

            if (request.FontSizePt is { } size)
            {
                font.Size = (float)size;
            }

            if (request.Bold is { } bold)
            {
                font.Bold = bold ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            }
        }

        // The text colour goes through the FONT's fill, not the shape's: the shape
        // fill is already spoken for by ApplyStyle, and the two are independent
        // properties on the host. A shape can therefore be a red delay body with a
        // white label, which is exactly what entity guide section 17 requires.
        //
        // The chain is TextRange2.Font (Font2) -> Fill (FillFormat) -> ForeColor
        // (ColorFormat) -> RGB, verified by reflection against the installed
        // office.dll 16.0.0 rather than assumed: Font2 exposes UnderlineColor
        // directly but has no Font.Color, so the fill is the only colour surface
        // on a Font2. Every proxy is read once into a local and written through,
        // per the ownership rule, and none is force-released because they are
        // Excel-owned and shared.
        OfficeShapeStyle style = OfficeStyleMapper.Map(request);
        if (style.TextRgb is { } textRgb
            && textRange.Font is { } textFont
            && textFont.Fill is { } textFill)
        {
            // Solid() before the colour write, for the same reason ApplyStyle does
            // it for a shape fill: the host keeps whichever fill type it last had,
            // and a text fill inherited from the theme can silently ignore a
            // ForeColor write. Making it solid first is what makes the colour land.
            textFill.Solid();
            if (textFill.ForeColor is { } textForeColor)
            {
                textForeColor.RGB = textRgb;
            }

            textFill.Visible = MsoTriState.msoTrue;
        }
    }

    /// <summary>
    /// Maps a resolved scene alignment onto the host's paragraph-alignment
    /// constant.
    /// </summary>
    /// <param name="alignment">The alignment the scene resolved.</param>
    /// <returns>The host alignment constant.</returns>
    /// <remarks>
    /// The mapping is closed because <see cref="GanttTextAlignment"/> is. The
    /// default arm is a throw rather than a silent fallback: an alignment added
    /// to the scene enum must fail loudly here, so the guide and this method
    /// cannot drift apart unnoticed.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the alignment is not a defined member.</exception>
    private static MsoParagraphAlignment MapAlignment(GanttTextAlignment alignment) =>
        alignment switch
        {
            GanttTextAlignment.Left => MsoParagraphAlignment.msoAlignLeft,
            GanttTextAlignment.Centre => MsoParagraphAlignment.msoAlignCenter,
            GanttTextAlignment.Right => MsoParagraphAlignment.msoAlignRight,
            _ => throw new ArgumentOutOfRangeException(
                nameof(alignment),
                alignment,
                "The scene produced an alignment with no host mapping. Extend MapAlignment and the guide together."),
        };
}
