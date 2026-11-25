using System.Diagnostics;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

public partial class ExcelShapeWriter
{
    /// <inheritdoc />
    public ShapeWriteOutcome Update(OfficeShapeRequest request) => UpdateCore(request, applyAppearance: true);

    /// <inheritdoc />
    public ShapeWriteOutcome UpdateGeometry(OfficeShapeRequest request) =>
        UpdateCore(request, applyAppearance: false, applyContent: false);

    /// <inheritdoc />
    public ShapeWriteOutcome UpdateContent(OfficeShapeRequest request) =>
        UpdateCore(request, applyAppearance: false, applyContent: true);

    /// <inheritdoc />
    public ShapeGroupWriteOutcome ApplyGroupAppearance(IReadOnlyList<OfficeShapeRequest> requests)
    {
        ArgumentNullException.ThrowIfNull(requests);
        if (requests.Count == 0)
        {
            return ShapeGroupWriteOutcome.Ok();
        }

        // Diagnostic: the grouped pass only wins when shapes actually share an appearance.
        // The live full-refresh trace said 148 updates still cost ~10-18 ms each, which is
        // the per-shape figure, so the group cardinality is the fact that decides whether
        // the key is too granular. Recorded only; it changes no decision.
        long groupStarted = Stopwatch.GetTimestamp();

        if (!_batch.Active && RefusalFor(_protectionGuard.Query()) is { } guardRefusal)
        {
            return ShapeGroupWriteOutcome.Refused(requests[0].PrimitiveId, guardRefusal.Refusal!.Value);
        }

        Excel.Shapes? shapes = ResolveShapes(out ShapeWriteOutcome? refusal);
        if (shapes is null)
        {
            return ShapeGroupWriteOutcome.Refused(requests[0].PrimitiveId, refusal!.Refusal!.Value);
        }

        // Two independent buckets, because the two write families touch disjoint members:
        // a Fill/Line write never touches TextFrame2, and a TextFrame2 write never touches
        // Fill/Line. The CREATION key must match on both (a duplicate inherits both), but an
        // UPDATE can coalesce further by bucketing each family on the members it actually
        // writes. Measured on the owner's 148-shape workbook: the combined key produced 22
        // groups and the group pass cost 1.35-1.53 s; a text group costs ~36-40 ms while a
        // style group is ~8-10 ms, so merging text groups that differ only by fill (and
        // style groups that differ only by typography) removes the expensive groups without
        // skipping or repeating a single property write.
        Dictionary<StyleGroupKey, List<OfficeShapeRequest>> styleGroups = [];
        List<StyleGroupKey> styleOrder = [];
        Dictionary<TypographyGroupKey, List<OfficeShapeRequest>> textGroups = [];
        List<TypographyGroupKey> textOrder = [];
        Dictionary<string, Excel.Shape> resolvedShapes = new(requests.Count, StringComparer.Ordinal);
        List<string> mismatched = [];

        long preflightStarted = Stopwatch.GetTimestamp();
        foreach (OfficeShapeRequest request in requests)
        {
            if (!Validate(request, out ShapeWriteRefusal? invalid))
            {
                return ShapeGroupWriteOutcome.Refused(request.PrimitiveId, invalid!.Value);
            }

            // Resolve and ownership-prove every shape before writing any of it. A missing
            // or user-retagged shape refuses the whole pass rather than styling the rest
            // and then reporting a failure against a half-written set.
            Excel.Shape? shape = FindShapeByName(shapes, NameFor(request.PrimitiveId));
            if (shape is null || !CarriesOwnershipTagFor(shape, request.PrimitiveId))
            {
                return ShapeGroupWriteOutcome.Refused(request.PrimitiveId, ShapeWriteRefusal.NotFound);
            }

            resolvedShapes[request.PrimitiveId] = shape;

            // Geometry and content are verified here too, so a fully matching shape never
            // reaches the per-shape dispatch. The reads are the same ones UpdateContent
            // would perform; doing them once, in this pass, removes a second tag read and
            // a second dispatch per shape. A mismatch is named and the reconciler sends
            // only that shape down the write path.
            bool geometryDiffers = !MatchesGeometry(shape, request);
            bool contentDiffers = request.Kind == OfficeShapeKind.TextBox
                && !TextMatches(shape, request.Text);
            if (geometryDiffers || contentDiffers)
            {
                mismatched.Add(request.PrimitiveId);
            }

            StyleGroupKey styleKey = StyleGroupKey.For(request);
            if (!styleGroups.TryGetValue(styleKey, out List<OfficeShapeRequest>? styleGroup))
            {
                styleGroup = [];
                styleGroups.Add(styleKey, styleGroup);
                styleOrder.Add(styleKey);
            }

            styleGroup.Add(request);

            if (request.Kind == OfficeShapeKind.TextBox)
            {
                TypographyGroupKey textKey = TypographyGroupKey.For(request);
                if (!textGroups.TryGetValue(textKey, out List<OfficeShapeRequest>? textGroup))
                {
                    textGroup = [];
                    textGroups.Add(textKey, textGroup);
                    textOrder.Add(textKey);
                }

                textGroup.Add(request);
            }
        }

        double preflightMilliseconds = Stopwatch.GetElapsedTime(preflightStarted).TotalMilliseconds;

        long styleStarted = Stopwatch.GetTimestamp();
        foreach (StyleGroupKey key in styleOrder)
        {
            // Read-back verification: a converged group is read, not written.
            if (!AnyStyleDiffers(styleGroups[key], resolvedShapes))
            {
                continue;
            }

            if (!WriteGroup(styleGroups[key], ApplyStyle, ApplyStyleToRange, shapes, resolvedShapes, out ShapeGroupWriteOutcome? refused))
            {
                return refused!;
            }
        }

        double styleMilliseconds = Stopwatch.GetElapsedTime(styleStarted).TotalMilliseconds;

        long textStarted = Stopwatch.GetTimestamp();
        foreach (TypographyGroupKey key in textOrder)
        {
            if (!AnyTextFormattingDiffers(textGroups[key], resolvedShapes))
            {
                continue;
            }

            if (!WriteGroup(textGroups[key], ApplyTextFormatting, ApplyTextFormattingToRange, shapes, resolvedShapes, out ShapeGroupWriteOutcome? textRefused))
            {
                return textRefused!;
            }
        }

        double textMilliseconds = Stopwatch.GetElapsedTime(textStarted).TotalMilliseconds;

        if (_trace is not null)
        {
            int largestStyle = 0;
            foreach (StyleGroupKey key in styleOrder)
            {
                largestStyle = Math.Max(largestStyle, styleGroups[key].Count);
            }

            int largestText = 0;
            foreach (TypographyGroupKey key in textOrder)
            {
                largestText = Math.Max(largestText, textGroups[key].Count);
            }

            _trace(string.Create(
                CultureInfo.InvariantCulture,
                $"gantt.shape group-appearance shapes={requests.Count} stylegroups={styleOrder.Count} textgroups={textOrder.Count} maxstyle={largestStyle} maxtext={largestText} preflightms={preflightMilliseconds:0.0} stylems={styleMilliseconds:0.0} textms={textMilliseconds:0.0} ms={Stopwatch.GetElapsedTime(groupStarted).TotalMilliseconds:0.0}"));
        }

        return ShapeGroupWriteOutcome.Ok(mismatched);
    }

    /// <summary>Writes one same-appearance bucket through one shape or one range.</summary>
    /// <param name="group">The requests sharing this bucket's written members.</param>
    /// <param name="single">The single-shape writer for this family.</param>
    /// <param name="range">The range writer for this family.</param>
    /// <param name="shapes">The resolved shapes collection.</param>
    /// <param name="resolvedShapes">The ownership-proven proxies, keyed by primitive id.</param>
    /// <param name="refused">The typed refusal, or null on success.</param>
    /// <returns><see langword="true"/> when the bucket was written.</returns>
    private bool WriteGroup(
        List<OfficeShapeRequest> group,
        Action<Excel.Shape, OfficeShapeRequest> single,
        Action<Excel.ShapeRange, OfficeShapeRequest> range,
        Excel.Shapes shapes,
        Dictionary<string, Excel.Shape> resolvedShapes,
        out ShapeGroupWriteOutcome? refused)
    {
        OfficeShapeRequest first = group[0];
        refused = null;
#pragma warning disable CA1031 // A host refusal is a typed group result, not an escaped exception.
        try
        {
            if (group.Count == 1)
            {
                single(resolvedShapes[first.PrimitiveId], first);
                return true;
            }

            string[] names = new string[group.Count];
            for (int index = 0; index < group.Count; index++)
            {
                names[index] = NameFor(group[index].PrimitiveId);
            }

            range(GetShapeRange(shapes, names), first);
            return true;
        }
        catch (ArgumentOutOfRangeException)
        {
            // A scene value with no host mapping is a contract fault, not a host failure;
            // it must not be swallowed into a group refusal.
            throw;
        }
        catch (Exception)
        {
            refused = ShapeGroupWriteOutcome.Refused(first.PrimitiveId, ShapeWriteRefusal.HostRejected);
            return false;
        }
#pragma warning restore CA1031
    }

    /// <summary>Whether any shape in a style bucket no longer matches its request.</summary>
    /// <param name="group">The bucket's requests.</param>
    /// <param name="resolvedShapes">The ownership-proven proxies, keyed by primitive id.</param>
    /// <returns><see langword="true"/> when the bucket needs the block write.</returns>
    private static bool AnyStyleDiffers(
        List<OfficeShapeRequest> group,
        Dictionary<string, Excel.Shape> resolvedShapes)
    {
        foreach (OfficeShapeRequest request in group)
        {
            Excel.Shape shape = resolvedShapes[request.PrimitiveId];
            if (!StyleMatches(shape.Fill, shape.Line, OfficeStyleMapper.Map(request)))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>Whether any text box in a typography bucket no longer matches its request.</summary>
    /// <param name="group">The bucket's requests.</param>
    /// <param name="resolvedShapes">The ownership-proven proxies, keyed by primitive id.</param>
    /// <returns><see langword="true"/> when the bucket needs the block write.</returns>
    private bool AnyTextFormattingDiffers(
        List<OfficeShapeRequest> group,
        Dictionary<string, Excel.Shape> resolvedShapes)
    {
        foreach (OfficeShapeRequest request in group)
        {
            if (!TextFormattingMatches(resolvedShapes[request.PrimitiveId], request))
            {
                return true;
            }
        }

        return false;
    }

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
    /// <param name="applyContent">
    /// Whether to write a text box's string without repeating its formatting. The grouped
    /// update uses this after <see cref="ApplyGroupAppearance"/>; a full update implies it.
    /// </param>
    /// <returns>The typed result.</returns>
    private ShapeWriteOutcome UpdateCore(OfficeShapeRequest request, bool applyAppearance, bool applyContent = false)
    {
        ArgumentNullException.ThrowIfNull(request);

        if (!Validate(request, out ShapeWriteRefusal? invalid))
        {
            return ShapeWriteOutcome.Refused(invalid!.Value);
        }

        // ADR-0008 D4's first statement, before any workbook access. ADR-0054: inside a
        // reconcile batch the target was already checked when it was resolved and events are
        // off for the whole pass, so the repeated query can only return the same answer.
        if (!_batch.Active && RefusalFor(_protectionGuard.Query()) is { } guardRefusal)
        {
            return guardRefusal;
        }

        Excel.Shapes? shapes = ResolveShapes(out ShapeWriteOutcome? refusal);
        if (shapes is null)
        {
            return refusal!;
        }

        Excel.Shape? existing = FindShapeByName(shapes, NameFor(request.PrimitiveId));
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
        if (!ApplyGeometry(existing, request.Kind, request.Geometry, request.HostOwnsExtent))
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

        TraceBandShape(
            applyAppearance ? "update" : (applyContent ? "content" : "move"),
            existing,
            request.PrimitiveId,
            bandBefore);

        // The full update writes the whole shape: content and style. It is the path a
        // caller takes when it has no grouped appearance pass (the tests, and any future
        // single-shape caller).
        if (applyAppearance)
        {
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

        // The grouped path: the appearance and typography were written once for the
        // group, so only the string is left for a text box. The move-only path stops
        // at the geometry.
        return (!applyContent || request.Kind != OfficeShapeKind.TextBox)
            ? ShapeWriteOutcome.Ok()
            : TryApplyTextContent(existing, request)
                ? ShapeWriteOutcome.Ok()
                : ShapeWriteOutcome.Refused(ShapeWriteRefusal.HostRejected);
    }

}
