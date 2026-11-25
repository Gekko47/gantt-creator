using System.Globalization;
using GanttCreator.Core;
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
/// <param name="ownershipScheme">
/// The ownership scheme to stamp and prove, or <see langword="null"/> for the live
/// chart's scheme. A staging scheme (R6.2 D4) writes the staging name prefix and tag
/// instead, so a staged composition and the live chart can share one worksheet without
/// either ownership scan touching the other's shapes.
/// </param>
/// <param name="reuseTemplates">
/// Whether an open batch may create a later shape by duplicating an earlier shape with
/// the same appearance, or <see langword="false"/> to create every shape normally. The
/// editable-export staging writer opts in; the live writer does not.
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
/// <para>
/// <b>One ownership scheme per instance (R6.2 D4).</b> The adapter can be constructed
/// with <see cref="ShapeOwnershipScheme.Staging"/> instead of the default live scheme.
/// A staging instance stamps the staging name prefix and staging tag on every shape it
/// creates and proves only those on delete and list, so the live chart and a staged
/// composition can share one worksheet without either scan ever touching the other's
/// shapes. The tag proof remains a recomputation from the shape name in both schemes.
/// </para>
/// </remarks>
public partial class ExcelShapeWriter(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null,
    Action<string>? trace = null,
    ShapeOwnershipScheme? ownershipScheme = null,
    bool reuseTemplates = false) : IShapeWritePort, IShapeAppearanceHost
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <summary>
    /// Whether a batch may create a later shape by duplicating an earlier one with the
    /// same appearance (R6.8 follow-up, measured 2026-11-19).
    /// </summary>
    /// <remarks>
    /// <para>
    /// Opt-in, and enabled only for the editable-export staging writer. Measured on the
    /// 20-activity copy: a styled text box costs 34.6 ms to create and 8.6 ms to
    /// duplicate-and-rewrite; a rectangle 24.9 ms against 2.3 ms; a line 14.4 ms against
    /// 2.1 ms. The copy creates 340 shapes, so duplication is the only lever that
    /// removes the style-write block (4.2 s of a 9.9 s lifecycle).
    /// </para>
    /// <para>
    /// The live writer keeps creating every shape normally: a refresh creates only the
    /// handful of shapes the scene gained, and the live path's behaviour stays exactly
    /// as measured before this row.
    /// </para>
    /// </remarks>
    /// <summary>
    /// Gets whether an open batch may create a shape by duplicating an earlier shape with
    /// the same appearance, rather than by calling <c>AddShape</c> and re-applying the style.
    /// </summary>
    /// <remarks>
    /// Opt-in, because the staging and live compositions own their create paths separately. A
    /// caller that builds the writer without the flag silently keeps the old create cost, so
    /// the live factory is pinned by a test rather than read from its constructor call.
    /// </remarks>
    internal bool ReusesTemplates { get; } = reuseTemplates;

    /// <summary>
    /// Returns the appearance collaborator, creating it on first use.
    /// </summary>
    /// <returns>The appearance collaborator bound to this writer.</returns>
    /// <remarks>
    /// A method rather than a lazy property, because the collaborator needs
    /// <c>this</c> and a field initializer cannot reference it. The collaborator
    /// receives this writer as its host, so every style and text write still passes
    /// through the <c>internal virtual</c> seams below and a test subclass's overrides
    /// stay on the path.
    /// </remarks>
    private ShapeAppearanceWriter Appearance() => _appearance ??= new ShapeAppearanceWriter(this);

    /// <summary>
    /// The ownership scheme this adapter stamps and proves: the live chart's scheme by
    /// default, or the staging scheme (R6.2 D4). One instance is one scheme for its
    /// whole life; the two families are distinguished by name prefix and tag, never by
    /// instance identity.
    /// </summary>
    private readonly ShapeOwnershipScheme _ownership = ownershipScheme ?? ShapeOwnershipScheme.Live;

    /// <summary>Optional diagnostic sink for the splitter band's live geometry.</summary>
    private readonly Action<string>? _trace = trace;

    /// <summary>The appearance collaborator, created lazily by <c>Appearance()</c>.</summary>
    private ShapeAppearanceWriter? _appearance;

    /// <summary>
    /// The batch state: the hoisted shapes collection, the proxy and tag memos, the
    /// template cache, the name-index flag, and the per-step create timings (ADR-0054).
    /// </summary>
    /// <remarks>
    /// One instance for the writer's whole life, opened and closed by
    /// <see cref="BeginBatch"/>. Every member is reset on open and cleared on close, so a
    /// read outside a batch sees the same empty state the individual fields used to hold;
    /// the contract that the tag memo is empty outside a batch lives on
    /// <see cref="ShapeWriteBatch.Tags"/>.
    /// </remarks>
    private readonly ShapeWriteBatch _batch = new(trace);

    /// <summary>
    /// The appearance members a template duplicate must share with the request. Geometry,
    /// primitive identifier, z-order, and text content are deliberately absent: each
    /// duplicate rewrites the first three and the text is written per shape.
    /// </summary>
    /// <remarks>
    /// For a line the two direction signs are part of the key. Excel's <c>Duplicate</c>
    /// keeps the source line's drawing direction, and a left/top/width/height write does
    /// not flip it, so a right-to-left source must not serve a left-to-right request.
    /// </remarks>
    internal readonly record struct ShapeTemplateKey(
        OfficeShapeKind Kind,
        bool HostOwnsExtent,
        ColourHex? FillColour,
        ColourHex? StrokeColour,
        double? LineWidthPt,
        GanttHatchPattern HatchPattern,
        GanttStrokeDash StrokeDash,
        string? FontFamily,
        double? FontSizePt,
        bool? Bold,
        GanttTextAlignment? Alignment,
        ColourHex? TextColour,
        double LeftInsetPt,
        int LineDxSign,
        int LineDySign)
    {
        internal static ShapeTemplateKey For(OfficeShapeRequest request)
        {
            int dx = 0;
            int dy = 0;
            if (request.Kind == OfficeShapeKind.Line
                && request.Geometry.From is { } from
                && request.Geometry.To is { } to)
            {
                dx = Math.Sign(to.X - from.X);
                dy = Math.Sign(to.Y - from.Y);
            }

            return new ShapeTemplateKey(
                request.Kind,
                request.HostOwnsExtent,
                request.FillColour,
                request.StrokeColour,
                request.LineWidthPt,
                request.HatchPattern,
                request.StrokeDash,
                request.FontFamily,
                request.FontSizePt,
                request.Bold,
                request.Alignment,
                request.TextColour,
                request.LeftInsetPt,
                dx,
                dy);
        }
    }

    /// <summary>
    /// The members a Fill/Line write depends on. Deliberately excludes every typography
    /// member: a style write cannot change them, so two shapes that differ only in font or
    /// text colour can share one style range.
    /// </summary>
    private readonly record struct StyleGroupKey(
        OfficeShapeKind Kind,
        bool HostOwnsExtent,
        ColourHex? FillColour,
        ColourHex? StrokeColour,
        double? LineWidthPt,
        GanttHatchPattern HatchPattern,
        GanttStrokeDash StrokeDash,
        int LineDxSign,
        int LineDySign)
    {
        internal static StyleGroupKey For(OfficeShapeRequest request)
        {
            int dx = 0;
            int dy = 0;
            if (request.Kind == OfficeShapeKind.Line
                && request.Geometry.From is { } from
                && request.Geometry.To is { } to)
            {
                dx = Math.Sign(to.X - from.X);
                dy = Math.Sign(to.Y - from.Y);
            }

            return new StyleGroupKey(
                request.Kind,
                request.HostOwnsExtent,
                request.FillColour,
                request.StrokeColour,
                request.LineWidthPt,
                request.HatchPattern,
                request.StrokeDash,
                dx,
                dy);
        }
    }

    /// <summary>
    /// The members a <c>TextFrame2</c> write depends on. Deliberately excludes Fill/Line:
    /// a typography write cannot change the shape's fill or stroke, so two text boxes that
    /// differ only in fill can share one typography range.
    /// </summary>
    private readonly record struct TypographyGroupKey(
        string? FontFamily,
        double? FontSizePt,
        bool? Bold,
        GanttTextAlignment? Alignment,
        ColourHex? TextColour,
        double LeftInsetPt)
    {
        internal static TypographyGroupKey For(OfficeShapeRequest request) =>
            new(
                request.FontFamily,
                request.FontSizePt,
                request.Bold,
                request.Alignment,
                request.TextColour,
                request.LeftInsetPt);
    }


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
        if (_batch.Active)
        {
            throw new InvalidOperationException(
                "A shape-write batch is already open. Batches do not nest, because the inner "
                + "scope would release the resolution the outer one is still using.");
        }

        _batch.Open(ReusesTemplates);
        return new BatchScope(this);
    }

    /// <summary>Ends a batch and releases the state it was holding.</summary>
    private sealed class BatchScope(ExcelShapeWriter owner) : IDisposable
    {
        /// <inheritdoc />
        /// <remarks>
        /// Closing reports the create-step record before the state is cleared, and only when
        /// the caller asked for trace records: a measurement of the pass, not a product
        /// behaviour.
        /// </remarks>
        public void Dispose() => owner._batch.Close();
    }

    /// <summary>
    /// Builds the open batch's complete name index exactly once, so an ABSENT name can be
    /// answered without asking the host.
    /// </summary>
    /// <param name="shapes">The batch's resolved shapes collection.</param>
    /// <remarks>
    /// The scan and its measured rationale live on <see cref="ShapeWriteBatch.EnsureNameIndex"/>;
    /// this method exists only to hand that class the two shape seams, so a test subclass
    /// that overrides <see cref="GetShapeCount"/> or <see cref="GetShapeAt"/> stays on the
    /// path.
    /// </remarks>
    private void EnsureBatchNameIndex(Excel.Shapes shapes) =>
        _batch.EnsureNameIndex(shapes, GetShapeCount, GetShapeAt);

    /// <summary>Remembers a resolved proxy for the rest of the batch, when one is open.</summary>
    /// <param name="name">The exact shape name.</param>
    /// <param name="shape">The resolved shape, or <see langword="null"/>.</param>
    /// <returns>The shape, so a lookup can return it directly.</returns>
    private Excel.Shape? Remember(string name, Excel.Shape? shape) => _batch.Remember(name, shape);

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

    /// <summary>
    /// Records a refused create with the primitive it concerned and the typed reason, then
    /// returns the refusal unchanged.
    /// </summary>
    /// <param name="request">The request that was refused.</param>
    /// <param name="refusal">The typed reason the write refused.</param>
    /// <param name="step">
    /// The step that refused: <c>validate</c>, <c>name-length</c>, <c>guard</c>,
    /// <c>resolve</c>, <c>already-exists</c>, <c>create</c>, <c>style</c>, or
    /// <c>ownership</c>.
    /// </param>
    /// <returns>The refused outcome, unchanged.</returns>
    /// <remarks>
    /// The 2026-11-19 live defect drove this record. The editable copy refused
    /// deterministically at one primitive of a full export composition: 24 creates
    /// succeeded, the next refused, and the editable composer collapses every write
    /// refusal to <see cref="StagingCompositionRefusal.HostRejected"/>, so the log held
    /// no record of <em>which</em> primitive or <em>why</em>. One line per refusal makes
    /// the next occurrence self-diagnosing without changing the typed outcome any caller
    /// already handles.
    /// </remarks>
    private ShapeWriteOutcome RefuseCreate(OfficeShapeRequest request, ShapeWriteRefusal refusal, string step)
    {
        _trace?.Invoke(string.Create(
            CultureInfo.InvariantCulture,
            $"gantt.shape create-refused id={request.PrimitiveId} kind={request.Kind} refusal={refusal} step={step} {DescribeGeometry(request.Geometry)}"));
        return ShapeWriteOutcome.Refused(refusal);
    }

    /// <summary>Renders a request's geometry into one bounded diagnostic field.</summary>
    /// <param name="geometry">The request's geometry.</param>
    /// <returns>A single-token description of the bounds or the line endpoints.</returns>
    /// <remarks>
    /// A zero extent is legal in <see cref="RectD"/> but refused by
    /// <see cref="IsUsableBounds"/>, so the numbers themselves are the evidence for that
    /// refusal; a summary such as "invalid" would hide the defect that produced it.
    /// </remarks>
    private static string DescribeGeometry(OfficeShapeGeometry geometry) => geometry switch
    {
        { Bounds: { } bounds } => string.Create(
            CultureInfo.InvariantCulture,
            $"b={bounds.X:0.###},{bounds.Y:0.###},{bounds.Width:0.###},{bounds.Height:0.###}"),
        { From: { } from, To: { } to } => string.Create(
            CultureInfo.InvariantCulture,
            $"from={from.X:0.###},{from.Y:0.###} to={to.X:0.###},{to.Y:0.###}"),
        _ => "b=none",
    };

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

}
