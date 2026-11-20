using System.Security.Cryptography;
using System.Text.Json;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office;

/// <summary>In-memory signatures for the last fully reconciled shape payloads.</summary>
/// <remarks>Stores identifiers, hashes, and z-order only; rendered/user text is never retained.</remarks>
public sealed class ShapeDeltaCache
{
    private IReadOnlyDictionary<string, string>? _signatures;
    private IReadOnlyDictionary<string, ShapePartSignatures>? _parts;
    private IReadOnlyList<string>? _backToFront;
    private string? _layout;

    /// <summary>Gets the last successful snapshot, or null after invalidation.</summary>
    public ShapeDeltaSnapshot? Snapshot => _signatures is null || _backToFront is null
        ? null
        : new ShapeDeltaSnapshot(_signatures, _backToFront)
        {
            Parts = _parts ?? new Dictionary<string, ShapePartSignatures>(),
            Layout = _layout ?? string.Empty,
        };

    /// <summary>Commits desired payload signatures after complete host success.</summary>
    /// <param name="desired">The desired shape set the host accepted.</param>
    /// <param name="layout">
    /// The measured worksheet layout the shapes were placed against, or null when the caller does
    /// not compute one (in which case the next pass verifies geometry against the host, as before).
    /// </param>
    public void Commit(IReadOnlyList<OfficeShapeRequest> desired, string? layout = null)
    {
        ArgumentNullException.ThrowIfNull(desired);
        _layout = layout;
        Dictionary<string, string> signatures = new(StringComparer.Ordinal);
        Dictionary<string, ShapePartSignatures> parts = new(StringComparer.Ordinal);
        foreach (OfficeShapeRequest request in desired)
        {
            signatures.Add(request.PrimitiveId, Signature(request));
            parts.Add(request.PrimitiveId, new ShapePartSignatures(
                request.Kind.ToString(),
                GeometrySignature(request),
                AppearanceSignature(request)));
        }

        _signatures = signatures;
        _parts = parts;
        _backToFront = [.. desired.Select(request => request.PrimitiveId)];
    }

    internal static string Signature(OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        return Hash(request);
    }

    /// <summary>
    /// Hashes the part of a request that decides WHERE the shape is: its box.
    /// </summary>
    /// <param name="request">The shape request.</param>
    /// <returns>The geometry part's hash.</returns>
    /// <remarks>
    /// The request is split into parts so a reconciliation can tell a shape that MOVED from a
    /// shape that looks different. Moving one is a handful of property writes; a full update
    /// re-applies fill, stroke, hatch, font, text, alignment and colour as well, and measured
    /// ~15–18 ms per primitive against ~0.5 ms for a move (ADR-0053).
    /// <para>
    /// The KIND is deliberately not in this hash. It is stored beside it, because a shape whose
    /// kind changed (a bar becoming a diamond) must never be moved as though it were the same
    /// shape; the reconciler requires kind equality before it plans a move.
    /// </para>
    /// </remarks>
    internal static string GeometrySignature(OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        return Hash(request.Geometry);
    }

    /// <summary>
    /// Hashes the part of a request that decides how the shape LOOKS and what it SAYS:
    /// fill, stroke, hatch, typography, text, alignment and text colour.
    /// </summary>
    /// <param name="request">The shape request.</param>
    /// <returns>The appearance part's hash.</returns>
    /// <remarks>
    /// <see cref="OfficeShapeRequest.ZLayer"/> is deliberately in neither part. It orders
    /// the plan and the single z-order pass; it is not a property of the shape, so a change
    /// to it alone must not be mistaken for a move or a restyle.
    /// </remarks>
    internal static string AppearanceSignature(OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        return Hash(new
        {
            request.FillColour,
            request.StrokeColour,
            request.LineWidthPt,
            request.HatchPattern,
            request.StrokeDash,
            request.FontFamily,
            request.FontSizePt,
            request.Bold,
            request.Text,
            request.Alignment,
            request.TextColour,
        });
    }

    private static string Hash<T>(T payload) =>
        Convert.ToHexString(SHA256.HashData(JsonSerializer.SerializeToUtf8Bytes(payload)));

    /// <summary>Invalidates the snapshot after a possibly partial host mutation.</summary>
    public void Invalidate()
    {
        _signatures = null;
        _parts = null;
        _backToFront = null;
        _layout = null;
    }

    /// <summary>
    /// Fingerprints the measured worksheet layout the chart was placed against.
    /// </summary>
    /// <param name="grid">The measured panel grid.</param>
    /// <param name="rowHeightsWritten">How many body row heights this pass wrote.</param>
    /// <returns>A stable string that changes whenever the placement inputs change.</returns>
    /// <remarks>
    /// <para>
    /// <b>What it proves.</b> Excel moves an <c>xlMove</c> shape only when the rows or columns it
    /// is anchored to change: an inserted or deleted row, a hidden row, a row height, a column
    /// width, or the table's origin. Every one of those is an input to this fingerprint, and the
    /// measurement it is built from is the same value the scene was laid out with. Two passes
    /// with an identical fingerprint therefore describe an identical set of anchors, so a shape
    /// that still carries the payload the last pass wrote is still where the last pass put it —
    /// unless someone moved it by hand, which the owner has ruled the automatic pass need not
    /// repair (the explicit Refresh does, and it has no snapshot so it always verifies).
    /// </para>
    /// <para>
    /// <b>Why the written-heights count is in it.</b> A pass that normalised a height changed the
    /// sheet after measuring, so its own measurement cannot be the fingerprint of the state it
    /// left; including the count forces the next pass to verify rather than trust.
    /// </para>
    /// </remarks>
    internal static string LayoutFingerprint(PanelCellGrid grid, int rowHeightsWritten)
    {
        ArgumentNullException.ThrowIfNull(grid);
        System.Text.StringBuilder builder = new();
        _ = builder.Append(rowHeightsWritten.ToString(System.Globalization.CultureInfo.InvariantCulture));
        _ = builder.Append('|').Append(grid.RowHeightsPt.Count.ToString(System.Globalization.CultureInfo.InvariantCulture));
        foreach (double height in grid.RowHeightsPt)
        {
            _ = builder.Append('|').Append(height.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture));
        }

        _ = builder.Append('|').Append(grid.Columns.Count.ToString(System.Globalization.CultureInfo.InvariantCulture));
        foreach (PanelColumn column in grid.Columns)
        {
            _ = builder.Append('|').Append(column.WidthPt.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture));
        }

        _ = builder.Append('|').Append(grid.OriginTopPt.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture));
        _ = builder.Append('|').Append(grid.OriginLeftPt.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture));
        _ = builder.Append('|').Append(grid.HeaderHeightPt.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture));
        _ = builder.Append('|').Append(grid.TopPaddingHeightPt.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture));
        _ = builder.Append('|').Append(grid.BottomPaddingHeightPt.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture));
        return builder.ToString();
    }
}

/// <summary>One successful desired-shape signature snapshot.</summary>
/// <param name="Signatures">Stable shape IDs mapped to SHA-256 payload hashes.</param>
/// <param name="BackToFront">The desired shape order.</param>
public sealed record ShapeDeltaSnapshot(
    IReadOnlyDictionary<string, string> Signatures,
    IReadOnlyList<string> BackToFront)
{
    /// <summary>
    /// Stable shape IDs mapped to their geometry and appearance part hashes; empty for a
    /// snapshot built without parts, in which case every changed shape gets a full update.
    /// </summary>
    public IReadOnlyDictionary<string, ShapePartSignatures> Parts { get; init; } =
        new Dictionary<string, ShapePartSignatures>();

    /// <summary>
    /// The layout fingerprint the shapes were placed against, or empty when it was not recorded.
    /// </summary>
    /// <remarks>
    /// An empty value never matches a computed fingerprint, so a snapshot written before this
    /// property existed makes the next pass verify as it always did.
    /// </remarks>
    public string Layout { get; init; } = string.Empty;
}

/// <summary>One shape's kind and part hashes from the last successful apply.</summary>
/// <param name="Kind">The shape kind last applied; a change forces the full update path.</param>
/// <param name="Geometry">The hash of the shape's box.</param>
/// <param name="Appearance">The hash of its fill, stroke, hatch, typography, text and alignment.</param>
public sealed record ShapePartSignatures(string Kind, string Geometry, string Appearance);
