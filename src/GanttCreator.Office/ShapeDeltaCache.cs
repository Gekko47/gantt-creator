using System.Security.Cryptography;
using System.Text.Json;

namespace GanttCreator.Office;

/// <summary>In-memory signatures for the last fully reconciled shape payloads.</summary>
/// <remarks>Stores identifiers, hashes, and z-order only; rendered/user text is never retained.</remarks>
public sealed class ShapeDeltaCache
{
    private IReadOnlyDictionary<string, string>? _signatures;
    private IReadOnlyDictionary<string, ShapePartSignatures>? _parts;
    private IReadOnlyList<string>? _backToFront;

    /// <summary>Gets the last successful snapshot, or null after invalidation.</summary>
    public ShapeDeltaSnapshot? Snapshot => _signatures is null || _backToFront is null
        ? null
        : new ShapeDeltaSnapshot(_signatures, _backToFront) { Parts = _parts ?? new Dictionary<string, ShapePartSignatures>() };

    /// <summary>Commits desired payload signatures after complete host success.</summary>
    public void Commit(IReadOnlyList<OfficeShapeRequest> desired)
    {
        ArgumentNullException.ThrowIfNull(desired);
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
}

/// <summary>One shape's kind and part hashes from the last successful apply.</summary>
/// <param name="Kind">The shape kind last applied; a change forces the full update path.</param>
/// <param name="Geometry">The hash of the shape's box.</param>
/// <param name="Appearance">The hash of its fill, stroke, hatch, typography, text and alignment.</param>
public sealed record ShapePartSignatures(string Kind, string Geometry, string Appearance);
