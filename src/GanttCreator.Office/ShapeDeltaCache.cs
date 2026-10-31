using System.Security.Cryptography;
using System.Text.Json;

namespace GanttCreator.Office;

/// <summary>In-memory signatures for the last fully reconciled shape payloads.</summary>
/// <remarks>Stores identifiers, hashes, and z-order only; rendered/user text is never retained.</remarks>
public sealed class ShapeDeltaCache
{
    private IReadOnlyDictionary<string, string>? _signatures;
    private IReadOnlyList<string>? _backToFront;

    /// <summary>Gets the last successful snapshot, or null after invalidation.</summary>
    public ShapeDeltaSnapshot? Snapshot => _signatures is null || _backToFront is null
        ? null
        : new ShapeDeltaSnapshot(_signatures, _backToFront);

    /// <summary>Commits desired payload signatures after complete host success.</summary>
    public void Commit(IReadOnlyList<OfficeShapeRequest> desired)
    {
        ArgumentNullException.ThrowIfNull(desired);
        Dictionary<string, string> signatures = new(StringComparer.Ordinal);
        foreach (OfficeShapeRequest request in desired)
        {
            signatures.Add(request.PrimitiveId, Signature(request));
        }

        _signatures = signatures;
        _backToFront = [.. desired.Select(request => request.PrimitiveId)];
    }

    internal static string Signature(OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        byte[] payload = JsonSerializer.SerializeToUtf8Bytes(request);
        return Convert.ToHexString(SHA256.HashData(payload));
    }

    /// <summary>Invalidates the snapshot after a possibly partial host mutation.</summary>
    public void Invalidate()
    {
        _signatures = null;
        _backToFront = null;
    }
}

/// <summary>One successful desired-shape signature snapshot.</summary>
/// <param name="Signatures">Stable shape IDs mapped to SHA-256 payload hashes.</param>
/// <param name="BackToFront">The desired shape order.</param>
public sealed record ShapeDeltaSnapshot(
    IReadOnlyDictionary<string, string> Signatures,
    IReadOnlyList<string> BackToFront);
